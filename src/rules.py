"""Deterministic, offline row-level employment pre-screening rules."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from .config import FIELD_MAP
from .normalization import (
    normalize_company_name,
    normalize_credit_code,
    normalize_phone,
    normalize_text,
)
from .reference import ReferenceLibrary, ReferenceRecord


class ReviewStage(str, Enum):
    INITIAL = "INITIAL"
    FINAL = "FINAL"


@dataclass(frozen=True)
class AuditIssue:
    rule: str
    field: str
    message: str
    student_value: str
    reference_value: str
    highlight_fields: tuple[str, ...]


@dataclass
class StudentAuditResult:
    row_number: int
    student_name: str
    student_id: str
    company_name: str
    review_stage: ReviewStage
    is_ok: bool = True
    issues: list[AuditIssue] = field(default_factory=list)

    def add_issue(self, issue: AuditIssue) -> None:
        self.issues.append(issue)
        self.is_ok = False


def _raw(row: dict[str, Any], field_key: str) -> str:
    value = row.get(FIELD_MAP[field_key], "")
    return str(value or "").strip()


def _issue(
    rule: str,
    field: str,
    message: str,
    student_value: str = "",
    reference_value: str = "",
    highlight_fields: tuple[str, ...] = (),
) -> AuditIssue:
    return AuditIssue(rule, field, message, student_value, reference_value, highlight_fields)


class AuditEngine:
    """Apply the shared initial/final-stage rules to one exported student row."""

    def __init__(self, reference_library: ReferenceLibrary) -> None:
        self._reference_library = reference_library

    def audit(
        self,
        row: dict[str, Any],
        row_number: int,
        review_stage: ReviewStage,
    ) -> StudentAuditResult:
        company_name = _raw(row, "company_name")
        result = StudentAuditResult(
            row_number=row_number,
            student_name=_raw(row, "student_name"),
            student_id=_raw(row, "student_id"),
            company_name=company_name,
            review_stage=review_stage,
        )
        reference = self._check_company_identity(row, result)
        if reference is not None:
            self._check_reference_field(
                result,
                row,
                reference,
                "industry",
                "industry_code",
                "industry_match",
                "单位行业与历史参考库不一致",
            )
            self._check_reference_field(
                result,
                row,
                reference,
                "company_type",
                "company_type_code",
                "company_type_match",
                "单位性质与历史参考库不一致",
            )
        self._check_job_category(row, result)
        self._check_student_contact_matches(row, result)
        return result

    def _check_company_identity(
        self, row: dict[str, Any], result: StudentAuditResult
    ) -> ReferenceRecord | None:
        credit_code = _raw(row, "credit_code")
        company_name = _raw(row, "company_name")
        reference = self._reference_library.find(credit_code, company_name)
        if credit_code:
            if reference is None:
                result.add_issue(
                    _issue(
                        "credit_code_history_match",
                        FIELD_MAP["credit_code"],
                        "统一社会信用代码未命中历史参考库，请人工核实",
                        credit_code,
                        "",
                        (FIELD_MAP["credit_code"],),
                    )
                )
                return None
            if reference.reference_conflict:
                result.add_issue(
                    _issue(
                        "credit_code_reference_conflict",
                        FIELD_MAP["credit_code"],
                        "历史参考库中该统一社会信用代码存在冲突记录，请人工核实",
                        credit_code,
                        "",
                        (FIELD_MAP["credit_code"], FIELD_MAP["company_name"]),
                    )
                )
                return None
            if normalize_company_name(company_name) != normalize_company_name(reference.company_name):
                result.add_issue(
                    _issue(
                        "company_name_match",
                        FIELD_MAP["company_name"],
                        "统一社会信用代码对应企业名称与学生填写企业名称不一致",
                        company_name,
                        reference.company_name,
                        (FIELD_MAP["credit_code"], FIELD_MAP["company_name"]),
                    )
                )
            return reference
        if reference is None:
            result.add_issue(
                _issue(
                    "company_name_no_credit_history_match",
                    FIELD_MAP["company_name"],
                    "企业无统一社会信用代码且历史参考库无记录，请人工核实",
                    company_name,
                    "",
                    (FIELD_MAP["company_name"],),
                )
            )
            return None
        if reference.reference_conflict:
            result.add_issue(
                _issue(
                    "company_name_reference_conflict",
                    FIELD_MAP["company_name"],
                    "历史参考库中该企业名称存在冲突记录，请人工核实",
                    company_name,
                    "",
                    (FIELD_MAP["company_name"],),
                )
            )
            return None
        return reference

    def _check_reference_field(
        self,
        result: StudentAuditResult,
        row: dict[str, Any],
        reference: ReferenceRecord,
        name_key: str,
        code_key: str,
        rule: str,
        message: str,
    ) -> None:
        student_code = _raw(row, code_key)
        reference_code = getattr(reference, code_key)
        if student_code and reference_code:
            equal = normalize_text(student_code) == normalize_text(reference_code)
            student_value, reference_value = student_code, reference_code
            field = FIELD_MAP[code_key]
        else:
            student_value = _raw(row, name_key)
            reference_value = getattr(reference, name_key)
            equal = normalize_text(student_value) == normalize_text(reference_value)
            field = FIELD_MAP[name_key]
        if not equal:
            result.add_issue(
                _issue(
                    rule,
                    field,
                    message,
                    student_value,
                    reference_value,
                    (FIELD_MAP[name_key], FIELD_MAP[code_key]),
                )
            )

    def _check_job_category(self, row: dict[str, Any], result: StudentAuditResult) -> None:
        category = _raw(row, "job_category")
        if normalize_text(category) == "其他人员":
            result.add_issue(
                _issue(
                    "job_category_not_other",
                    FIELD_MAP["job_category"],
                    "工作职位类别不能填写为其他人员",
                    category,
                    "",
                    (FIELD_MAP["job_category"],),
                )
            )

    def _check_student_contact_matches(
        self, row: dict[str, Any], result: StudentAuditResult) -> None:
        contacts = [
            _raw(row, "company_phone"),
            _raw(row, "company_mobile"),
        ]
        normalized_contacts = {normalize_phone(phone) for phone in contacts if normalize_phone(phone)}
        for field_key, rule, message in (
            ("student_phone", "company_contact_not_student_phone", "用人单位联系人电话与学生本人联系电话相同"),
            ("family_phone", "company_contact_not_family_phone", "用人单位联系人电话与学生家庭联系电话相同"),
        ):
            student_phone = _raw(row, field_key)
            normalized_student_phone = normalize_phone(student_phone)
            if normalized_student_phone and normalized_student_phone in normalized_contacts:
                result.add_issue(
                    _issue(
                        rule,
                        FIELD_MAP[field_key],
                        message,
                        student_phone,
                        ", ".join(contacts),
                        (FIELD_MAP[field_key], FIELD_MAP["company_phone"], FIELD_MAP["company_mobile"]),
                    )
                )
