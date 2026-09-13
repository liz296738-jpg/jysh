"""Non-persistent GUI review sessions that reuse the tested audit components."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .audit_workbook import audit_workbook
from .batch_rules import apply_batch_contact_phone_duplicates
from .config import FIELD_MAP
from .excel_parser import ExcelParser
from .reference import ReferenceLibrary
from .rules import AuditEngine, ReviewStage, StudentAuditResult


ISSUE_CATEGORIES = {
    "企业信息": {
        "credit_code_history_match",
        "credit_code_reference_conflict",
        "company_name_match",
        "company_name_no_credit_history_match",
        "company_name_reference_conflict",
    },
    "电话": {
        "company_contact_not_student_phone",
        "company_contact_not_family_phone",
        "batch_company_contact_phone_duplicate",
    },
    "单位行业": {"industry_match"},
    "单位性质": {"company_type_match"},
    "工作职位": {"job_category_not_other"},
}


@dataclass(frozen=True)
class ReviewSession:
    input_path: Path
    review_stage: ReviewStage
    results: tuple[StudentAuditResult, ...]

    @property
    def total_count(self) -> int:
        return len(self.results)

    @property
    def normal_count(self) -> int:
        return sum(result.is_ok for result in self.results)

    @property
    def issue_count(self) -> int:
        return self.total_count - self.normal_count

    def category_count(self, category: str) -> int:
        rules = ISSUE_CATEGORIES[category]
        return sum(any(issue.rule in rules for issue in result.issues) for result in self.results)


def filter_results(
    results: tuple[StudentAuditResult, ...], category: str
) -> tuple[StudentAuditResult, ...]:
    if category == "全部异常":
        return tuple(result for result in results if result.issues)
    return tuple(
        result
        for result in results
        if any(issue.rule in ISSUE_CATEGORIES[category] for issue in result.issues)
    )


class ReviewController:
    """Runs one in-memory review at a time; it never stores task history."""

    def __init__(self, reference_path: str | Path) -> None:
        self._reference_path = Path(reference_path)

    def audit(self, input_path: str | Path, review_stage: str | ReviewStage) -> ReviewSession:
        stage = review_stage if isinstance(review_stage, ReviewStage) else ReviewStage(review_stage.upper())
        parsed = ExcelParser(FIELD_MAP).read(Path(input_path))
        engine = AuditEngine(ReferenceLibrary.from_workbook(self._reference_path))
        results = [
            engine.audit(row, index + 2, stage)
            for index, row in enumerate(parsed.rows)
        ]
        apply_batch_contact_phone_duplicates(results, parsed.rows)
        return ReviewSession(Path(input_path), stage, tuple(results))

    def export(self, session: ReviewSession) -> Path:
        return audit_workbook(
            session.input_path,
            session.review_stage,
            self._reference_path,
        ).output_path
