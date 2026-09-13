"""Cross-row offline pre-screening rules."""

from __future__ import annotations

from typing import Any, Sequence

from .config import FIELD_MAP
from .normalization import normalize_company_name, normalize_credit_code, normalize_phone
from .rules import AuditIssue, StudentAuditResult


def _company_identity(row: dict[str, Any]) -> tuple[str, str]:
    credit_code = normalize_credit_code(row.get(FIELD_MAP["credit_code"]))
    if credit_code:
        return ("credit", credit_code)
    return ("name", normalize_company_name(row.get(FIELD_MAP["company_name"])))


def apply_batch_contact_phone_duplicates(
    results: Sequence[StudentAuditResult], rows: Sequence[dict[str, Any]]
) -> None:
    """Flag one contact number reused across distinct employer identities in a batch."""
    by_phone: dict[str, list[tuple[int, tuple[str, str]]]] = {}
    for index, row in enumerate(rows):
        identity = _company_identity(row)
        phones = {
            normalize_phone(row.get(FIELD_MAP[field_key]))
            for field_key in ("company_phone", "company_mobile")
        }
        for phone in phones - {""}:
            by_phone.setdefault(phone, []).append((index, identity))

    for entries in by_phone.values():
        if len({identity for _, identity in entries}) < 2:
            continue
        for index, _ in entries:
            issue = AuditIssue(
                rule="batch_company_contact_phone_duplicate",
                field="company_contact_phone",
                message="本批次不同用人单位使用了相同联系人电话，请人工核实",
                student_value="",
                reference_value="",
                highlight_fields=(FIELD_MAP["company_phone"], FIELD_MAP["company_mobile"]),
            )
            if issue not in results[index].issues:
                results[index].add_issue(issue)
