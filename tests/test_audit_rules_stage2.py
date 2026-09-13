from __future__ import annotations

from functools import cache
from dataclasses import replace
from pathlib import Path

from src.batch_rules import apply_batch_contact_phone_duplicates
from src.config import FIELD_MAP
from src.excel_parser import ExcelParser
from src.normalization import normalize_company_name, normalize_credit_code, normalize_phone
from src.reference import ReferenceLibrary, build_reference_records
from src.rules import AuditEngine, ReviewStage


PROJECT_ROOT = Path(__file__).resolve().parents[1]


@cache
def _trusted_reference():
    """Derive an anonymous test baseline from a completed historical enterprise."""
    history = ExcelParser(FIELD_MAP).read(PROJECT_ROOT / "samples" / "历史参考总库.xlsx")
    return next(
        record
        for record in build_reference_records(history.rows)
        if not record.reference_conflict
        and record.credit_code
        and record.company_name
        and record.company_type_code
        and record.industry_code
    )


def _student_row(reference, **changes):
    row = {column: "" for column in FIELD_MAP.values()}
    row.update(
        {
            "xm": "匿名学生",
            "xh": "ANON-001",
            "dwmc": reference.company_name,
            "dwzzjgdm": reference.credit_code,
            "dwxz": reference.company_type,
            "dwxzdm": reference.company_type_code,
            "dwhy": reference.industry,
            "dwhydm": reference.industry_code,
            "gzzwlb": "工程技术人员",
        }
    )
    row.update(changes)
    return row


def _audit(reference, row, row_number=2):
    return AuditEngine(ReferenceLibrary([reference])).audit(row, row_number, ReviewStage.INITIAL)


def test_normalization_removes_only_safe_formatting_variants():
    assert normalize_credit_code(" 91-3100 00_abcd ") == "91310000ABCD"
    assert normalize_company_name(" 某公司（中国）  有限公司 ") == "某公司(中国) 有限公司"
    assert normalize_phone(" +86 138-0013-8000 ") == "13800138000"
    assert normalize_phone("010-1234 5678") == "01012345678"


def test_matched_trusted_company_with_equal_codes_is_ok():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference))

    assert result.is_ok is True
    assert result.review_stage is ReviewStage.INITIAL


def test_matching_credit_code_with_different_company_name_is_an_issue():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, dwmc="匿名错误企业"))

    assert result.is_ok is False
    assert result.issues[0].message == "统一社会信用代码对应企业名称与学生填写企业名称不一致"
    assert result.issues[0].student_value == "匿名错误企业"
    assert result.issues[0].reference_value == reference.company_name


def test_unknown_credit_code_requires_manual_review():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, dwzzjgdm="999999999999999999"))

    assert [issue.message for issue in result.issues] == ["统一社会信用代码未命中历史参考库，请人工核实"]


def test_conflicting_reference_credit_code_requires_manual_review():
    reference = replace(_trusted_reference(), reference_conflict=True, conflict_fields=("company_name",))

    result = _audit(reference, _student_row(reference))

    assert [issue.message for issue in result.issues] == ["历史参考库中该统一社会信用代码存在冲突记录，请人工核实"]


def test_blank_credit_code_uses_only_the_no_code_company_name_index():
    reference = replace(_trusted_reference(), credit_code="")

    result = _audit(reference, _student_row(reference, dwzzjgdm=""))

    assert result.is_ok is True


def test_blank_credit_code_without_history_record_requires_manual_review():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, dwzzjgdm=""))

    assert [issue.message for issue in result.issues] == [
        "企业无统一社会信用代码且历史参考库无记录，请人工核实"
    ]


def test_industry_code_difference_is_an_issue_after_trusted_match():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, dwhydm="WRONG"))

    assert result.is_ok is False
    assert result.issues[0].rule == "industry_match"
    assert result.issues[0].student_value == "WRONG"
    assert result.issues[0].reference_value == reference.industry_code


def test_company_type_code_difference_is_an_issue_after_trusted_match():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, dwxzdm="WRONG"))

    assert result.is_ok is False
    assert result.issues[0].rule == "company_type_match"
    assert result.issues[0].student_value == "WRONG"
    assert result.issues[0].reference_value == reference.company_type_code


def test_other_person_job_category_is_an_issue():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, gzzwlb=" 其他人员 "))

    assert [issue.message for issue in result.issues] == ["工作职位类别不能填写为其他人员"]
    assert result.issues[0].field == "gzzwlb"


def test_company_contact_matching_student_phone_is_an_issue():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, mobilePhone="13800138000", lxrdh="+86 138-0013-8000"))

    assert [issue.message for issue in result.issues] == ["用人单位联系人电话与学生本人联系电话相同"]


def test_company_contact_matching_family_phone_is_an_issue():
    reference = _trusted_reference()

    result = _audit(reference, _student_row(reference, jtdh="010-12345678", lxrsj="010 1234 5678"))

    assert [issue.message for issue in result.issues] == ["用人单位联系人电话与学生家庭联系电话相同"]


def test_batch_marks_different_companies_sharing_contact_phone():
    reference = _trusted_reference()
    other = replace(reference, credit_code="911111111111111111", company_name="匿名另一企业")
    rows = [
        _student_row(reference, lxrdh="010-12345678"),
        _student_row(other, lxrsj="010 1234 5678"),
    ]
    engine = AuditEngine(ReferenceLibrary([reference, other]))
    results = [engine.audit(row, index + 2, ReviewStage.FINAL) for index, row in enumerate(rows)]

    apply_batch_contact_phone_duplicates(results, rows)

    assert all(not result.is_ok for result in results)
    assert all(
        any(issue.message == "本批次不同用人单位使用了相同联系人电话，请人工核实" for issue in result.issues)
        for result in results
    )


def test_batch_allows_same_company_to_reuse_hr_phone():
    reference = _trusted_reference()
    rows = [
        _student_row(reference, xh="ANON-001", lxrdh="010-12345678"),
        _student_row(reference, xh="ANON-002", lxrsj="010 1234 5678"),
    ]
    engine = AuditEngine(ReferenceLibrary([reference]))
    results = [engine.audit(row, index + 2, ReviewStage.FINAL) for index, row in enumerate(rows)]

    apply_batch_contact_phone_duplicates(results, rows)

    assert all(result.is_ok for result in results)


def test_one_student_can_retain_multiple_independent_issues():
    reference = _trusted_reference()
    row = _student_row(
        reference,
        dwmc="匿名错误企业",
        gzzwlb="其他人员",
        mobilePhone="13800138000",
        lxrsj="13800138000",
    )

    result = _audit(reference, row)

    assert {issue.rule for issue in result.issues} == {
        "company_name_match",
        "job_category_not_other",
        "company_contact_not_student_phone",
    }
