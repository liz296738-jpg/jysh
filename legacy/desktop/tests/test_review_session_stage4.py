from __future__ import annotations

from pathlib import Path
from shutil import copy2

from src.review_session import ReviewController, filter_results
from src.rules import AuditIssue, ReviewStage, StudentAuditResult
from src.gui import application_root


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SAMPLES = PROJECT_ROOT / "samples"


def test_initial_and_final_sessions_audit_immediately_without_exporting(tmp_path: Path):
    controller = ReviewController(PROJECT_ROOT / "企业参考库.xlsx")
    for stage, filename in (("initial", "初审样例.xlsx"), ("final", "终审样例.xlsx")):
        input_path = tmp_path / filename
        copy2(SAMPLES / filename, input_path)
        expected_output = input_path.with_name(
            f"{input_path.stem}_{'初审' if stage == 'initial' else '终审'}审核结果.xlsx"
        )

        session = controller.audit(input_path, stage)

        assert session.review_stage.value == stage.upper()
        assert session.total_count == (2 if stage == "initial" else 5)
        assert not expected_output.exists()


def test_export_is_deferred_until_requested(tmp_path: Path):
    input_path = tmp_path / "初审样例.xlsx"
    copy2(SAMPLES / "初审样例.xlsx", input_path)
    controller = ReviewController(PROJECT_ROOT / "企业参考库.xlsx")
    session = controller.audit(input_path, "initial")

    output_path = controller.export(session)

    assert output_path.exists()
    assert output_path.name == "初审样例_初审审核结果.xlsx"


def test_filter_results_groups_issues_by_review_category():
    phone_issue = AuditIssue("company_contact_not_student_phone", "mobilePhone", "电话异常", "", "", ())
    job_issue = AuditIssue("job_category_not_other", "gzzwlb", "岗位异常", "", "", ())
    phone_result = StudentAuditResult(2, "匿名学生", "A-1", "匿名企业", ReviewStage.INITIAL, False, [phone_issue])
    job_result = StudentAuditResult(3, "匿名学生", "A-2", "匿名企业", ReviewStage.INITIAL, False, [job_issue])

    assert filter_results((phone_result, job_result), "电话") == (phone_result,)
    assert filter_results((phone_result, job_result), "工作职位") == (job_result,)
    assert filter_results((phone_result, job_result), "全部异常") == (phone_result, job_result)


def test_frozen_app_uses_the_executable_directory_for_the_reference_library(monkeypatch, tmp_path: Path):
    executable = tmp_path / "就业预审工具.exe"
    monkeypatch.setattr("sys.frozen", True, raising=False)
    monkeypatch.setattr("sys.executable", str(executable))

    assert application_root() == tmp_path
