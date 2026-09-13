from pathlib import Path

import pytest
from openpyxl import Workbook

from app import create_app


def _write_reference(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["dwmc", "dwzzjgdm", "dwxz", "dwhy"])
    sheet.append(["甲有限公司", "911111111111111111", "民营企业", "制造业"])
    sheet.append(["乙有限公司", "922222222222222222", "国有企业", "信息传输、软件和信息技术服务业"])
    workbook.save(path)
    workbook.close()


@pytest.fixture()
def managed_client(tmp_path):
    reference_path = tmp_path / "企业参考库.xlsx"
    _write_reference(reference_path)
    app = create_app(
        {
            "TESTING": True,
            "TEMP_ROOT": tmp_path / "temp",
            "REFERENCE_PATH": reference_path,
            "REFERENCE_BACKUP_ROOT": tmp_path / "backups" / "reference",
            "REFERENCE_CHANGE_LOG": tmp_path / "reference_change_log.jsonl",
            "REFERENCE_ADMIN_PASSWORD": "test-admin-password",
        }
    )
    return app.test_client(), reference_path, tmp_path


def _login(client):
    response = client.post("/reference/login", data={"password": "test-admin-password"})
    assert response.status_code == 302


def test_reference_page_lists_and_searches_by_name_or_credit_code(managed_client):
    client, _, _ = managed_client
    page = client.get("/reference")
    assert page.status_code == 200
    assert "高可信企业参考库" in page.get_data(as_text=True)
    assert "当前共 2 家企业" in page.get_data(as_text=True)
    assert "甲有限公司" in page.get_data(as_text=True)
    assert "乙有限公司" not in client.get("/reference?q=甲有").get_data(as_text=True)
    assert "甲有限公司" in client.get("/reference?q=甲有").get_data(as_text=True)
    assert "乙有限公司" in client.get("/reference?q=922222").get_data(as_text=True)


def test_add_and_edit_are_persisted_with_backup_and_audit_log(managed_client):
    client, reference_path, tmp_path = managed_client
    _login(client)
    added = client.post(
        "/reference/new",
        data={"dwmc": "丙有限公司", "dwzzjgdm": "933333333333333333", "dwxz": "民营企业", "dwhy": "制造业"},
    )
    assert added.status_code == 302
    assert "丙有限公司" in client.get("/reference?q=933333").get_data(as_text=True)
    edited = client.post(
        "/reference/933333333333333333/edit",
        data={"dwmc": "丙有限公司", "dwzzjgdm": "933333333333333333", "dwxz": "国有企业", "dwhy": "制造业"},
    )
    assert edited.status_code == 302
    page = client.get("/reference?q=933333").get_data(as_text=True)
    assert "国有企业" in page
    assert reference_path.exists()
    assert list((tmp_path / "backups" / "reference").glob("*.xlsx"))
    assert '"action": "ADD"' in (tmp_path / "reference_change_log.jsonl").read_text(encoding="utf-8")
    assert '"action": "EDIT"' in (tmp_path / "reference_change_log.jsonl").read_text(encoding="utf-8")


def test_reference_management_rejects_invalid_or_conflicting_writes(managed_client):
    client, _, _ = managed_client
    _login(client)
    duplicate_code = client.post(
        "/reference/new",
        data={"dwmc": "新名称", "dwzzjgdm": "911111111111111111", "dwxz": "民营企业", "dwhy": "制造业"},
    )
    assert duplicate_code.status_code == 400
    duplicate_name = client.post(
        "/reference/new",
        data={"dwmc": "甲有限公司", "dwzzjgdm": "933333333333333333", "dwxz": "民营企业", "dwhy": "制造业"},
    )
    assert duplicate_name.status_code == 400
    invalid = client.post(
        "/reference/new",
        data={"dwmc": "新名称", "dwzzjgdm": "30", "dwxz": "民营企业", "dwhy": "制造业"},
    )
    assert invalid.status_code == 400


def test_writes_require_an_authenticated_admin_session(managed_client):
    client, _, _ = managed_client
    response = client.post(
        "/reference/new",
        data={"dwmc": "丙有限公司", "dwzzjgdm": "933333333333333333", "dwxz": "民营企业", "dwhy": "制造业"},
    )
    assert response.status_code == 403
