from io import BytesIO
from pathlib import Path

import pytest

from app import create_app


ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture()
def client(tmp_path):
    app = create_app({"TESTING": True, "TEMP_ROOT": tmp_path / "temp", "REFERENCE_PATH": ROOT / "企业参考库.xlsx"})
    return app.test_client()


def test_home_and_health(client):
    assert client.get("/").status_code == 200
    assert "上传初审数据" in client.get("/").get_data(as_text=True)
    assert client.get("/health").get_json() == {"status": "ok"}


def test_rejects_non_xlsx_upload(client):
    response = client.post("/audit/initial", data={"file": (BytesIO(b"no"), "bad.csv")})
    assert response.status_code == 400


def test_initial_upload_result_and_export(client):
    payload = (ROOT / "samples" / "初审样例.xlsx").read_bytes()
    response = client.post("/audit/initial", data={"file": (BytesIO(payload), "upload.xlsx")}, content_type="multipart/form-data")
    assert response.status_code == 302
    result = client.get(response.headers["Location"])
    assert result.status_code == 200
    assert "初审" in result.get_data(as_text=True)
    assert result.headers["Cache-Control"] == "no-store"
    task_id = response.headers["Location"].rsplit("/", 1)[-1]
    exported = client.post(f"/export/{task_id}")
    assert exported.status_code == 200
    assert exported.headers["Content-Type"].startswith("application/vnd.openxmlformats")


def test_unknown_task_is_not_found(client):
    assert client.get("/results/does-not-exist").status_code == 404


def test_final_upload_shows_result_page_and_export_cleans_task(client):
    payload = (ROOT / "samples" / "终审样例.xlsx").read_bytes()
    response = client.post(
        "/audit/final",
        data={"file": (BytesIO(payload), "upload.xlsx")},
        content_type="multipart/form-data",
    )
    assert response.status_code == 302
    task_id = response.headers["Location"].rsplit("/", 1)[-1]
    assert client.get(response.headers["Location"]).status_code == 200
    assert (Path(client.application.config["TEMP_ROOT"]) / task_id).is_dir()
    exported = client.post(f"/export/{task_id}")
    assert exported.status_code == 200
    assert not (Path(client.application.config["TEMP_ROOT"]) / task_id).exists()


def test_rejects_workbook_without_required_headers(client):
    from openpyxl import Workbook

    workbook = Workbook()
    workbook.active.append(["not-a-supported-header"])
    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    response = client.post(
        "/audit/initial",
        data={"file": (buffer, "invalid.xlsx")},
        content_type="multipart/form-data",
    )
    assert response.status_code == 400


def test_rejects_upload_over_limit(tmp_path):
    app = create_app(
        {
            "TESTING": True,
            "TEMP_ROOT": tmp_path / "temp",
            "REFERENCE_PATH": ROOT / "企业参考库.xlsx",
            "MAX_CONTENT_LENGTH": 1,
        }
    )
    response = app.test_client().post(
        "/audit/initial",
        data={"file": (BytesIO(b"too-large"), "upload.xlsx")},
        content_type="multipart/form-data",
    )
    assert response.status_code == 413
