from __future__ import annotations

import json
import os
import shutil
import uuid
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path

from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_file, url_for

from src.audit_workbook import audit_workbook
from src.batch_rules import apply_batch_contact_phone_duplicates
from src.config import FIELD_MAP
from src.excel_parser import ExcelParser
from src.reference import ReferenceLibrary
from src.reference_manager import ReferenceManagementError, ReferenceManager
from src.rules import AuditEngine, ReviewStage


ROOT = Path(__file__).resolve().parent


def _serialize(result):
    return {"row_number": result.row_number, "student_name": result.student_name, "student_id": result.student_id, "company_name": result.company_name, "is_ok": result.is_ok, "issues": [{"rule": i.rule, "field": i.field, "message": i.message, "student_value": i.student_value, "reference_value": i.reference_value} for i in result.issues]}


def create_app(config=None):
    app = Flask(__name__)
    app.config.update(SECRET_KEY=os.getenv("SECRET_KEY") or os.urandom(32), MAX_CONTENT_LENGTH=int(os.getenv("MAX_UPLOAD_MB", "20")) * 1024 * 1024, TEMP_ROOT=ROOT / "temp", REFERENCE_PATH=ROOT / "企业参考库.xlsx", TASK_TTL_MINUTES=int(os.getenv("TASK_TTL_MINUTES", "120")), REFERENCE_BACKUP_ROOT=ROOT / "backups" / "reference", REFERENCE_CHANGE_LOG=ROOT / "reference_change_log.jsonl", REFERENCE_ADMIN_PASSWORD=os.getenv("REFERENCE_ADMIN_PASSWORD", ""))
    if config:
        app.config.update(config)
    Path(app.config["TEMP_ROOT"]).mkdir(parents=True, exist_ok=True)

    def cleanup():
        cutoff = datetime.now(timezone.utc) - timedelta(minutes=app.config["TASK_TTL_MINUTES"])
        for item in Path(app.config["TEMP_ROOT"]).iterdir():
            if item.is_dir() and datetime.fromtimestamp(item.stat().st_mtime, timezone.utc) < cutoff:
                shutil.rmtree(item, ignore_errors=True)

    def task_path(task_id):
        path = Path(app.config["TEMP_ROOT"]) / task_id
        if not path.is_dir(): abort(404)
        return path

    def reference_ready():
        try: ReferenceLibrary.from_workbook(app.config["REFERENCE_PATH"], require_high_confidence=True); return True
        except Exception: return False

    def manager():
        return ReferenceManager(app.config["REFERENCE_PATH"], app.config["REFERENCE_BACKUP_ROOT"], app.config["REFERENCE_CHANGE_LOG"])


    @app.before_request
    def _cleanup(): cleanup()

    @app.get("/")
    def index():
        try: reference_count = len(ReferenceLibrary.from_workbook(app.config["REFERENCE_PATH"], require_high_confidence=True).by_credit_code)
        except Exception: reference_count = 0
        return render_template("index.html", reference_ready=reference_ready(), reference_count=reference_count)

    @app.get("/reference")
    def reference_list():
        try:
            rows = manager().records()
        except Exception:
            return render_template("error.html", message="企业参考库异常，请联系管理员。"), 503
        query = request.args.get("q", "").strip()
        if query:
            normalized = query.upper().replace(" ", "")
            rows = [row for row in rows if query.casefold() in row["dwmc"].casefold() or normalized in row["dwzzjgdm"].upper()]
        per_page = 100 if request.args.get("per_page") == "100" else 50
        page = max(request.args.get("page", 1, type=int), 1)
        total_pages = max((len(rows) + per_page - 1) // per_page, 1)
        page = min(page, total_pages)
        start = (page - 1) * per_page
        return render_template("reference_list.html", rows=rows[start:start + per_page], total=len(manager().records()), query=query, page=page, total_pages=total_pages, per_page=per_page, is_admin=True)

    @app.route("/reference/new", methods=["GET", "POST"])
    def reference_new():
        current_manager = manager()
        if request.method == "POST":
            try:
                current_manager.save("ADD", request.form)
            except ReferenceManagementError as error:
                return render_template("reference_form.html", record=request.form, choices=current_manager.choices(), title="新增企业", action_url=url_for("reference_new"), error=str(error)), 400
            flash("企业已成功加入高可信参考库。")
            return redirect(url_for("reference_list"))
        return render_template("reference_form.html", record={}, choices=current_manager.choices(), title="新增企业", action_url=url_for("reference_new"), error=None)

    @app.route("/reference/<credit_code>/edit", methods=["GET", "POST"])
    def reference_edit(credit_code):
        current_manager = manager()
        original = current_manager.find(credit_code)
        if original is None: abort(404)
        if request.method == "POST":
            try:
                current_manager.save("EDIT", request.form, credit_code)
            except ReferenceManagementError as error:
                return render_template("reference_form.html", record=request.form, choices=current_manager.choices(), title="编辑企业", action_url=url_for("reference_edit", credit_code=credit_code), error=str(error)), 400
            flash("企业信息已更新并保存。")
            return redirect(url_for("reference_list", q=request.form.get("dwzzjgdm", "")))
        return render_template("reference_form.html", record=original, choices=current_manager.choices(), title="编辑企业", action_url=url_for("reference_edit", credit_code=credit_code), error=None)

    @app.get("/health")
    def health(): return jsonify(status="ok")

    @app.post("/audit/<stage>")
    def audit(stage):
        if stage not in {"initial", "final"}: abort(404)
        if not reference_ready(): return render_template("error.html", message="参考库不可用，请联系维护人员。"), 503
        upload = request.files.get("file")
        if not upload or not upload.filename or not upload.filename.lower().endswith(".xlsx"):
            return render_template("error.html", message="仅支持 .xlsx 文件。"), 400
        task_id = uuid.uuid4().hex; folder = Path(app.config["TEMP_ROOT"]) / task_id; folder.mkdir()
        input_path = folder / "original.xlsx"; upload.save(input_path)
        try:
            parsed = ExcelParser(FIELD_MAP).read(input_path); engine = AuditEngine(ReferenceLibrary.from_workbook(app.config["REFERENCE_PATH"], require_high_confidence=True)); review_stage = ReviewStage(stage.upper()); results = [engine.audit(row, index + 2, review_stage) for index, row in enumerate(parsed.rows)]; apply_batch_contact_phone_duplicates(results, parsed.rows)
        except Exception as error:
            shutil.rmtree(folder, ignore_errors=True); return render_template("error.html", message=f"Excel 无法审核：{error}"), 400
        data = {"task_id": task_id, "review_stage": stage, "created_at": datetime.now(timezone.utc).isoformat(), "results": [_serialize(r) for r in results]}
        (folder / "result.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return redirect(url_for("results", task_id=task_id))

    @app.get("/results/<task_id>")
    def results(task_id):
        data = json.loads((task_path(task_id) / "result.json").read_text(encoding="utf-8")); total=len(data["results"]); normal=sum(r["is_ok"] for r in data["results"]); response=app.make_response(render_template("results.html", data=data, total=total, normal=normal, problem=total-normal)); response.headers["Cache-Control"]="no-store"; return response

    @app.post("/export/<task_id>")
    def export(task_id):
        folder=task_path(task_id); data=json.loads((folder / "result.json").read_text(encoding="utf-8")); outcome=audit_workbook(folder / "original.xlsx", data["review_stage"], app.config["REFERENCE_PATH"], require_high_confidence_reference=True)
        workbook_bytes = outcome.output_path.read_bytes()
        download_name = outcome.output_path.name
        shutil.rmtree(folder, ignore_errors=True)
        return send_file(BytesIO(workbook_bytes), as_attachment=True, download_name=download_name, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return app


app = create_app()

if __name__ == "__main__": app.run(debug=False)
