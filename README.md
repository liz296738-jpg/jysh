# 高校就业信息辅助预审系统

用于高校招生就业处的离线式结构化就业数据辅助校验。它不作正式审批决定，不上传学生数据；终审中的协议、合同、PDF 和扫描件仍须人工审核。

## 本地运行

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

访问 `http://127.0.0.1:5000`。生产启动命令：`gunicorn app:app --workers 1 --threads 4 --timeout 120`。

## Render

Build Command：`pip install -r requirements.txt`
Start Command：`gunicorn app:app --workers 1 --threads 4 --timeout 120`

生产环境请设置高强度 `SECRET_KEY`；可设置 `MAX_UPLOAD_MB`（默认 20）和 `TASK_TTL_MINUTES`（默认 120）。上传及审核结果仅保存在 `temp/<task_id>/`，下载后删除，超时自动清理。

企业参考库位于 `企业参考库.xlsx`，只允许企业级必要字段。请勿将真实学生 Excel、`.env` 或临时任务提交到 Git。
