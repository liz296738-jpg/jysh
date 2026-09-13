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

## 企业高可信参考库

正式审核只读取项目根目录的 `企业参考库.xlsx`。当前版本包含 4356 家高可信企业，字段固定为 `dwmc`、`dwzzjgdm`、`dwxz`、`dwhy`。

该库由历史就业数据按“宁缺毋滥”原则完成一致性筛选后静态发布；同一信用代码存在多个名称、单位性质或行业，同一名称存在多个信用代码，核心字段缺失或为占位值 `30` 的企业均不进入正式库。程序加载时会校验该库的列结构、企业数量、核心字段及信用代码唯一性；校验失败时不会继续审核。

未命中高可信参考库不表示企业信息必然错误，只表示需要人工核实。旧版 `tools/build_reference.py` 已停用，不能再由历史样例重新生成或覆盖正式参考库。

### 参考库管理

访问 `/reference` 可分页查看并按单位名称或统一社会信用代码包含搜索。新增和编辑需要设置环境变量 `REFERENCE_ADMIN_PASSWORD` 后通过管理员口令登录；未设置该变量时，管理写入默认拒绝。

每次保存均由后端校验四项字段、18 位信用代码格式、信用代码唯一性和名称—信用代码一对一关系。系统先写入并校验临时 Excel，再原子替换正式库；旧文件保存到 `backups/reference/`（保留最近 30 个），企业字段变更写入 `reference_change_log.jsonl`。这些运行时文件不提交 Git。

Render 的默认本地磁盘不保证重启或重新部署后的数据持久性。因此生产启用参考库编辑前，必须将正式库、备份目录和变更日志放到持久化磁盘或其他经批准的持久化存储；本次未改变部署架构。
