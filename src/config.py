"""Centralized source-column definitions for employment exports."""

FIELD_MAP = {
    "student_name": "xm",
    "student_id": "xh",
    "student_phone": "mobilePhone",
    "family_phone": "jtdh",
    "company_name": "dwmc",
    "credit_code": "dwzzjgdm",
    "company_type": "dwxz",
    "company_type_code": "dwxzdm",
    "industry": "dwhy",
    "industry_code": "dwhydm",
    "company_contact": "dwlxr",
    "company_phone": "lxrdh",
    "company_mobile": "lxrsj",
    "job_name": "gwmc",
    "job_category": "gzzwlb",
    "job_category_code": "gzzwlbdm",
    "audit_status": "jyshzt",
    "reason": "reason",
}

TRUSTED_AUDIT_STATUS = "审核完成"

# 当前静态高可信企业参考库的数据完整性校验值，不是审核业务规则。
EXPECTED_REFERENCE_COUNT = 4356
HIGH_CONFIDENCE_REFERENCE_HEADERS = ("dwmc", "dwzzjgdm", "dwxz", "dwhy")
