from tools.build_high_confidence_reference import (
    build_trusted_reference,
    is_missing_reference_value,
    normalize_credit_code,
    normalize_text,
    validate_final_reference,
    validate_final_workbook,
    write_outputs,
)


def row(number, name, code, nature, industry):
    return {
        "row_number": number,
        "dwmc": name,
        "dwzzjgdm": code,
        "dwxz": nature,
        "dwhy": industry,
    }


def test_normalization_only_changes_allowed_formatting():
    assert normalize_credit_code(" 91ab 12\t ") == "91AB12"
    assert normalize_text("\u00a0 某有限公司\u200b ") == "某有限公司"
    assert is_missing_reference_value("30")
    assert is_missing_reference_value("   ")


def test_identical_records_are_reduced_to_one_trusted_company():
    outcome = build_trusted_reference(
        [row(index, "甲有限公司", "91AAA", "民营企业", "制造业") for index in range(2, 12)]
    )
    assert outcome.trusted_rows == [{"dwmc": "甲有限公司", "dwzzjgdm": "91AAA", "dwxz": "民营企业", "dwhy": "制造业"}]
    assert outcome.statistics["complete_duplicate_historical_records"] == 9
    validate_final_reference(outcome.trusted_rows)


def test_conflicting_name_nature_and_industry_remove_entire_code_group():
    outcome = build_trusted_reference(
        [
            row(2, "甲有限公司", "91AAA", "民营企业", "制造业"),
            row(3, "乙有限公司", "91AAA", "国有企业", "信息传输、软件和信息技术服务业"),
        ]
    )
    assert outcome.trusted_rows == []
    reasons = outcome.removed_by_code["91AAA"]["reasons"]
    assert "同一统一社会信用代码对应多个单位名称" in reasons
    assert "同一企业历史单位性质不一致" in reasons
    assert "同一企业历史单位行业不一致" in reasons


def test_missing_placeholder_does_not_create_conflict_but_all_missing_field_removes_code():
    retained = build_trusted_reference(
        [
            row(2, "甲有限公司", "91AAA", "民营企业", "制造业"),
            row(3, "甲有限公司", "91AAA", "30", "30"),
        ]
    )
    assert len(retained.trusted_rows) == 1
    removed = build_trusted_reference([row(2, "甲有限公司", "91AAA", "30", "30")])
    assert removed.trusted_rows == []
    assert "单位性质无有效历史参考值" in removed.removed_by_code["91AAA"]["reasons"]
    assert "单位行业无有效历史参考值" in removed.removed_by_code["91AAA"]["reasons"]


def test_same_name_with_multiple_codes_removes_all_involved_codes():
    outcome = build_trusted_reference(
        [
            row(2, "甲有限公司", "91AAA", "民营企业", "制造业"),
            row(3, "甲有限公司", "91BBB", "民营企业", "制造业"),
        ]
    )
    assert outcome.trusted_rows == []
    for code in ("91AAA", "91BBB"):
        assert "同一单位名称对应多个统一社会信用代码" in outcome.removed_by_code[code]["reasons"]


def test_missing_credit_code_never_enters_trusted_reference():
    outcome = build_trusted_reference([row(2, "甲有限公司", "30", "民营企业", "制造业")])
    assert outcome.trusted_rows == []
    assert outcome.statistics["invalid_credit_code_records"] == 1


def test_written_reference_is_independently_reloaded_and_validated(tmp_path):
    outcome = build_trusted_reference([row(2, "甲有限公司", "91AAA", "民营企业", "制造业")])
    input_path = tmp_path / "source.xlsx"
    input_path.touch()
    paths = write_outputs(input_path, tmp_path / "企业参考库_高可信.xlsx", outcome)
    validate_final_workbook(paths["trusted"])
