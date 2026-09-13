from __future__ import annotations

from functools import cache
from pathlib import Path
import warnings

import openpyxl
import pytest

from src.config import FIELD_MAP
from src.excel_parser import ExcelParser, MissingRequiredColumnsError
from src.reference import (
    REFERENCE_HEADERS,
    ReferenceLibrary,
    build_reference_records,
    write_reference_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SAMPLES = PROJECT_ROOT / "samples"


@cache
def _parsed_sample(filename: str):
    return ExcelParser(FIELD_MAP).read(SAMPLES / filename)


def _write_workbook(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def _history_headers() -> list[str]:
    return list(dict.fromkeys(FIELD_MAP.values()))


def _history_row(**overrides: object) -> list[object]:
    values: dict[str, object] = {header: "" for header in _history_headers()}
    values.update(
        {
            "jyshzt": "审核完成",
            "dwmc": "示例企业",
            "dwzzjgdm": "91310000TEST00001X",
            "dwxz": "民营企业",
            "dwxzdm": "300",
            "dwhy": "信息传输、软件和信息技术服务业",
            "dwhydm": "I",
        }
    )
    values.update(overrides)
    return [values[header] for header in _history_headers()]


def test_initial_and_final_samples_are_readable_with_one_field_map() -> None:
    initial = _parsed_sample("初审样例.xlsx")
    final = _parsed_sample("终审样例.xlsx")

    assert initial.headers == final.headers
    assert initial.headers.issuperset(FIELD_MAP.values())
    assert initial.row_count == 2
    assert final.row_count == 5


def test_history_sample_is_readable() -> None:
    parsed = _parsed_sample("历史参考总库.xlsx")

    assert parsed.row_count == 2
    assert len(parsed.headers) == 101


def test_parser_silences_known_missing_default_style_warning() -> None:
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        ExcelParser(FIELD_MAP).read(SAMPLES / "历史参考总库.xlsx")

    assert not any("Workbook contains no default style" in str(item.message) for item in caught)


def test_only_completed_history_rows_become_reference_samples(tmp_path: Path) -> None:
    history = tmp_path / "history.xlsx"
    _write_workbook(
        history,
        _history_headers(),
        [_history_row(), _history_row(jyshzt="审核中", dwmc="不应进入企业")],
    )

    records = build_reference_records(ExcelParser(FIELD_MAP).read(history).rows)

    assert len(records) == 1
    assert records[0].company_name == "示例企业"


def test_reference_workbook_excludes_student_sensitive_fields(tmp_path: Path) -> None:
    output = tmp_path / "企业参考库.xlsx"
    write_reference_workbook(output, build_reference_records([dict(zip(_history_headers(), _history_row()))]))

    workbook = openpyxl.load_workbook(output, read_only=True)
    headers = [cell.value for cell in workbook.active[1]]
    workbook.close()

    assert headers == REFERENCE_HEADERS
    assert not set(headers) & {"xm", "xh", "sfzh", "mobilePhone", "jtdh", "jtdz", "email"}


def test_consistent_credit_code_creates_non_conflicting_reference_record() -> None:
    records = build_reference_records(
        [
            dict(zip(_history_headers(), _history_row(dwzzjgdm=" 91310000TEST00001X "))),
            dict(zip(_history_headers(), _history_row(dwzzjgdm="91310000test00001x"))),
        ]
    )

    assert len(records) == 1
    assert records[0].credit_code == "91310000TEST00001X"
    assert records[0].reference_conflict is False


def test_conflicting_credit_code_is_retained_and_marked_for_manual_review() -> None:
    records = build_reference_records(
        [
            dict(zip(_history_headers(), _history_row())),
            dict(zip(_history_headers(), _history_row(dwmc="另一企业名称"))),
        ]
    )

    assert len(records) == 1
    assert records[0].reference_conflict is True
    assert "company_name" in records[0].conflict_fields


def test_records_without_credit_code_are_grouped_by_normalized_company_name() -> None:
    records = build_reference_records(
        [
            dict(zip(_history_headers(), _history_row(dwzzjgdm="", dwmc="  示例 企业  "))),
            dict(zip(_history_headers(), _history_row(dwzzjgdm="", dwmc="示例 企业"))),
        ]
    )

    assert len(records) == 1
    assert records[0].credit_code == ""
    assert records[0].reference_conflict is False


def test_reference_library_prefers_credit_code_and_uses_name_only_when_code_empty() -> None:
    records = build_reference_records(
        [
            dict(zip(_history_headers(), _history_row())),
            dict(zip(_history_headers(), _history_row(dwzzjgdm="", dwmc="无代码企业"))),
        ]
    )
    library = ReferenceLibrary(records)

    assert library.find("91310000TEST00001X", "不存在的企业").credit_code == "91310000TEST00001X"
    assert library.find("", "示例企业") is None
    assert library.find("", "无代码企业").credit_code == ""
    assert library.find("错误代码", "示例企业") is None


def test_missing_required_column_raises_clear_error(tmp_path: Path) -> None:
    path = tmp_path / "missing.xlsx"
    headers = _history_headers()
    headers.remove("gzzwlb")
    _write_workbook(path, headers, [])

    with pytest.raises(MissingRequiredColumnsError, match="gzzwlb"):
        ExcelParser(FIELD_MAP).read(path)
