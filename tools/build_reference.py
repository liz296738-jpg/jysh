"""Build the offline enterprise reference workbook from the completed history export."""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from src.config import FIELD_MAP
from src.excel_parser import ExcelParser
from src.reference import build_reference_records, write_reference_workbook


def main() -> None:
    history_path = PROJECT_ROOT / "samples" / "历史参考总库.xlsx"
    output_path = PROJECT_ROOT / "企业参考库.xlsx"
    history = ExcelParser(FIELD_MAP).read(history_path)
    records = build_reference_records(history.rows)
    write_reference_workbook(output_path, records)
    conflict_count = sum(record.reference_conflict for record in records)
    print(f"历史有效样本数量：{sum(row.get('jyshzt') == '审核完成' for row in history.rows)}")
    print(f"参考企业数量：{len(records)}")
    print(f"历史冲突企业数量：{conflict_count}")


if __name__ == "__main__":
    main()
