#!/usr/bin/env python3
"""
受講人数集計 CLI スクリプト

lists/ 内のすべての Excel ファイルを処理し、
月別の受講人数を集計して Pivot 形式の Excel ファイルを出力します。

使用方法:
  python scripts/aggregate.py
"""

import sys
from pathlib import Path

# プロジェクトルートを sys.path に追加
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from services.aggregator import (
    load_excel,
    parse_target_month,
    aggregate,
    save_monthly_result,
    build_pivot,
    to_excel_bytes,
)


def main():
    """メイン処理"""
    lists_dir = project_root / "lists"
    output_dir = project_root / "outputs"
    results_dir = output_dir / "results"
    output_file = output_dir / "monthly_stats.xlsx"

    # lists ディレクトリの確認
    if not lists_dir.exists():
        print(f"Error: {lists_dir} が見つかりません")
        return 1

    xlsx_files = sorted(lists_dir.glob("*.xlsx"))
    if not xlsx_files:
        print(f"Warning: {lists_dir} に Excel ファイルがありません")
        return 0

    # 既存の CSV を削除
    results_dir.mkdir(parents=True, exist_ok=True)
    for csv in results_dir.glob("*.csv"):
        csv.unlink()

    print(f"Processing {len(xlsx_files)} files from {lists_dir.name}/")
    print("-" * 60)

    processed = 0
    for file_path in xlsx_files:
        target_month = parse_target_month(file_path.name)
        if not target_month:
            print(f"  Skipped: {file_path.name} (invalid filename)")
            continue

        try:
            df = load_excel(file_path)
            result = aggregate(df, target_month)
            if result is not None and len(result) > 0:
                save_monthly_result(result, target_month, results_dir)
                print(f"  {target_month}: {len(result)} rows")
                processed += 1
            else:
                print(f"  {target_month}: no data")
        except Exception as e:
            print(f"  Error ({target_month}): {e}")
            return 1

    print("-" * 60)

    if processed == 0:
        print("Error: No data processed")
        return 1

    # Pivot 生成
    print(f"\nGenerating pivot...")
    pivot = build_pivot(results_dir)

    if pivot.empty:
        print("Error: Failed to generate pivot")
        return 1

    # Excel 出力
    excel_bytes = to_excel_bytes(pivot)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file.write_bytes(excel_bytes)

    # 結果表示
    print(f"Output: {output_file}")
    print(f"  Rows: {pivot.shape[0]}")
    print(f"  Columns: {pivot.shape[1]}")
    print(f"  Size: {output_file.stat().st_size / 1024:.1f} KB")

    # 月別統計
    month_cols = [col for col in pivot.columns if col not in ["学年", "教室", "講座名", "M/C", "担当"]]
    if month_cols:
        print(f"\nAnnual Summary:")
        total = 0
        for month in month_cols:
            count = int(pivot[month].sum())
            total += count
            print(f"  {month}: {count:,}")
        print(f"  Total: {total:,}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
