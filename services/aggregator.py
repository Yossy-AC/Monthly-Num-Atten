"""
受講人数集計ロジック

Row 4をヘッダーとして読み込み、1ヶ月分を集計してストアに保存。
build_pivot() で全月分をマージしてピボットを生成する。
"""
from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd

# ── 列番号マッピング（0-indexed） ──
# Row 4がヘッダー：A=0, B=1, ..., AA=26
COLUMN_INDICES = {
    "add_date": 2,      # Column C: 受講追加日付
    "cancel_date": 6,   # Column G: 受講取消日付
    "course": 9,        # Column J: 講座名
    "class_type": 10,   # Column K: 【マスター】【コア】
    "classroom": 11,    # Column L: 受講教室
    "grade": 15,        # Column P: 学年コード (31=高1, 32=高2, 33=高3)
    "teacher": 26,      # Column AA: 担当
    "gender": 17,       # Column R: 性別
    "school": 18,       # Column S: 在籍校
    "department": 28,   # Column AC: 学科
}

# 学年コード → 表示名
GRADE_LABELS = {31: "高1", 32: "高2", 33: "高3"}
# 集計対象学年コード
TARGET_GRADES = set(GRADE_LABELS.keys())

KEY_COLS = ["学年", "教室", "講座名", "マスター/コア", "担当", "在籍校", "学科", "性別"]
MONTH_ORDER = ["4月", "5月", "6月", "7月", "8月", "9月",
               "10月", "11月", "12月", "1月", "2月", "3月"]

RESULTS_DIR = Path("outputs/results")


def parse_target_month(filename: str) -> pd.Period | None:
    """ファイル名末尾の _YYMM からターゲット月を抽出。例: _2409 → 2024-09"""
    m = re.search(r"_(\d{2})(\d{2})\.", filename)
    if m:
        year = 2000 + int(m.group(1))
        month = int(m.group(2))
        return pd.Period(f"{year}-{month:02d}", freq="M")
    return None


def load_excel(file: bytes | Path) -> pd.DataFrame:
    if isinstance(file, bytes):
        return pd.read_excel(io.BytesIO(file), header=3)
    return pd.read_excel(file, header=3)


def _resolve_course_name(course: str, class_type: str) -> str:
    """アドバンス/ハイレベル講座は K列の値を付加して別講座扱い"""
    if pd.isna(class_type):
        return course

    class_str = str(class_type).strip()
    course_str = str(course).strip()

    if "ｱﾄﾞﾊﾞﾝｽ" in course_str or "ﾊｲﾚﾍﾞﾙ" in course_str:
        return f"{course_str}{class_str}"

    return course_str


def aggregate(df: pd.DataFrame, target_month: pd.Period | None = None) -> pd.DataFrame:
    """
    対象月1ヶ月分の受講予定人数を集計して返す。

    基準日 = target_month の前月末（例：2024-09 → 2024-08-31）

    Returns
    -------
    pd.DataFrame
        列: 学年, 教室, 講座名, マスター/コア, 担当, {N月}
    """
    idx = COLUMN_INDICES

    add_date_col = pd.to_datetime(df.iloc[:, idx["add_date"]], errors="coerce")
    cancel_date_col = pd.to_datetime(df.iloc[:, idx["cancel_date"]], errors="coerce")

    valid_dates = add_date_col.dropna()
    if len(valid_dates) == 0:
        return pd.DataFrame()

    if target_month is None:
        target_month = valid_dates.max().to_period("M") + 1

    # 基準日 = target_month の前月末
    cutoff_month = target_month - 1
    month_end = cutoff_month.to_timestamp(freq="M")
    month_label = f"{target_month.month}月"

    # 基準日時点で受講中の行をフィルタ
    active_mask = (add_date_col <= month_end) & (
        cancel_date_col.isna() | (cancel_date_col > month_end)
    )

    if not active_mask.any():
        return pd.DataFrame()

    active_df = df[active_mask].copy()

    # 集計対象学年のみに絞り込み（31=高1, 32=高2, 33=高3）
    grade_col = active_df.iloc[:, idx["grade"]]
    grade_mask = grade_col.isin(TARGET_GRADES)
    active_df = active_df[grade_mask].copy()
    if active_df.empty:
        return pd.DataFrame()

    # 学年コードを表示名に変換
    active_df["_grade_label"] = active_df.iloc[:, idx["grade"]].map(GRADE_LABELS)

    active_df["_course_key"] = active_df.apply(
        lambda r: _resolve_course_name(r.iloc[idx["course"]], r.iloc[idx["class_type"]]),
        axis=1,
    )
    active_df["_class_type"] = active_df.iloc[:, idx["class_type"]].fillna("")

    grouped = (
        active_df.groupby(
            [
                "_grade_label",
                active_df.iloc[:, idx["classroom"]],
                "_course_key",
                "_class_type",
                active_df.iloc[:, idx["teacher"]],
                active_df.iloc[:, idx["school"]],
                active_df.iloc[:, idx["department"]],
                active_df.iloc[:, idx["gender"]],
            ],
            dropna=False,
        )
        .size()
        .reset_index(name=month_label)
    )

    grouped.columns = KEY_COLS + [month_label]
    return grouped


def save_monthly_result(df: pd.DataFrame, target_month: pd.Period,
                        results_dir: Path = RESULTS_DIR) -> None:
    """1ヶ月分の集計結果を CSV に保存（同月を再アップロードすると上書き）"""
    results_dir.mkdir(parents=True, exist_ok=True)
    df.to_csv(results_dir / f"{target_month}.csv", index=False, encoding="utf-8-sig")


def build_pivot(results_dir: Path = RESULTS_DIR) -> pd.DataFrame:
    """保存済みの全月分を読み込み、年度ピボットを生成する"""
    files = sorted(results_dir.glob("*.csv"))
    if not files:
        return pd.DataFrame()

    # 月ラベル → DataFrame のマッピングを構築
    month_dfs: dict[str, pd.DataFrame] = {}
    for f in files:
        month_df = pd.read_csv(f, dtype=str)
        month_col = [c for c in month_df.columns if c not in KEY_COLS]
        if month_col:
            month_dfs[month_col[0]] = month_df

    if not month_dfs:
        return pd.DataFrame()

    # 会計年度順にマージ
    result: pd.DataFrame | None = None
    for month_label in MONTH_ORDER:
        if month_label not in month_dfs:
            continue
        month_df = month_dfs[month_label]
        if result is None:
            result = month_df.copy()
        else:
            result = result.merge(month_df, on=KEY_COLS, how="outer")

    if result is None:
        return pd.DataFrame()

    # 欠損を 0 で埋める
    available_months = [c for c in MONTH_ORDER if c in result.columns]
    for col in available_months:
        result[col] = result[col].fillna(0).astype(int)

    result = result[KEY_COLS + available_months]
    return result


def to_excel_bytes(result: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="月次受講人数")
    return buf.getvalue()
