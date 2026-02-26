"""
受講人数集計ロジック

Row 4をヘッダーとして読み込み、1ヶ月分を集計してCSV保存。
build_pivot() で全月分をマージしてピボットを生成する。
"""
from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd

# ── 列番号マッピング（0-indexed、Row 4がヘッダー） ──
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
TARGET_GRADES = set(GRADE_LABELS.keys())

KEY_COLS = ["学年", "教室", "講座名", "M/C", "担当"]
MONTH_ORDER = ["4月", "5月", "6月", "7月", "8月", "9月",
               "10月", "11月", "12月", "1月", "2月", "3月"]

RESULTS_DIR = Path("outputs/results")


def parse_target_month(filename: str) -> pd.Period | None:
    """ファイル名末尾の _YYMM からターゲット月を抽出。例: _2504 → 2025-04"""
    m = re.search(r"_(\d{2})(\d{2})\.", filename)
    if not m:
        return None
    year_suffix = int(m.group(1))
    month = int(m.group(2))
    # YY = 年度下2桁。4月～12月はそのまま、1月～3月は翌暦年
    fiscal_year = 2000 + year_suffix
    year = fiscal_year if month >= 4 else fiscal_year + 1
    return pd.Period(f"{year}-{month:02d}", freq="M")


def load_excel(file: bytes | Path) -> pd.DataFrame:
    src = io.BytesIO(file) if isinstance(file, bytes) else file
    return pd.read_excel(src, header=3)


def aggregate(df: pd.DataFrame, target_month: pd.Period | None = None) -> pd.DataFrame:
    """対象月1ヶ月分の受講人数を集計（全操作ベクトル化）"""
    idx = COLUMN_INDICES

    add_date = pd.to_datetime(df.iloc[:, idx["add_date"]], errors="coerce", format="mixed")
    cancel_date = pd.to_datetime(df.iloc[:, idx["cancel_date"]], errors="coerce", format="mixed")

    if add_date.dropna().empty:
        return pd.DataFrame()

    if target_month is None:
        target_month = add_date.dropna().max().to_period("M") + 1

    # 基準日 = target_month の前月末
    cutoff = (target_month - 1).to_timestamp(freq="M")
    month_label = f"{target_month.month}月"

    # アクティブ行フィルタ
    active = (add_date <= cutoff) & (cancel_date.isna() | (cancel_date > cutoff))
    if not active.any():
        return pd.DataFrame()

    # 学年フィルタ
    grade = df.iloc[:, idx["grade"]]
    mask = active & grade.isin(TARGET_GRADES)
    if not mask.any():
        return pd.DataFrame()

    sub = df.loc[mask]

    # 担当フィルタ：「0」「-」「」を除外
    teacher = sub.iloc[:, idx["teacher"]].fillna("").astype(str).str.strip()
    teacher_mask = ~teacher.isin(["0", "-", ""])
    if not teacher_mask.any():
        return pd.DataFrame()
    sub = sub[teacher_mask]

    # 講座名解決（ベクトル化）
    course = sub.iloc[:, idx["course"]].astype(str).str.strip()
    class_type = sub.iloc[:, idx["class_type"]]
    class_str = class_type.fillna("").astype(str).str.strip()
    needs_suffix = course.str.contains("ｱﾄﾞﾊﾞﾝｽ|ﾊｲﾚﾍﾞﾙ", na=False) & class_str.ne("")
    resolved_course = course.where(~needs_suffix, course + class_str)

    # グループ化用 DataFrame を一括構築
    group_df = pd.DataFrame({
        "学年": sub.iloc[:, idx["grade"]].map(GRADE_LABELS),
        "教室": sub.iloc[:, idx["classroom"]],
        "講座名": resolved_course,
        "M/C": class_str.values,
        "担当": teacher.loc[teacher_mask],
    })

    result = group_df.groupby(KEY_COLS, dropna=False).size().reset_index(name=month_label)
    return result


def save_monthly_result(df: pd.DataFrame, target_month: pd.Period,
                        results_dir: Path = RESULTS_DIR) -> None:
    """1ヶ月分の集計結果を CSV に保存"""
    results_dir.mkdir(parents=True, exist_ok=True)
    df.to_csv(results_dir / f"{target_month}.csv", index=False, encoding="utf-8-sig")


def build_pivot(results_dir: Path = RESULTS_DIR) -> pd.DataFrame:
    """保存済みの全月CSVを読み込み、concat + groupby でピボット生成"""
    files = list(results_dir.glob("*.csv"))
    if not files:
        return pd.DataFrame()

    frames = []
    for f in files:
        mdf = pd.read_csv(f, dtype=str)
        month_cols = [c for c in mdf.columns if c not in KEY_COLS]
        if not month_cols:
            continue
        col = month_cols[0]
        mdf[col] = pd.to_numeric(mdf[col], errors="coerce").fillna(0).astype(int)
        frames.append(mdf)

    if not frames:
        return pd.DataFrame()

    merged = pd.concat(frames, ignore_index=True)

    # 各月列を集約（同一キーが複数CSVに跨る場合の安全策）
    available = [c for c in MONTH_ORDER if c in merged.columns]
    for col in available:
        merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0).astype(int)

    result = merged.groupby(KEY_COLS, dropna=False)[available].sum().reset_index()
    result = result[KEY_COLS + available]
    return result


def to_excel_bytes(result: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="月次受講人数")
    return buf.getvalue()
