"""
受講人数集計ロジック

列名はサンプルExcel受領後に COLUMN_MAP を更新すること。
"""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd

# ── 列名マッピング（サンプルExcel受領後に実際の列名に合わせて変更） ──
COLUMN_MAP = {
    "student_id": "生徒コード",
    "course": "講座名",
    "event": "イベント種別",
    "date": "処理日",
    "grade": "学年",
    "classroom": "教室",
    "teacher": "担当",
}

# 受講中とみなすイベント
ENROLL_EVENTS = {"入塾", "入講"}
# 退出とみなすイベント
LEAVE_EVENTS = {"退塾", "退講"}
# 変更系（受講継続）
CHANGE_EVENTS = {"クラス変更", "受講教室変更"}


def load_excel(file: bytes | Path) -> pd.DataFrame:
    if isinstance(file, bytes):
        return pd.read_excel(io.BytesIO(file))
    return pd.read_excel(file)


def aggregate(df: pd.DataFrame) -> pd.DataFrame:
    """
    月末時点の講座別受講人数を集計して返す。

    Returns
    -------
    pd.DataFrame
        列: 講座名, 学年, 教室, 担当, 年月, 受講人数
    """
    col = COLUMN_MAP

    # 日付列を datetime に変換
    df[col["date"]] = pd.to_datetime(df[col["date"]])
    df["年月"] = df[col["date"]].dt.to_period("M")

    # イベント別に分類
    df["_is_enroll"] = df[col["event"]].isin(ENROLL_EVENTS)
    df["_is_leave"] = df[col["event"]].isin(LEAVE_EVENTS)

    # 全講座・全月の組み合わせを取得
    courses = df[col["course"]].unique()
    months = sorted(df["年月"].unique())

    records = []
    for course in courses:
        sub = df[df[col["course"]] == course].sort_values(col["date"])
        # 講座のメタ情報（学年・教室・担当）は最新レコードから取得
        meta = sub.iloc[-1]

        active: set[str] = set()
        month_idx = 0

        for month in months:
            # その月以前のイベントを処理
            month_end = month.to_timestamp("M")
            period_rows = sub[sub[col["date"]] <= month_end]
            for _, row in period_rows.iterrows():
                sid = str(row[col["student_id"]])
                if row["_is_enroll"]:
                    active.add(sid)
                elif row["_is_leave"]:
                    active.discard(sid)

            records.append({
                "講座名": course,
                "学年": meta.get(col["grade"], ""),
                "教室": meta.get(col["classroom"], ""),
                "担当": meta.get(col["teacher"], ""),
                "年月": str(month),
                "受講人数": len(active),
            })

    result = pd.DataFrame(records)
    return result


def to_excel_bytes(result: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="月次受講人数")
    return buf.getvalue()
