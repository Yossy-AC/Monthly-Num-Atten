Role: Dedicated engineer and assistant for a university prep school English teacher.
Style: Conclusion first, concise, direct, no token waste.
Prohibited: greetings, prefaces, apologies, emojis/kaomoji.

---

## Project Overview
月次受講人数集計Webアプリ。本部から送付されるExcel（生徒受講情報）をアップロードし、
講座別・月別の受講人数を集計・Excel出力する。

## Tech Stack
- Backend: FastAPI + uvicorn
- Frontend: HTMX + Jinja2
- Package manager: uv
- Excel処理: pandas + openpyxl

## Structure
```
Course-Stats-Analyzer/
├── main.py                 # FastAPIアプリ本体
├── routers/upload.py       # /upload (POST), /download (GET)
├── services/aggregator.py  # 集計コアロジック（COLUMN_INDICESで列位置管理）
├── templates/              # base.html, index.html, result.html
├── static/style.css
├── uploads/                # アップロードExcel一時保存
└── outputs/                # 集計済みExcel出力先
```

## Run
```
uv run uvicorn main:app --reload
```
→ http://127.0.0.1:8000

## Column Mapping (確定版)
Excelファイルの列構造（Row 4がヘッダー、0-indexed）

```python
COLUMN_INDICES = {
    "add_date": 2,      # Column C: 受講追加日付
    "cancel_date": 6,   # Column G: 受講取消日付
    "course": 9,        # Column J: 講座名
    "class_type": 10,   # Column K: 【マスター】【コア】
    "classroom": 11,    # Column L: 受講教室
    "grade": 15,        # Column P: 学年コード (31=高1, 32=高2, 33=高3)
    "teacher": 26,      # Column AA: 担当
    "gender": 17,       # Column R: 性別コード (1=男, 2=女)
    "school": 18,       # Column S: 在籍校
    "department": 28,   # Column AC: 学科
}
```

## 集計仕様
- **集計対象**: 高1/高2/高3のみ（学年コード31/32/33）
- **月末判定**: `add_date <= cutoff_date AND (cancel_date IS NULL OR cancel_date > cutoff_date)`
- **出力形式**: Pivot（固定列+月列）
- **固定列順序**: 学年 | 教室 | 講座名 | マスター/コア | 担当 | 在籍校 | 学科 | 性別
- **月列順序**: 4月～3月（会計年度順）
