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
├── services/aggregator.py  # 集計コアロジック（COLUMN_MAPで列名管理）
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

## Column Mapping (要更新)
サンプルExcel受領後、`services/aggregator.py` の `COLUMN_MAP` を実際の列名に変更する。

```python
COLUMN_MAP = {
    "student_id": "生徒コード",   # 要確認
    "course":     "講座名",       # 要確認
    "event":      "イベント種別", # 要確認
    "date":       "処理日",       # 要確認
    "grade":      "学年",         # 要確認
    "classroom":  "教室",         # 要確認
    "teacher":    "担当",         # 要確認
}
```

## Event Classification
- 入塾・入講 → 受講者セットに追加
- 退塾・退講 → 受講者セットから除外
- クラス変更・受講教室変更 → 受講継続（セット変更なし）
