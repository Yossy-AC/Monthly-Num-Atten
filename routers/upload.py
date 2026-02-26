from fastapi import APIRouter, Request, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import io

from services.aggregator import (
    load_excel, aggregate, to_excel_bytes,
    parse_target_month, save_monthly_result, build_pivot,
)

router = APIRouter()
templates = Jinja2Templates(directory="templates")

# 直近のピボット結果をメモリに保持（シングルユーザー前提）
_last_result = None


@router.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
    global _last_result
    content = await file.read()
    df = load_excel(content)
    target_month = parse_target_month(file.filename or "")

    # 対象月1ヶ月分を集計して保存
    monthly = aggregate(df, target_month=target_month)
    if target_month is not None and not monthly.empty:
        save_monthly_result(monthly, target_month)

    # 保存済み全月分からピボットを構築
    _last_result = build_pivot()
    if _last_result.empty:
        _last_result = monthly  # 初回アップロード時のフォールバック

    rows = _last_result.to_dict(orient="records")
    return templates.TemplateResponse(
        "result.html", {"request": request, "rows": rows}
    )


@router.get("/download")
async def download():
    if _last_result is None:
        return HTMLResponse("集計データなし", status_code=400)
    data = to_excel_bytes(_last_result)
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=monthly_stats.xlsx"},
    )
