from fastapi import APIRouter, Request, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import io

from services.aggregator import load_excel, aggregate, to_excel_bytes

router = APIRouter()
templates = Jinja2Templates(directory="templates")

# 直近の集計結果をメモリに保持（シングルユーザー前提）
_last_result = None


@router.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
    global _last_result
    content = await file.read()
    df = load_excel(content)
    _last_result = aggregate(df)

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
