import copy
import os
from pathlib import Path

from fastapi import Body, FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
import io

from . import claude_client, file_parsing
from .excel_export import build_workbook
from .schema import EMPTY_PAYLOAD

MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB

SPREADSHEET_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv", ".tsv"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
SUPPORTED_EXTENSIONS = SPREADSHEET_EXTENSIONS | IMAGE_EXTENSIONS | {".pdf"}

app = FastAPI(title="Drawing to BOM & Cost Estimator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _normalize_payload(payload: dict) -> dict:
    merged = copy.deepcopy(EMPTY_PAYLOAD)
    if not isinstance(payload, dict):
        return merged

    project = payload.get("project")
    if isinstance(project, dict):
        merged["project"].update(project)

    for key in ("items", "bom", "hardware", "assumptions"):
        value = payload.get(key)
        if isinstance(value, list):
            merged[key] = value

    return merged


@app.get("/api/health")
def health():
    return {"ok": True, "aiConfigured": bool(os.environ.get("ANTHROPIC_API_KEY"))}


@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...)):
    filename = file.filename or "upload"
    ext = Path(filename).suffix.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file type '{ext}'. Upload a PDF, image (PNG/JPG/WEBP), or Excel/CSV file.",
        )

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=400, detail="File is too large (max 20 MB).")

    extracted_text = None
    if ext in SPREADSHEET_EXTENSIONS:
        try:
            extracted_text = file_parsing.extract_text(filename, data)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Could not read spreadsheet: {exc}")

    try:
        payload = claude_client.analyze_drawing(
            filename=filename,
            content_type=file.content_type or "",
            data=data,
            extracted_text=extracted_text,
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc))
    except ValueError as exc:
        raise HTTPException(status_code=502, detail=f"AI analysis failed: {exc}")

    return {"ok": True, "data": _normalize_payload(payload)}


@app.post("/api/export/xlsx")
def export_xlsx(payload: dict = Body(...)):
    normalized = _normalize_payload(payload)
    try:
        workbook_bytes = build_workbook(normalized)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not build workbook: {exc}")

    project_name = normalized["project"].get("projectName") or "BOM_Cost_Estimation"
    safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in project_name).strip() or "BOM_Cost_Estimation"
    filename = f"{safe_name}.xlsx"

    return StreamingResponse(
        io.BytesIO(workbook_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


FRONTEND_DIR = Path(__file__).resolve().parent.parent.parent / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
