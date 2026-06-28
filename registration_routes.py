# Registration addon routes - add to your existing FastAPI app with:
# from registration_routes import router as registration_router
# app.include_router(registration_router)

import json
import os
import re
import shutil
import uuid
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import APIRouter, Form, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, StreamingResponse
from starlette.datastructures import UploadFile as StarletteUploadFile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("REGISTRATION_DATA_DIR", APP_DIR / "registration_data"))
UPLOAD_DIR = DATA_DIR / "uploads"
DATA_FILE = DATA_DIR / "registrations.json"
ADMIN_PASSWORD = os.getenv("REGISTRATION_ADMIN_PASSWORD", "BeshooWarZone")
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "8"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024

DATA_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

router = APIRouter()

ALLOWED_IMAGE_TYPES = {"image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp"}
REQUIRED_PLAYER_FIELDS = ["name", "age", "birthdate", "national_id", "university", "college", "gender"]
FILE_FIELDS = ["photo", "id_card", "university_card"]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_data() -> Dict[str, Any]:
    if not DATA_FILE.exists():
        return {"teams": []}
    try:
        with DATA_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if "teams" not in data or not isinstance(data["teams"], list):
            return {"teams": []}
        return data
    except Exception:
        return {"teams": []}


def save_data(data: Dict[str, Any]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = DATA_FILE.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(DATA_FILE)


def normalize_team_name(name: str) -> str:
    name = re.sub(r"\s+", " ", (name or "").strip())
    return name.casefold()


def slugify(value: str) -> str:
    value = re.sub(r"[^\w\-]+", "_", value.strip(), flags=re.UNICODE)
    return value[:80] or "file"


def require_admin(request: Request) -> None:
    supplied = request.headers.get("x-admin-password") or request.query_params.get("p") or request.query_params.get("password")
    if supplied != ADMIN_PASSWORD:
        raise HTTPException(status_code=401, detail="Unauthorized")


def public_team(team: Dict[str, Any], request: Optional[Request] = None, include_files: bool = True) -> Dict[str, Any]:
    item = json.loads(json.dumps(team, ensure_ascii=False))
    if include_files and request is not None:
        for player in item.get("players", []):
            files = player.get("files", {})
            player["file_urls"] = {}
            for kind in FILE_FIELDS:
                if files.get(kind):
                    player["file_urls"][kind] = str(request.base_url).rstrip("/") + f"/api/registration-file/{item['id']}/{player['id']}/{kind}"
    return item


def validate_players(players: List[Dict[str, Any]], old_team: Optional[Dict[str, Any]] = None) -> None:
    if not (12 <= len(players) <= 15):
        raise HTTPException(status_code=400, detail="كل فريق لازم يكون من 12 إلى 15 فرد.")

    male_count = 0
    female_count = 0
    seen_national_ids = set()

    for idx, player in enumerate(players, start=1):
        for field in REQUIRED_PLAYER_FIELDS:
            if not str(player.get(field, "")).strip():
                raise HTTPException(status_code=400, detail=f"بيانات اللاعب رقم {idx} ناقصة: {field}")

        try:
            age = int(player.get("age"))
            if age <= 0 or age > 100:
                raise ValueError
            player["age"] = age
        except Exception:
            raise HTTPException(status_code=400, detail=f"سن اللاعب رقم {idx} غير صحيح.")

        national_id = re.sub(r"\D+", "", str(player.get("national_id", "")))
        if len(national_id) != 14:
            raise HTTPException(status_code=400, detail=f"الرقم القومي للاعب رقم {idx} لازم يكون 14 رقم.")
        if national_id in seen_national_ids:
            raise HTTPException(status_code=400, detail=f"الرقم القومي مكرر داخل نفس الفريق عند اللاعب رقم {idx}.")
        seen_national_ids.add(national_id)
        player["national_id"] = national_id

        gender = str(player.get("gender", "")).strip().lower()
        if gender in ["male", "ذكر", "m"]:
            player["gender"] = "ذكر"
            male_count += 1
        elif gender in ["female", "انثى", "أنثى", "f"]:
            player["gender"] = "أنثى"
            female_count += 1
        else:
            raise HTTPException(status_code=400, detail=f"نوع اللاعب رقم {idx} لازم يكون ذكر أو أنثى.")

        # Basic date format check
        try:
            datetime.strptime(str(player.get("birthdate")), "%Y-%m-%d")
        except Exception:
            raise HTTPException(status_code=400, detail=f"تاريخ ميلاد اللاعب رقم {idx} غير صحيح.")

    if male_count < 6 or female_count < 6:
        raise HTTPException(status_code=400, detail="كل فريق لازم يكون فيه على الأقل 6 ذكور و6 إناث.")


def ensure_team_name_unique(data: Dict[str, Any], team_name: str, exclude_team_id: Optional[str] = None) -> None:
    normalized = normalize_team_name(team_name)
    if not normalized:
        raise HTTPException(status_code=400, detail="اسم المنتخب مطلوب.")
    for team in data.get("teams", []):
        if exclude_team_id and team.get("id") == exclude_team_id:
            continue
        if normalize_team_name(team.get("team_name", "")) == normalized:
            raise HTTPException(status_code=409, detail="اسم المنتخب مستخدم قبل كده، اختار اسم تاني.")


async def save_uploaded_file(upload: StarletteUploadFile, team_id: str, player_id: str, kind: str) -> str:
    if not upload or not getattr(upload, "filename", ""):
        raise HTTPException(status_code=400, detail=f"ملف {kind} مطلوب.")

    content_type = upload.content_type or ""
    ext = ALLOWED_IMAGE_TYPES.get(content_type)
    if not ext:
        raise HTTPException(status_code=400, detail="الصور المسموحة: JPG / PNG / WEBP فقط.")

    data = await upload.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=400, detail=f"حجم الصورة لا يزيد عن {MAX_UPLOAD_MB}MB.")

    folder = UPLOAD_DIR / team_id / player_id
    folder.mkdir(parents=True, exist_ok=True)
    filename = f"{kind}_{uuid.uuid4().hex}{ext}"
    path = folder / filename
    with path.open("wb") as f:
        f.write(data)
    return str(path.relative_to(DATA_DIR))


def delete_team_files(team_id: str) -> None:
    folder = UPLOAD_DIR / team_id
    if folder.exists():
        shutil.rmtree(folder, ignore_errors=True)


async def build_team_from_form(request: Request, existing: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    form = await request.form()
    team_name = str(form.get("team_name", "")).strip()
    players_json = str(form.get("players_json", "[]"))

    try:
        players_raw = json.loads(players_json)
    except Exception:
        raise HTTPException(status_code=400, detail="صيغة بيانات اللاعبين غير صحيحة.")

    if not isinstance(players_raw, list):
        raise HTTPException(status_code=400, detail="بيانات اللاعبين لازم تكون قائمة.")

    team_id = existing.get("id") if existing else uuid.uuid4().hex
    old_players_by_id = {p.get("id"): p for p in (existing or {}).get("players", [])}
    players: List[Dict[str, Any]] = []

    for idx, raw in enumerate(players_raw):
        player_id = raw.get("id") or uuid.uuid4().hex
        client_key = str(raw.get("client_key") or raw.get("id") or idx)
        old_player = old_players_by_id.get(player_id, {})
        old_files = old_player.get("files", {})

        player = {
            "id": player_id,
            "name": str(raw.get("name", "")).strip(),
            "age": raw.get("age", ""),
            "birthdate": str(raw.get("birthdate", "")).strip(),
            "national_id": str(raw.get("national_id", "")).strip(),
            "university": str(raw.get("university", "")).strip(),
            "college": str(raw.get("college", "")).strip(),
            "gender": str(raw.get("gender", "")).strip(),
            "files": dict(old_files),
        }

        for kind in FILE_FIELDS:
            upload = form.get(f"{kind}_{client_key}")
            if isinstance(upload, StarletteUploadFile) and upload.filename:
                player["files"][kind] = await save_uploaded_file(upload, team_id, player_id, kind)
            elif not player["files"].get(kind):
                raise HTTPException(status_code=400, detail=f"صورة {kind} مطلوبة للاعب رقم {idx + 1}.")

        players.append(player)

    validate_players(players, old_team=existing)

    return {
        "id": team_id,
        "team_name": team_name,
        "created_at": (existing or {}).get("created_at") or now_iso(),
        "updated_at": now_iso(),
        "players": players,
    }



@router.get("/api/team-name-available")
def team_name_available(name: str):
    data = load_data()
    normalized = normalize_team_name(name)
    if not normalized:
        return {"available": False, "message": "اكتب اسم المنتخب"}
    for team in data.get("teams", []):
        if normalize_team_name(team.get("team_name", "")) == normalized:
            return {"available": False, "message": "الاسم مستخدم قبل كده"}
    return {"available": True, "message": "الاسم متاح"}


@router.post("/api/register-team")
async def register_team(request: Request):
    data = load_data()
    team = await build_team_from_form(request)
    ensure_team_name_unique(data, team["team_name"])

    # Check national ID uniqueness across teams.
    existing_ids = {p.get("national_id") for t in data.get("teams", []) for p in t.get("players", [])}
    for p in team["players"]:
        if p.get("national_id") in existing_ids:
            delete_team_files(team["id"])
            raise HTTPException(status_code=409, detail=f"الرقم القومي {p.get('national_id')} مسجل قبل كده في فريق آخر.")

    data["teams"].append(team)
    save_data(data)
    return {"status": "success", "team_id": team["id"], "team_name": team["team_name"]}


@router.get("/api/registrations")
def list_registrations(request: Request):
    require_admin(request)
    data = load_data()
    teams = []
    for team in data.get("teams", []):
        males = sum(1 for p in team.get("players", []) if p.get("gender") == "ذكر")
        females = sum(1 for p in team.get("players", []) if p.get("gender") == "أنثى")
        teams.append({
            "id": team.get("id"),
            "team_name": team.get("team_name"),
            "created_at": team.get("created_at"),
            "updated_at": team.get("updated_at"),
            "players_count": len(team.get("players", [])),
            "males": males,
            "females": females,
        })
    return {"teams": teams}


@router.get("/api/registrations/{team_id}")
def get_registration(team_id: str, request: Request):
    require_admin(request)
    data = load_data()
    for team in data.get("teams", []):
        if team.get("id") == team_id:
            return public_team(team, request=request)
    raise HTTPException(status_code=404, detail="الفريق غير موجود.")


@router.put("/api/registrations/{team_id}")
async def update_registration(team_id: str, request: Request):
    require_admin(request)
    data = load_data()
    for idx, old_team in enumerate(data.get("teams", [])):
        if old_team.get("id") == team_id:
            updated = await build_team_from_form(request, existing=old_team)
            ensure_team_name_unique(data, updated["team_name"], exclude_team_id=team_id)

            # Check national IDs across other teams.
            other_ids = {p.get("national_id") for t in data.get("teams", []) if t.get("id") != team_id for p in t.get("players", [])}
            for p in updated["players"]:
                if p.get("national_id") in other_ids:
                    raise HTTPException(status_code=409, detail=f"الرقم القومي {p.get('national_id')} مسجل في فريق آخر.")

            data["teams"][idx] = updated
            save_data(data)
            return {"status": "success", "team_id": team_id}
    raise HTTPException(status_code=404, detail="الفريق غير موجود.")


@router.delete("/api/registrations/{team_id}")
def delete_registration(team_id: str, request: Request):
    require_admin(request)
    data = load_data()
    before = len(data.get("teams", []))
    data["teams"] = [t for t in data.get("teams", []) if t.get("id") != team_id]
    if len(data["teams"]) == before:
        raise HTTPException(status_code=404, detail="الفريق غير موجود.")
    delete_team_files(team_id)
    save_data(data)
    return {"status": "success"}


@router.get("/api/registration-file/{team_id}/{player_id}/{kind}")
def get_registration_file(team_id: str, player_id: str, kind: str, request: Request):
    require_admin(request)
    if kind not in FILE_FIELDS:
        raise HTTPException(status_code=404, detail="نوع الملف غير صحيح.")
    data = load_data()
    for team in data.get("teams", []):
        if team.get("id") == team_id:
            for player in team.get("players", []):
                if player.get("id") == player_id:
                    rel = player.get("files", {}).get(kind)
                    if not rel:
                        raise HTTPException(status_code=404, detail="الملف غير موجود.")
                    path = DATA_DIR / rel
                    if not path.exists():
                        raise HTTPException(status_code=404, detail="الملف غير موجود على السيرفر.")
                    return FileResponse(path)
    raise HTTPException(status_code=404, detail="الملف غير موجود.")


@router.get("/api/registrations/export")
def export_registrations(request: Request):
    require_admin(request)
    data = load_data()
    wb = Workbook()
    ws = wb.active
    ws.title = "Teams Registrations"
    headers = [
        "اسم المنتخب", "تاريخ التسجيل", "اسم اللاعب", "السن", "تاريخ الميلاد", "الرقم القومي",
        "الجامعة", "الكلية", "النوع", "الصورة الشخصية", "صورة البطاقة", "صورة كارنيه الجامعة"
    ]
    ws.append(headers)

    base = str(request.base_url).rstrip("/")
    password = request.headers.get("x-admin-password") or request.query_params.get("p") or request.query_params.get("password") or ""

    for team in data.get("teams", []):
        for player in team.get("players", []):
            urls = {}
            for kind in FILE_FIELDS:
                if player.get("files", {}).get(kind):
                    urls[kind] = f"{base}/api/registration-file/{team['id']}/{player['id']}/{kind}?p={password}"
                else:
                    urls[kind] = ""
            ws.append([
                team.get("team_name", ""),
                team.get("created_at", ""),
                player.get("name", ""),
                player.get("age", ""),
                player.get("birthdate", ""),
                player.get("national_id", ""),
                player.get("university", ""),
                player.get("college", ""),
                player.get("gender", ""),
                urls["photo"],
                urls["id_card"],
                urls["university_card"],
            ])

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1E3A8A")
        cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        max_len = 12
        col_letter = col[0].column_letter
        for cell in col:
            max_len = max(max_len, len(str(cell.value or ""))[:80] if False else min(len(str(cell.value or "")), 80))
        ws.column_dimensions[col_letter].width = max_len + 4

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"warzone_registrations_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
