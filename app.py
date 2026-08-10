import os
import json
import sqlite3
import threading
import datetime
from zoneinfo import ZoneInfo

import requests
import msal
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from apscheduler.schedulers.background import BackgroundScheduler

# ----------------------------------------------------------------------------
# Config (ตั้งค่าผ่าน .env)
# ----------------------------------------------------------------------------
TENANT_ID = os.getenv("TENANT_ID", "")
CLIENT_ID = os.getenv("CLIENT_ID", "")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Tasks.ReadWrite", "Group.Read.All", "User.Read"]
GRAPH = "https://graph.microsoft.com/v1.0"

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5")
ANTHROPIC_VERSION = "2023-06-01"

TZ = os.getenv("TZ", "Asia/Bangkok")
BRIEF_HOUR = int(os.getenv("BRIEF_HOUR", "7"))
BRIEF_MINUTE = int(os.getenv("BRIEF_MINUTE", "30"))
DUE_SOON_DAYS = int(os.getenv("DUE_SOON_DAYS", "3"))

DATA_DIR = os.getenv("DATA_DIR", "data")
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "itdaily.db")
CACHE_PATH = os.path.join(DATA_DIR, "token_cache.json")
HERE = os.path.dirname(os.path.abspath(__file__))


def now_tz():
    return datetime.datetime.now(ZoneInfo(TZ))


# ----------------------------------------------------------------------------
# MSAL token cache (delegated / device-code flow)
# ----------------------------------------------------------------------------
def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())
    return cache


def save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(cache.serialize())


def build_msal(cache):
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)


def get_token():
    """คืน access token จาก refresh token ที่ cache ไว้ (silent). ถ้ายังไม่ auth คืน None"""
    cache = load_cache()
    app_ = build_msal(cache)
    accounts = app_.get_accounts()
    if not accounts:
        return None
    result = app_.acquire_token_silent(SCOPES, account=accounts[0])
    save_cache(cache)
    if result and "access_token" in result:
        return result["access_token"]
    return None


def start_device_flow():
    """เริ่ม device-code flow, คืน user_code ให้ผู้ใช้เอาไปกรอกที่ verification_uri
    การ auth จริงทำใน background thread แล้วเซฟ token cache เมื่อสำเร็จ"""
    cache = load_cache()
    app_ = build_msal(cache)
    flow = app_.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(json.dumps(flow, ensure_ascii=False))

    def _complete():
        app_.acquire_token_by_device_flow(flow)  # block จนกว่าผู้ใช้จะยืนยัน
        save_cache(cache)

    threading.Thread(target=_complete, daemon=True).start()
    return {
        "user_code": flow["user_code"],
        "verification_uri": flow["verification_uri"],
        "expires_in": flow.get("expires_in"),
    }


# ----------------------------------------------------------------------------
# Microsoft Graph — ดึง Planner tasks ของฉัน
# ----------------------------------------------------------------------------
_plan_cache = {}
_bucket_cache = {}


def graph_get(token, url):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_my_tasks(token):
    tasks, url = [], f"{GRAPH}/me/planner/tasks"
    while url:
        data = graph_get(token, url)
        tasks.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return tasks


def plan_name(token, pid):
    if not pid:
        return ""
    if pid not in _plan_cache:
        try:
            _plan_cache[pid] = graph_get(token, f"{GRAPH}/planner/plans/{pid}").get("title", "")
        except Exception:
            _plan_cache[pid] = ""
    return _plan_cache[pid]


def bucket_name(token, bid):
    if not bid:
        return ""
    if bid not in _bucket_cache:
        try:
            _bucket_cache[bid] = graph_get(token, f"{GRAPH}/planner/buckets/{bid}").get("name", "")
        except Exception:
            _bucket_cache[bid] = ""
    return _bucket_cache[bid]


_me_id = None


def get_me_id(token):
    global _me_id
    if not _me_id:
        _me_id = graph_get(token, f"{GRAPH}/me").get("id")
    return _me_id


def list_plans(token):
    data = graph_get(token, f"{GRAPH}/me/planner/plans")
    return [{"id": p.get("id"), "title": p.get("title", "")} for p in data.get("value", [])]


def list_buckets(token, plan_id):
    data = graph_get(token, f"{GRAPH}/planner/plans/{plan_id}/buckets")
    return [{"id": b.get("id"), "name": b.get("name", "")} for b in data.get("value", [])]


def create_task(token, plan_id, bucket_id, title, due, assign_me):
    body = {"planId": plan_id, "title": title}
    if bucket_id:
        body["bucketId"] = bucket_id
    if due:
        body["dueDateTime"] = f"{due}T17:00:00Z"
    if assign_me:
        uid = get_me_id(token)
        body["assignments"] = {
            uid: {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}
        }
    r = requests.post(
        f"{GRAPH}/planner/tasks",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
        timeout=30,
    )
    r.raise_for_status()
    _plan_cache.clear()
    _bucket_cache.clear()
    return r.json()


def update_task(token, task_id, fields):
    # Planner ต้องส่ง If-Match ด้วย etag ล่าสุด (optimistic concurrency)
    cur = graph_get(token, f"{GRAPH}/planner/tasks/{task_id}")
    etag = cur.get("@odata.etag")
    r = requests.patch(
        f"{GRAPH}/planner/tasks/{task_id}",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "If-Match": etag,
        },
        json=fields,
        timeout=30,
    )
    r.raise_for_status()
    return True


def build_task_view(token):
    raw = fetch_my_tasks(token)
    today = now_tz().date()
    items = []
    for t in raw:
        pct = t.get("percentComplete", 0)
        due = None
        if t.get("dueDateTime"):
            try:
                due = (
                    datetime.datetime.fromisoformat(t["dueDateTime"].replace("Z", "+00:00"))
                    .astimezone(ZoneInfo(TZ))
                    .date()
                )
            except Exception:
                due = None
        completed = None
        if t.get("completedDateTime"):
            try:
                completed = (
                    datetime.datetime.fromisoformat(t["completedDateTime"].replace("Z", "+00:00"))
                    .astimezone(ZoneInfo(TZ))
                    .date()
                )
            except Exception:
                completed = None
        item = {
            "id": t.get("id"),
            "title": t.get("title", ""),
            "percent": pct,
            "due": due.isoformat() if due else None,
            "completed": completed.isoformat() if completed else None,
            "plan": plan_name(token, t.get("planId")),
            "bucket": bucket_name(token, t.get("bucketId")),
            "priority": t.get("priority", 5),
        }
        if pct >= 100:
            item["status"] = "done"
        elif due and due < today:
            item["status"] = "overdue"
        elif due and due == today:
            item["status"] = "today"
        elif due and (due - today).days <= DUE_SOON_DAYS:
            item["status"] = "soon"
        elif pct > 0:
            item["status"] = "inprogress"
        else:
            item["status"] = "backlog"
        items.append(item)
    items.sort(key=lambda x: (x["due"] or "9999-99-99", x["priority"]))
    return items


# ----------------------------------------------------------------------------
# AI — สรุปเคสรายวันด้วย Claude
# ----------------------------------------------------------------------------
def summarize(items):
    if not ANTHROPIC_API_KEY:
        return None
    active = [i for i in items if i["status"] != "done"]
    lines = [
        f"- [{i['status']}] {i['title']} | plan: {i['plan'] or '-'} | due: {i['due'] or '-'} | {i['percent']}%"
        for i in active
    ]
    task_text = "\n".join(lines) if lines else "(ไม่มีงานค้าง)"
    system = (
        "คุณเป็นผู้ช่วยของ IT admin ทำหน้าที่สรุปงานรายวันจาก Microsoft Planner "
        "ตอบเป็นภาษาไทย กระชับ ตรงประเด็น เก็บศัพท์เทคนิคภาษาอังกฤษไว้ตามเดิม"
    )
    user = (
        f"นี่คือ task ที่ถูก assign ให้ฉันวันนี้ ({now_tz().date().isoformat()}):\n"
        f"{task_text}\n\n"
        "ช่วยสรุปและวางแผนวันนี้ ตอบกลับเป็น JSON เท่านั้น ห้ามมี markdown หรือข้อความอื่น รูปแบบ:\n"
        '{"headline":"สรุปหนึ่งบรรทัด","summary":"2-4 ประโยค",'
        '"today_focus":[{"title":"ชื่องาน","why":"เหตุผลสั้นๆ","plan":"ชื่อ plan"}],'
        '"risks":["งานเสี่ยง/เลย deadline"]}\n'
        "เลือก today_focus 3-5 อย่างที่ควรทำก่อน โดยดูจาก overdue > today > soon และ priority"
    )
    body = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 1500,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    r = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": ANTHROPIC_VERSION,
            "content-type": "application/json",
        },
        json=body,
        timeout=60,
    )
    r.raise_for_status()
    data = r.json()
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    text = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    try:
        return json.loads(text)
    except Exception:
        return {"headline": "", "summary": text, "today_focus": [], "risks": []}


# ----------------------------------------------------------------------------
# Storage — เก็บ brief รายวันไว้ทำ history / audit
# ----------------------------------------------------------------------------
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS briefs (date TEXT PRIMARY KEY, payload TEXT, created_at TEXT)"
    )
    return conn


def save_brief(date_str, payload):
    conn = db()
    conn.execute(
        "INSERT OR REPLACE INTO briefs(date,payload,created_at) VALUES(?,?,?)",
        (date_str, json.dumps(payload, ensure_ascii=False), now_tz().isoformat()),
    )
    conn.commit()
    conn.close()


def get_brief(date_str):
    conn = db()
    row = conn.execute("SELECT payload FROM briefs WHERE date=?", (date_str,)).fetchone()
    conn.close()
    return json.loads(row[0]) if row else None


def generate_daily():
    token = get_token()
    if not token:
        return {"connected": False}
    items = build_task_view(token)
    date_str = now_tz().date().isoformat()
    payload = {
        "connected": True,
        "date": date_str,
        "generated_at": now_tz().isoformat(),
        "tasks": items,
        "ai": summarize(items),
    }
    save_brief(date_str, payload)
    return payload


# ----------------------------------------------------------------------------
# FastAPI
# ----------------------------------------------------------------------------
app = FastAPI(title="IT Daily")


@app.get("/", response_class=HTMLResponse)
def index():
    with open(os.path.join(HERE, "index.html"), encoding="utf-8") as f:
        return f.read()


@app.get("/api/status")
def status():
    return {
        "connected": get_token() is not None,
        "config_ok": bool(TENANT_ID and CLIENT_ID),
        "ai_enabled": bool(ANTHROPIC_API_KEY),
        "schedule": f"{BRIEF_HOUR:02d}:{BRIEF_MINUTE:02d} (Mon-Fri, {TZ})",
    }


@app.post("/auth/start")
def auth_start():
    try:
        return start_device_flow()
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)


@app.get("/api/data")
def api_data():
    date_str = now_tz().date().isoformat()
    return get_brief(date_str) or generate_daily()


@app.post("/api/refresh")
def api_refresh():
    return generate_daily()


def _require_token():
    token = get_token()
    if not token:
        return None, JSONResponse({"error": "ยังไม่ได้เชื่อม Microsoft"}, status_code=401)
    return token, None


@app.get("/api/plans")
def api_plans():
    token, err = _require_token()
    if err:
        return err
    try:
        return list_plans(token)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)


@app.get("/api/buckets")
def api_buckets(plan_id: str):
    token, err = _require_token()
    if err:
        return err
    try:
        return list_buckets(token, plan_id)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)


@app.post("/api/tasks")
async def api_create_task(req: Request):
    token, err = _require_token()
    if err:
        return err
    b = await req.json()
    if not b.get("title") or not b.get("plan_id"):
        return JSONResponse({"error": "ต้องมี title และ plan_id"}, status_code=400)
    try:
        task = create_task(
            token,
            b["plan_id"],
            b.get("bucket_id") or None,
            b["title"],
            b.get("due") or None,
            bool(b.get("assign_me", True)),
        )
        return {"ok": True, "id": task.get("id")}
    except requests.HTTPError as e:
        return JSONResponse({"error": _graph_err(e)}, status_code=400)


@app.patch("/api/tasks/{task_id}")
async def api_update_task(task_id: str, req: Request):
    token, err = _require_token()
    if err:
        return err
    b = await req.json()
    fields = {}
    if "percent" in b:
        fields["percentComplete"] = int(b["percent"])
    if "title" in b:
        fields["title"] = b["title"]
    if "due" in b:
        fields["dueDateTime"] = (f"{b['due']}T17:00:00Z" if b["due"] else None)
    if not fields:
        return JSONResponse({"error": "ไม่มีข้อมูลให้อัปเดต"}, status_code=400)
    try:
        update_task(token, task_id, fields)
        return {"ok": True}
    except requests.HTTPError as e:
        return JSONResponse({"error": _graph_err(e)}, status_code=400)


def _graph_err(e):
    try:
        return e.response.json().get("error", {}).get("message", str(e))
    except Exception:
        return str(e)


scheduler = BackgroundScheduler(timezone=TZ)


@app.on_event("startup")
def _startup():
    scheduler.add_job(
        generate_daily,
        "cron",
        day_of_week="mon-fri",
        hour=BRIEF_HOUR,
        minute=BRIEF_MINUTE,
        id="daily_brief",
        replace_existing=True,
    )
    scheduler.start()


@app.on_event("shutdown")
def _shutdown():
    scheduler.shutdown(wait=False)
