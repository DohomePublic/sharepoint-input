#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Web server สำหรับหน้าเพิ่มข้อมูลเข้า SharePoint List 'DemoApp'

เสิร์ฟไฟล์ webapp/index.html และให้ JSON API สำหรับฟอร์ม
ใช้ Python standard library เท่านั้น (+ requests ผ่าน scripts/sp_upload.py)
ไม่ต้องติดตั้ง Flask/Django

    python webapp/server.py                 # พอร์ต 8000
    python webapp/server.py --port 8080
    python webapp/server.py --dry-run       # ไม่เขียนลง SharePoint จริง

เปิดเบราว์เซอร์ที่ http://127.0.0.1:8000

Endpoints
    GET  /             หน้าเว็บ (webapp/index.html)
    GET  /api/fields   นิยามฟิลด์บนฟอร์ม (JSON)
    GET  /api/health   สถานะการเชื่อมต่อ (JSON)
    GET  /api/schema   คอลัมน์ที่เขียนได้ของ list (JSON)
    POST /api/submit   บันทึกข้อมูล {"mode": "...", "fields": {...}} (JSON)
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))
sys.path.insert(0, os.path.join(ROOT_DIR, "scripts"))

try:
    from sp_upload import Config, GraphClient, build_resolver, load_field_map, coerce
except ImportError as exc:  # pragma: no cover
    print(f"[web] นำเข้า scripts/sp_upload.py ไม่ได้: {exc}\n"
          f"[web] ติดตั้ง dependency ก่อน:  pip install -r requirements.txt", flush=True)
    raise

# ---------------------------------------------------------------------------
# นิยามฟิลด์บนฟอร์ม — อิงคอลัมน์จริงของ list DemoApp
# key = ชื่อคอลัมน์ (display name หรือ internal name) ที่ส่งให้ resolver
# ---------------------------------------------------------------------------
FORM_FIELDS = [
    {"key": "Title", "label": "Title (เลขที่/ชื่อคำขอ)", "type": "text",
     "required": True, "placeholder": "Rungroj Chatch Limited Partnership..."},
    {"key": "Registered_Name", "label": "Registered_Name (ชื่อจดทะเบียน)", "type": "text",
     "required": True, "placeholder": "Rungroj Chatchawal Limited Partnership"},
    {"key": "Customer_id", "label": "Customer_id (รหัสลูกค้า)", "type": "text",
     "required": True, "placeholder": "101047587",
     "help": "ใช้เป็น key สำหรับ upsert — ถ้ามีอยู่แล้วจะอัปเดตรายการเดิม"},
    {"key": "Type1", "label": "Type1 (ประเภทลูกค้า)", "type": "select",
     "options": ["", "Existing", "Lead", "ค้าปลีก"]},
    {"key": "Status", "label": "Status (สถานะ)", "type": "select", "required": True,
     "options": ["", "Draft", "รอการพิจารณาเบื้องต้น", "ผ่านการพิจารณาเบื้องต้น",
                 "ไม่ผ่านการพิจารณาเบื้องต้น", "อนุมัติ-KYC", "ไม่อนุมัติ"]},
    {"key": "Type_Request", "label": "Type_Request (ประเภทคำขอ)", "type": "select",
     "options": ["", "คำขอเปิดวงเงินลูกค้าใหม่", "เปิดวงเงินลูกค้าใหม่",
                 "คำขอเพิ่มวงเงิน", "เพิ่มวงเงิน", "C.ติดตามชุดเปิดตัวจริง"]},
    {"key": "business_type", "label": "business_type (ประเภทธุรกิจ)", "type": "select",
     "options": ["", "รับเหมาก่อสร้าง", "รับเหมาก่อสร้าง(เพิ่ม)", "ร้านค้าช่วง",
                 "อสังหาริมทรัพย์", "อสังหาริมทรัพย์(เพิ่ม)", "หน่วยงานราชการ",
                 "อื่นๆ", "อื่นๆ(เพิ่ม)"]},
    {"key": "Customer Name", "label": "Customer Name (ชื่อลูกค้า)", "type": "text"},
    {"key": "type_teams", "label": "type_teams (ทีม)", "type": "select",
     "options": ["", "Store Operation", "Wholesales (WS)", "Project Sales (PS)", "Retail",
                 "Steel Key Account", "Key Account", "ค้าปลีก", "ตะวันตก", "ผู้แทนขาย"]},
    {"key": "branch", "label": "branch (สาขา)", "type": "text"},
    {"key": "Owner", "label": "Owner (ผู้รับผิดชอบ)", "type": "text",
     "placeholder": "name@dohome.co.th"},
    {"key": "registration_number", "label": "registration_number (เลขทะเบียน)", "type": "text"},
    {"key": "Estimated_annual_income", "label": "Estimated_annual_income (รายได้ต่อปี)",
     "type": "select", "options": ["", "10 - 50 ล้านบาท", "มากกว่า 50 ล้านบาท"]},
    {"key": "limit", "label": "limit (วงเงินที่ขอ)", "type": "text", "placeholder": "1000000",
     "help": "คอลัมน์นี้เป็น Text ในลิสต์ — ระบบจะตัดเครื่องหมาย , ให้อัตโนมัติ"},
    {"key": "province", "label": "province (จังหวัด)", "type": "text"},
    {"key": "district", "label": "district (อำเภอ/เขต)", "type": "text"},
    {"key": "telephone", "label": "telephone (โทรศัพท์)", "type": "text"},
    {"key": "contact_name", "label": "contact_name (ผู้ติดต่อ)", "type": "text"},
    {"key": "contact_number", "label": "contact_number (เบอร์ผู้ติดต่อ)", "type": "text"},
]

REQUIRED_KEYS = [f["key"] for f in FORM_FIELDS if f.get("required")]

STATE = {"dry_run": False, "cfg": None, "client": None, "resolver": None,
         "connected": False, "error": None}


# ---------------------------------------------------------------------------
def connect() -> None:
    """เชื่อมต่อ Microsoft Graph และเตรียมตัว map ชื่อคอลัมน์"""
    cfg = Config()
    STATE["cfg"] = cfg
    if STATE["dry_run"] and not (cfg.client_id and cfg.tenant_id and cfg.client_secret):
        STATE["error"] = "ไม่มี credential — ทำงานในโหมด DRY-RUN (ไม่เขียนจริง)"
        STATE["connected"] = False
        return
    try:
        cfg.require_credentials()
        gc = GraphClient(cfg)
        gc.resolve_site_and_list()
        cols = gc.columns()
        fmap = load_field_map(os.path.join(ROOT_DIR, "config", "field_map.json"))
        STATE["client"] = gc
        STATE["resolver"] = build_resolver(cols, fmap)
        STATE["connected"] = True
        STATE["error"] = None
    except SystemExit as exc:
        STATE["connected"] = False
        STATE["error"] = (f"เชื่อมต่อ SharePoint ไม่สำเร็จ (exit {exc.code}) "
                          f"— ดูรายละเอียดใน console")
    except Exception as exc:
        STATE["connected"] = False
        STATE["error"] = f"{type(exc).__name__}: {exc}"


def build_payload(fields: dict):
    """แปลงค่าจากฟอร์ม -> payload ที่ส่งให้ Graph (ชื่อ internal + ค่า normalize แล้ว)"""
    payload, ignored = {}, []
    for f in FORM_FIELDS:
        raw = fields.get(f["key"], "")
        value = coerce(raw)
        if value is None:
            continue
        name = f.get("sp_key", f["key"])
        if STATE["resolver"]:
            target = STATE["resolver"](name)
            if not target:
                ignored.append(name)
                continue
        else:
            target = name
        payload[target] = value
    return payload, ignored


# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "DemoAppUploader/1.1"

    # -- helpers ----------------------------------------------------------- #
    def _send(self, data: bytes, ctype: str, code: int = 200) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(data)

    def _json(self, obj, code: int = 200) -> None:
        self._send(json.dumps(obj, ensure_ascii=False, indent=2).encode("utf-8"),
                   "application/json; charset=utf-8", code)

    def _file(self, name: str) -> None:
        path = os.path.join(BASE_DIR, name)
        if not os.path.isfile(path):
            return self._send(
                f"<h1>404</h1><p>ไม่พบไฟล์ {name} ใน {BASE_DIR}</p>".encode("utf-8"),
                "text/html; charset=utf-8", 404)
        ctype = ("text/html; charset=utf-8" if name.endswith(".html")
                 else "text/css; charset=utf-8" if name.endswith(".css")
                 else "application/javascript; charset=utf-8" if name.endswith(".js")
                 else "application/octet-stream")
        with open(path, "rb") as fh:
            self._send(fh.read(), ctype)

    def log_message(self, fmt, *args):
        print(f"[web] {self.address_string()} {fmt % args}", flush=True)

    # -- routes ------------------------------------------------------------ #
    def do_OPTIONS(self):
        self._send(b"", "text/plain", 204)

    def do_GET(self):
        path = urlparse(self.path).path

        if path in ("/", "/index.html"):
            return self._file("index.html")

        if path == "/api/fields":
            return self._json(FORM_FIELDS)

        if path == "/api/health":
            cfg = STATE["cfg"]
            return self._json({
                "status": "ok" if (STATE["connected"] or STATE["dry_run"]) else "error",
                "connected": STATE["connected"],
                "dry_run": STATE["dry_run"],
                "hostname": cfg.hostname if cfg else None,
                "site_path": cfg.site_path if cfg else None,
                "list_name": cfg.list_name if cfg else None,
                "site_id": STATE["client"].site_id if STATE["client"] else None,
                "list_id": STATE["client"].list_id if STATE["client"] else None,
                "error": STATE["error"],
                "time": datetime.now(timezone.utc).isoformat(),
            })

        if path == "/api/schema":
            if not STATE["connected"]:
                return self._json({"error": "ยังไม่ได้เชื่อมต่อ SharePoint",
                                   "detail": STATE["error"]}, 503)
            cols = STATE["client"].columns()
            return self._json({"count": len(cols), "columns": [
                {"name": c["name"], "displayName": c.get("displayName"),
                 "readOnly": bool(c.get("readOnly")), "required": bool(c.get("required"))}
                for c in cols]})

        if path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
            return

        return self._send("<h1>404</h1><p><a href='/'>กลับหน้าฟอร์ม</a></p>".encode("utf-8"),
                          "text/html; charset=utf-8", 404)

    def do_POST(self):
        if urlparse(self.path).path != "/api/submit":
            return self._json({"status": "error", "message": "ไม่พบ endpoint นี้"}, 404)
        try:
            n = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(n).decode("utf-8") or "{}")
            return self._json(*self.handle_submit(body))
        except json.JSONDecodeError:
            return self._json({"status": "error", "message": "รูปแบบ JSON ไม่ถูกต้อง"}, 400)
        except Exception as exc:
            traceback.print_exc()
            return self._json({"status": "error", "message": "เกิดข้อผิดพลาดภายในระบบ",
                               "detail": f"{type(exc).__name__}: {exc}"}, 500)

    # -- business logic ---------------------------------------------------- #
    def handle_submit(self, body: dict):
        mode = body.get("mode", "upsert")
        fields = body.get("fields") or {}

        missing = [f["label"] for f in FORM_FIELDS
                   if f.get("required") and not str(fields.get(f["key"], "")).strip()]
        if missing:
            return ({"status": "error",
                     "message": "กรอกข้อมูลไม่ครบ: " + ", ".join(missing)}, 400)

        payload, ignored = build_payload(fields)
        base = {"payload": payload, "ignored": ignored, "field_count": len(payload)}

        if mode == "validate":
            return ({**base, "status": "validated"}, 200)

        if STATE["dry_run"] or not STATE["connected"]:
            return ({**base, "status": "dry_run"}, 200)

        gc, cfg = STATE["client"], STATE["cfg"]
        cust = str(fields.get("Customer_id", "")).strip()

        existing_id = None
        if mode == "upsert" and cust:
            key_field = STATE["resolver"]("Customer_id")
            if key_field:
                existing_id = gc.find_item_id(key_field, cust)

        if existing_id:
            resp = gc.update_item(existing_id, payload)
            ok, status, item_id = resp.status_code in (200, 204), "updated", existing_id
        else:
            resp = gc.create_item(payload)
            ok = resp.status_code in (200, 201)
            status = "created"
            item_id = resp.json().get("id") if ok else None

        if ok:
            link = (f"https://{cfg.hostname}{cfg.site_path}/Lists/{cfg.list_name}"
                    f"/DispForm.aspx?ID={item_id}")
            print(f"[web] {status} item {item_id} (Customer_id={cust})", flush=True)
            return ({**base, "status": status, "item_id": item_id, "link": link}, 200)

        print(f"[web] write failed {resp.status_code}: {resp.text[:300]}", flush=True)
        return ({**base, "status": "error",
                 "message": f"SharePoint ตอบกลับ HTTP {resp.status_code}",
                 "detail": resp.text[:800]}, 502)


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="หน้าเว็บเพิ่มข้อมูลเข้า SharePoint DemoApp")
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--dry-run", action="store_true", help="ไม่เขียนลง SharePoint จริง")
    args = ap.parse_args()

    if not os.path.isfile(os.path.join(BASE_DIR, "index.html")):
        print(f"[web] ไม่พบ index.html ใน {BASE_DIR}", flush=True)
        return 1

    STATE["dry_run"] = args.dry_run
    print(f"[web] dry-run = {args.dry_run}", flush=True)
    connect()
    if STATE["error"]:
        print(f"[web] {STATE['error']}", flush=True)

    try:
        httpd = ThreadingHTTPServer((args.host, args.port), Handler)
    except OSError as exc:
        print(f"[web] เปิดพอร์ต {args.port} ไม่ได้: {exc}\n"
              f"[web] ลองใช้พอร์ตอื่น เช่น --port 8080", flush=True)
        return 1

    print(f"[web] เปิดเบราว์เซอร์ที่  http://{args.host}:{args.port}   (Ctrl+C เพื่อหยุด)",
          flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[web] หยุดทำงานแล้ว", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
