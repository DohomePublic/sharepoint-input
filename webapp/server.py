#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
หน้าเว็บสำหรับเพิ่มข้อมูลเข้า SharePoint List 'DemoApp' และทดสอบการเชื่อมต่อ

ใช้ Python standard library เท่านั้น (+ requests ผ่าน scripts/sp_upload.py)
ไม่ต้องติดตั้ง Flask/Django

    python webapp/server.py --port 8000
    python webapp/server.py --port 8000 --dry-run     # ไม่เขียนจริง

เปิดเบราว์เซอร์ที่ http://127.0.0.1:8000

Endpoints
    GET  /            ฟอร์มเพิ่มข้อมูล
    POST /submit      บันทึกเข้า SharePoint (หรือจำลองเมื่อ --dry-run)
    GET  /api/health  สถานะการเชื่อมต่อ (JSON)
    GET  /api/schema  คอลัมน์ที่เขียนได้ของ list (JSON)
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from sp_upload import Config, GraphClient, build_resolver, load_field_map  # noqa: E402

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))

# ---------------------------------------------------------------------------
# นิยามฟิลด์บนฟอร์ม — อิงคอลัมน์จริงของ list DemoApp
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
     "type": "select",
     "options": ["", "10 - 50 ล้านบาท", "มากกว่า 50 ล้านบาท"]},
    {"key": "limit", "label": "limit (วงเงินที่ขอ)", "type": "text", "placeholder": "1000000",
     "help": "คอลัมน์นี้เป็น Text ในลิสต์ — ระบบจะตัดเครื่องหมาย , ให้อัตโนมัติ"},
    {"key": "province", "label": "province (จังหวัด)", "type": "text"},
    {"key": "district", "label": "district (อำเภอ/เขต)", "type": "text"},
    {"key": "telephone", "label": "telephone (โทรศัพท์)", "type": "text"},
    {"key": "contact_name", "label": "contact_name (ผู้ติดต่อ)", "type": "text"},
    {"key": "contact_number", "label": "contact_number (เบอร์ผู้ติดต่อ)", "type": "text"},
]

NUMERIC_TEXT_FIELDS = {"limit", "CraditApprove"}

STATE = {"dry_run": False, "cfg": None, "client": None, "resolver": None,
         "connected": False, "error": None, "log": []}


# ---------------------------------------------------------------------------
def esc(v) -> str:
    if v is None:
        return ""
    return (str(v).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def clean_amount(value: str) -> str:
    """'5,000,000' -> '5000000' (คอลัมน์ limit เก็บเป็น Text แต่ปนรูปแบบ)"""
    return value.strip().replace(",", "").replace(" ", "")


def connect() -> None:
    """เชื่อมต่อ Graph และเตรียม resolver ของคอลัมน์"""
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


def record_log(entry: dict) -> None:
    STATE["log"].insert(0, entry)
    del STATE["log"][20:]


# ---------------------------------------------------------------------------
def render_page(values=None, message=None, level="info") -> str:
    values = values or {}
    cfg = STATE["cfg"]
    target = (f"https://{cfg.hostname}{cfg.site_path} / {cfg.list_name}"
              if cfg else "ยังไม่ได้ตั้งค่า")

    if STATE["dry_run"]:
        badge = '<span class="badge warn">DRY-RUN — ไม่เขียนจริง</span>'
    elif STATE["connected"]:
        badge = '<span class="badge ok">เชื่อมต่อ SharePoint แล้ว</span>'
    else:
        badge = '<span class="badge err">ยังไม่ได้เชื่อมต่อ</span>'

    rows = []
    for f in FORM_FIELDS:
        key, val = f["key"], values.get(f["key"], "")
        req = " required" if f.get("required") else ""
        star = ' <span class="req">*</span>' if f.get("required") else ""
        help_html = f'<small>{esc(f["help"])}</small>' if f.get("help") else ""
        if f["type"] == "select":
            opts = "".join(
                f'<option value="{esc(o)}"{" selected" if str(val) == o else ""}>'
                f'{esc(o) if o else "— เลือก —"}</option>' for o in f["options"])
            ctl = f'<select name="{esc(key)}" id="{esc(key)}"{req}>{opts}</select>'
        else:
            ph = esc(f.get("placeholder", ""))
            ctl = (f'<input type="text" name="{esc(key)}" id="{esc(key)}" '
                   f'value="{esc(val)}" placeholder="{ph}"{req}>')
        rows.append(f'<div class="field"><label for="{esc(key)}">{esc(f["label"])}{star}'
                    f'</label>{ctl}{help_html}</div>')

    msg_html = f'<div class="msg {level}">{message}</div>' if message else ""
    err_html = f'<div class="msg warn">{esc(STATE["error"])}</div>' if STATE["error"] else ""

    log_html = ""
    if STATE["log"]:
        items = "".join(
            f'<tr><td>{esc(e["time"])}</td><td>{esc(e["action"])}</td>'
            f'<td>{esc(e["customer"])}</td><td class="{e["cls"]}">{esc(e["result"])}</td></tr>'
            for e in STATE["log"])
        log_html = ("<h2>ประวัติการส่ง (รอบนี้)</h2>"
                    "<table class='log'><thead><tr><th>เวลา (UTC)</th><th>การทำงาน</th>"
                    f"<th>Customer_id</th><th>ผลลัพธ์</th></tr></thead><tbody>{items}"
                    "</tbody></table>")

    fields_html = "".join(rows)

    return f"""<!DOCTYPE html>
<html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>เพิ่มข้อมูล — SharePoint DemoApp</title>
<style>
 :root {{ --dh:#e4002b; --bg:#f4f6f8; --line:#d8dee4; }}
 * {{ box-sizing:border-box; }}
 body {{ margin:0; background:var(--bg); color:#16202a;
        font-family:"Segoe UI",system-ui,"Noto Sans Thai",sans-serif; }}
 header {{ background:var(--dh); color:#fff; padding:18px 24px; }}
 header h1 {{ margin:0; font-size:19px; }}
 header p {{ margin:6px 0 0; font-size:13px; opacity:.92; }}
 .wrap {{ max-width:1000px; margin:22px auto; padding:0 16px 48px; }}
 .card {{ background:#fff; border:1px solid var(--line); border-radius:10px;
          padding:22px; margin-bottom:18px; }}
 .grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(290px,1fr)); gap:14px; }}
 .field {{ display:flex; flex-direction:column; }}
 label {{ font-size:13px; font-weight:600; margin-bottom:5px; }}
 input,select {{ padding:9px 10px; border:1px solid var(--line); border-radius:6px;
                 font-size:14px; background:#fff; font-family:inherit; }}
 input:focus,select:focus {{ outline:2px solid var(--dh); outline-offset:-1px; }}
 small {{ color:#5b6b7b; font-size:12px; margin-top:4px; }}
 .req {{ color:var(--dh); }}
 .actions {{ margin-top:20px; display:flex; gap:10px; flex-wrap:wrap; }}
 button {{ background:var(--dh); color:#fff; border:0; border-radius:6px;
           padding:11px 22px; font-size:14px; font-weight:600; cursor:pointer;
           font-family:inherit; }}
 button.sec {{ background:#5b6b7b; }}
 .badge {{ display:inline-block; padding:3px 10px; border-radius:11px;
           font-size:12px; font-weight:600; }}
 .ok {{ background:#d8f5e0; color:#116b33; }}
 .warn {{ background:#fff3cd; color:#8a6100; }}
 .err {{ background:#fde2e4; color:#9b1c25; }}
 .msg {{ padding:12px 14px; border-radius:7px; margin-bottom:16px; font-size:14px;
         line-height:1.55; }}
 .msg.info {{ background:#e4f0fb; color:#13497a; }}
 .msg.ok {{ background:#d8f5e0; color:#116b33; }}
 .msg.err {{ background:#fde2e4; color:#9b1c25; }}
 .msg.warn {{ background:#fff3cd; color:#8a6100; }}
 pre {{ background:#1e2630; color:#e8edf2; padding:12px; border-radius:6px;
        overflow-x:auto; font-size:12.5px; }}
 table.log {{ width:100%; border-collapse:collapse; font-size:13px; }}
 table.log th,table.log td {{ border-bottom:1px solid var(--line);
                              padding:7px 9px; text-align:left; }}
 table.log th {{ background:#eef2f6; }}
 h2 {{ font-size:15px; margin:22px 0 10px; }}
 code {{ background:#eef2f6; padding:1px 5px; border-radius:4px; font-size:13px; }}
 @media (max-width:640px) {{ .grid {{ grid-template-columns:1fr; }}
                             button {{ width:100%; }} }}
</style></head><body>
<header>
  <h1>เพิ่มข้อมูลเข้า SharePoint List — DemoApp</h1>
  <p>{esc(target)} &nbsp; {badge}</p>
</header>
<div class="wrap">
  {msg_html}{err_html}
  <form method="post" action="/submit" class="card">
    <div class="grid">{fields_html}</div>
    <div class="actions">
      <button type="submit" name="mode" value="upsert">บันทึกเข้า SharePoint (Upsert)</button>
      <button type="submit" name="mode" value="insert" class="sec">สร้างรายการใหม่เสมอ</button>
      <button type="submit" name="mode" value="validate" class="sec">ตรวจสอบอย่างเดียว</button>
    </div>
  </form>
  <div class="card">
    <h2 style="margin-top:0">เครื่องมือทดสอบ</h2>
    <p style="font-size:14px;line-height:1.7;margin:0">
      <a href="/api/health">/api/health</a> — ตรวจสถานะการเชื่อมต่อ<br>
      <a href="/api/schema">/api/schema</a> — ดูคอลัมน์ที่เขียนได้จริง (internal name)<br>
      ส่งเป็นชุดจากไฟล์: <code>python scripts/sp_upload.py --input data/upload.csv
      --key-column Customer_id</code>
    </p>
    {log_html}
  </div>
</div></body></html>"""


# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "DemoAppUploader/1.0"

    def _html(self, body: str, code: int = 200) -> None:
        data = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _json(self, obj, code: int = 200) -> None:
        data = json.dumps(obj, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        print(f"[web] {self.address_string()} {fmt % args}", flush=True)

    # -- GET --------------------------------------------------------------- #
    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/":
            return self._html(render_page())
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
        return self._html("<h1>404</h1><p><a href='/'>กลับหน้าฟอร์ม</a></p>", 404)

    # -- POST -------------------------------------------------------------- #
    def do_POST(self):
        if urlparse(self.path).path != "/submit":
            return self._html("<h1>404</h1>", 404)
        try:
            n = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(n).decode("utf-8")
            form = {k: v[0] for k, v in parse_qs(raw, keep_blank_values=True).items()}
            return self._html(self.handle_submit(form))
        except Exception:
            traceback.print_exc()
            return self._html(render_page(
                message="เกิดข้อผิดพลาดภายในระบบ — ดูรายละเอียดใน console",
                level="err"), 500)

    def handle_submit(self, form: dict) -> str:
        mode = form.pop("mode", "upsert")
        values = {f["key"]: form.get(f["key"], "") for f in FORM_FIELDS}

        missing = [f["label"] for f in FORM_FIELDS
                   if f.get("required") and not values.get(f["key"], "").strip()]
        if missing:
            return render_page(values,
                               "กรอกข้อมูลไม่ครบ: " + ", ".join(esc(m) for m in missing), "err")

        payload, unmapped = {}, []
        for f in FORM_FIELDS:
            raw_val = values.get(f["key"], "").strip()
            if not raw_val:
                continue
            sp_name = f.get("sp_key", f["key"])
            if sp_name in NUMERIC_TEXT_FIELDS:
                raw_val = clean_amount(raw_val)
            if STATE["resolver"]:
                target = STATE["resolver"](sp_name)
                if not target:
                    unmapped.append(sp_name)
                    continue
            else:
                target = sp_name
            payload[target] = raw_val

        pretty = json.dumps(payload, ensure_ascii=False, indent=2)
        note = (f"<br>คอลัมน์ที่ลิสต์ไม่มีหรือเขียนไม่ได้ (ข้าม): "
                f"{esc(', '.join(unmapped))}" if unmapped else "")
        cust = values.get("Customer_id", "")

        if mode == "validate":
            return render_page(values,
                               f"ตรวจสอบผ่าน — ข้อมูลที่จะส่ง {len(payload)} ฟิลด์:{note}"
                               f"<pre>{esc(pretty)}</pre>", "info")

        if STATE["dry_run"] or not STATE["connected"]:
            record_log({"time": datetime.now(timezone.utc).strftime("%H:%M:%S"),
                        "action": f"{mode} (dry-run)", "customer": cust,
                        "result": "จำลอง — ไม่เขียนจริง", "cls": "warn"})
            return render_page(values,
                               f"<b>DRY-RUN</b> — ไม่ได้เขียนลง SharePoint จริง "
                               f"({len(payload)} ฟิลด์):{note}<pre>{esc(pretty)}</pre>", "warn")

        gc = STATE["client"]
        existing_id = None
        if mode == "upsert" and cust:
            key_field = STATE["resolver"]("Customer_id")
            if key_field:
                existing_id = gc.find_item_id(key_field, cust)

        if existing_id:
            resp = gc.update_item(existing_id, payload)
            ok, verb, item_id = resp.status_code in (200, 204), "อัปเดต", existing_id
        else:
            resp = gc.create_item(payload)
            ok = resp.status_code in (200, 201)
            verb = "สร้าง"
            item_id = resp.json().get("id") if ok else None

        ts = datetime.now(timezone.utc).strftime("%H:%M:%S")
        if ok:
            cfg = STATE["cfg"]
            link = (f"https://{cfg.hostname}{cfg.site_path}/Lists/{cfg.list_name}"
                    f"/DispForm.aspx?ID={item_id}")
            record_log({"time": ts, "action": mode, "customer": cust,
                        "result": f"{verb}สำเร็จ (ID {item_id})", "cls": "ok"})
            return render_page({}, f"{verb}รายการสำเร็จ — item ID <b>{esc(item_id)}</b>"
                                   f' &nbsp;<a href="{esc(link)}" target="_blank">'
                                   f"เปิดใน SharePoint</a>{note}", "ok")

        record_log({"time": ts, "action": mode, "customer": cust,
                    "result": f"ผิดพลาด HTTP {resp.status_code}", "cls": "err"})
        return render_page(values,
                           f"บันทึกไม่สำเร็จ (HTTP {resp.status_code})"
                           f"<pre>{esc(resp.text[:800])}</pre>", "err")


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="หน้าเว็บเพิ่มข้อมูลเข้า SharePoint DemoApp")
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--dry-run", action="store_true", help="ไม่เขียนลง SharePoint จริง")
    args = ap.parse_args()

    STATE["dry_run"] = args.dry_run
    print(f"[web] dry-run = {args.dry_run}", flush=True)
    connect()
    if STATE["error"]:
        print(f"[web] {STATE['error']}", flush=True)

    httpd = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"[web] serving on http://{args.host}:{args.port}  (Ctrl+C to stop)", flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[web] stopped", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
