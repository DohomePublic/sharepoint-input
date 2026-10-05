#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Upload / upsert rows into a SharePoint Online list via Microsoft Graph.

Auth: Azure AD app registration (client credentials flow).
Secrets are read from environment variables ONLY. Never hard-code them.

Supported secret names (first non-empty wins, logged by NAME only):
    AZ_CLIENT_ID      | AZURE_CLIENT_ID
    AZ_TENANT_ID      | AZURE_TENANT_ID
    AZ_CLIENT_SECRET  | AZURE_CLIENT_SECRET

Optional variables (defaults shown):
    SP_HOSTNAME   = dohomegroup.sharepoint.com
    SP_SITE_PATH  = /sites/AC-Accounting
    SP_LIST_NAME  = DemoApp

Usage:
    python scripts/sp_upload.py --input data/upload.csv --dry-run
    python scripts/sp_upload.py --input data/upload.csv --key-column Customer_id
    python scripts/sp_upload.py --input data/upload.json --mode insert
    python scripts/sp_upload.py --show-schema
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests

# Endpoints (overridable only for local testing against a mock server)
GRAPH = os.environ.get("GRAPH_BASE_URL", "https://graph.microsoft.com/v1.0")
LOGIN = os.environ.get("LOGIN_BASE_URL", "https://login.microsoftonline.com")

DEFAULTS = {
    "SP_HOSTNAME": "dohomegroup.sharepoint.com",
    "SP_SITE_PATH": "/sites/AC-Accounting",
    "SP_LIST_NAME": "DemoApp",
}

# Graph refuses writes to these; they are system-managed.
READONLY_FIELDS = {
    "id", "ID", "ContentType", "Created", "Modified", "Author", "Editor",
    "_UIVersionString", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle",
    "DocIcon", "ItemChildCount", "FolderChildCount", "_ComplianceFlags",
    "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId",
    "AppAuthor", "AppEditor", "ContentTypeId", "ComplianceAssetId",
}


# --------------------------------------------------------------------------- #
# logging helpers
# --------------------------------------------------------------------------- #
def log(msg: str) -> None:
    print(f"[sp-upload] {msg}", flush=True)


def fail(msg: str, code: int = 1) -> None:
    print(f"::error::[sp-upload] {msg}", flush=True)
    sys.exit(code)


def pick_env(*names: str) -> Tuple[Optional[str], Optional[str]]:
    """Return (value, source_name) of the first env var that has a value."""
    for name in names:
        val = os.environ.get(name)
        if val and val.strip():
            return val.strip(), name
    return None, None


def mask(value: str) -> str:
    """Never print a secret. Show length only."""
    return f"<hidden, {len(value)} chars>"


# --------------------------------------------------------------------------- #
# config
# --------------------------------------------------------------------------- #
class Config:
    def __init__(self) -> None:
        self.client_id, src_id = pick_env("AZ_CLIENT_ID", "AZURE_CLIENT_ID")
        self.tenant_id, src_tn = pick_env("AZ_TENANT_ID", "AZURE_TENANT_ID")
        self.client_secret, src_sc = pick_env("AZ_CLIENT_SECRET", "AZURE_CLIENT_SECRET")

        log(f"client id   <- {src_id or 'MISSING'}")
        log(f"tenant id   <- {src_tn or 'MISSING'}")
        log(f"client secret <- {src_sc or 'MISSING'} "
            f"({mask(self.client_secret) if self.client_secret else 'no value'})")

        self.hostname = os.environ.get("SP_HOSTNAME") or DEFAULTS["SP_HOSTNAME"]
        self.site_path = os.environ.get("SP_SITE_PATH") or DEFAULTS["SP_SITE_PATH"]
        self.list_name = os.environ.get("SP_LIST_NAME") or DEFAULTS["SP_LIST_NAME"]
        if not self.site_path.startswith("/"):
            self.site_path = "/" + self.site_path

        log(f"target: https://{self.hostname}{self.site_path} :: list '{self.list_name}'")

    def require_credentials(self) -> None:
        missing = [n for n, v in (
            ("AZ_CLIENT_ID / AZURE_CLIENT_ID", self.client_id),
            ("AZ_TENANT_ID / AZURE_TENANT_ID", self.tenant_id),
            ("AZ_CLIENT_SECRET / AZURE_CLIENT_SECRET", self.client_secret),
        ) if not v]
        if missing:
            fail("missing credentials: " + ", ".join(missing) +
                 ". Add them under Settings > Secrets and variables > Actions.")


# --------------------------------------------------------------------------- #
# Graph client
# --------------------------------------------------------------------------- #
class GraphClient:
    def __init__(self, cfg: Config) -> None:
        self.cfg = cfg
        self.session = requests.Session()
        self._token: Optional[str] = None
        self._token_exp: float = 0.0
        self.site_id: Optional[str] = None
        self.list_id: Optional[str] = None

    # -- auth ------------------------------------------------------------- #
    def token(self) -> str:
        if self._token and time.time() < self._token_exp - 120:
            return self._token
        url = f"{LOGIN}/{self.cfg.tenant_id}/oauth2/v2.0/token"
        try:
            resp = self.session.post(url, data={
                "client_id": self.cfg.client_id,
                "client_secret": self.cfg.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            }, timeout=60)
        except Exception as exc:  # network / TLS / DNS
            fail(f"cannot reach {LOGIN}: {type(exc).__name__}: {exc}")
        if resp.status_code != 200:
            fail(f"token request failed ({resp.status_code}): {resp.text[:500]}")
        payload = resp.json()
        self._token = payload["access_token"]
        self._token_exp = time.time() + int(payload.get("expires_in", 3600))
        log("access token acquired (client credentials flow)")
        return self._token

    def request(self, method: str, url: str, **kw) -> requests.Response:
        if not url.startswith("http"):
            url = GRAPH + url
        headers = kw.pop("headers", {}) or {}
        headers["Authorization"] = f"Bearer {self.token()}"
        headers.setdefault("Accept", "application/json")

        for attempt in range(5):
            try:
                resp = self.session.request(method, url, headers=headers, timeout=90, **kw)
            except Exception as exc:
                if attempt == 4:
                    fail(f"network error calling Graph: {type(exc).__name__}: {exc}")
                wait = 2 ** attempt
                log(f"network error ({type(exc).__name__}); retrying in {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code in (429, 503, 504):
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                log(f"throttled ({resp.status_code}); retrying in {wait}s")
                time.sleep(wait)
                continue
            return resp
        return resp  # type: ignore[return-value]

    # -- discovery -------------------------------------------------------- #
    def resolve_site_and_list(self) -> None:
        r = self.request("GET", f"/sites/{self.cfg.hostname}:{self.cfg.site_path}")
        if r.status_code != 200:
            fail(f"cannot resolve site ({r.status_code}): {r.text[:500]}")
        self.site_id = r.json()["id"]
        log(f"site id: {self.site_id}")

        r = self.request("GET", f"/sites/{self.site_id}/lists?$select=id,name,displayName&$top=200")
        if r.status_code != 200:
            fail(f"cannot enumerate lists ({r.status_code}): {r.text[:500]}")
        wanted = self.cfg.list_name.strip().lower()
        for lst in r.json().get("value", []):
            if lst.get("displayName", "").lower() == wanted or lst.get("name", "").lower() == wanted:
                self.list_id = lst["id"]
                log(f"list id: {self.list_id} ({lst.get('displayName')})")
                return
        fail(f"list '{self.cfg.list_name}' not found on the site")

    def columns(self) -> List[Dict[str, Any]]:
        r = self.request(
            "GET",
            f"/sites/{self.site_id}/lists/{self.list_id}/columns"
            "?$select=name,displayName,readOnly,required,hidden,columnGroup",
        )
        if r.status_code != 200:
            fail(f"cannot read list columns ({r.status_code}): {r.text[:500]}")
        return r.json().get("value", [])

    # -- write ------------------------------------------------------------ #
    def find_item_id(self, key_field: str, key_value: str) -> Optional[str]:
        safe = str(key_value).replace("'", "''")
        url = (f"/sites/{self.site_id}/lists/{self.list_id}/items"
               f"?$expand=fields($select=id)&$filter=fields/{key_field} eq '{safe}'&$top=1")
        r = self.request("GET", url, headers={
            "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"})
        if r.status_code != 200:
            log(f"lookup on {key_field}='{key_value}' failed ({r.status_code}): {r.text[:200]}")
            return None
        items = r.json().get("value", [])
        return items[0]["id"] if items else None

    def create_item(self, fields: Dict[str, Any]) -> requests.Response:
        return self.request(
            "POST", f"/sites/{self.site_id}/lists/{self.list_id}/items",
            json={"fields": fields},
            headers={"Content-Type": "application/json"},
        )

    def update_item(self, item_id: str, fields: Dict[str, Any]) -> requests.Response:
        return self.request(
            "PATCH", f"/sites/{self.site_id}/lists/{self.list_id}/items/{item_id}/fields",
            json=fields,
            headers={"Content-Type": "application/json"},
        )


# --------------------------------------------------------------------------- #
# input handling
# --------------------------------------------------------------------------- #
def read_rows(path: str) -> List[Dict[str, Any]]:
    if not os.path.exists(path):
        fail(f"input file not found: {path}")
    if path.lower().endswith(".json"):
        with open(path, encoding="utf-8-sig") as fh:
            data = json.load(fh)
        rows = data["items"] if isinstance(data, dict) and "items" in data else data
        if not isinstance(rows, list):
            fail("JSON input must be a list of objects, or {'items': [...]}")
        return rows
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


def load_field_map(path: Optional[str]) -> Dict[str, str]:
    if not path:
        return {}
    if not os.path.exists(path):
        fail(f"field map not found: {path}")
    with open(path, encoding="utf-8-sig") as fh:
        return json.load(fh)


def build_resolver(columns: List[Dict[str, Any]],
                   field_map: Dict[str, str]):
    """Map an input header to a writable SharePoint internal field name."""
    by_internal = {c["name"]: c for c in columns}
    by_display = {}
    for c in columns:
        by_display.setdefault(c.get("displayName", "").strip().lower(), c)

    def resolve(header: str) -> Optional[str]:
        h = field_map.get(header, header).strip()
        col = by_internal.get(h) or by_display.get(h.lower())
        if not col:
            return None
        if col["name"] in READONLY_FIELDS or col.get("readOnly"):
            return None
        return col["name"]

    return resolve


AMOUNT_LIKE = re.compile(r"^\d{1,3}(,\d{3})+(\.\d+)?$")


def coerce(value: Any, clean_amounts: bool = True) -> Any:
    """ตัดช่องว่าง, แปลงค่าว่างเป็น None และ normalize ตัวเลขที่มีลูกน้ำคั่นหลักพัน"""
    if value is None:
        return None
    if isinstance(value, str):
        v = value.strip()
        if v == "":
            return None
        # '5,000,000' -> '5000000' (คอลัมน์จำนวนเงินในลิสต์เก็บเป็น Text แต่รูปแบบปนกัน)
        if clean_amounts and AMOUNT_LIKE.match(v):
            return v.replace(",", "")
        return v
    return value


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main() -> int:
    ap = argparse.ArgumentParser(description="Upload rows into a SharePoint list via Microsoft Graph")
    ap.add_argument("--input", help="CSV or JSON file with the rows to send")
    ap.add_argument("--field-map", default="config/field_map.json",
                    help="JSON mapping of input header -> SharePoint field (optional)")
    ap.add_argument("--key-column", help="Input column used to detect existing items (enables upsert)")
    ap.add_argument("--mode", choices=["insert", "upsert"], default="upsert")
    ap.add_argument("--batch-sleep", type=float, default=0.0, help="Seconds to pause between writes")
    ap.add_argument("--limit", type=int, default=0, help="Process only the first N rows (0 = all)")
    ap.add_argument("--keep-raw-numbers", action="store_true",
                    help="Do NOT strip thousand separators (default strips '5,000,000' -> '5000000')")
    ap.add_argument("--dry-run", action="store_true",
                    help="Validate input + field mapping without writing to SharePoint")
    ap.add_argument("--show-schema", action="store_true",
                    help="Print the list's writable columns and exit")
    args = ap.parse_args()

    cfg = Config()

    # ---- offline validation path (no credentials needed) ----------------- #
    if args.dry_run and not (cfg.client_id and cfg.client_secret and cfg.tenant_id):
        log("dry-run without credentials: validating input file only")
        if not args.input:
            fail("--input is required")
        rows = read_rows(args.input)
        fmap = load_field_map(args.field_map if os.path.exists(args.field_map) else None)
        headers = sorted({k for r in rows for k in r.keys()})
        log(f"rows parsed: {len(rows)}")
        log(f"headers ({len(headers)}): {', '.join(headers)}")
        if fmap:
            log(f"field map entries: {len(fmap)}")
            unmapped = [h for h in headers if h not in fmap]
            if unmapped:
                log(f"headers without an explicit mapping (will match by name): {', '.join(unmapped)}")
        if args.key_column and args.key_column not in headers:
            fail(f"--key-column '{args.key_column}' is not present in the input")
        log("input OK. Provide the secrets to perform a live run.")
        return 0

    cfg.require_credentials()
    gc = GraphClient(cfg)
    gc.resolve_site_and_list()
    cols = gc.columns()

    if args.show_schema:
        writable = [c for c in cols if not c.get("readOnly") and c["name"] not in READONLY_FIELDS]
        log(f"{len(writable)} writable columns:")
        for c in sorted(writable, key=lambda x: x["name"]):
            flags = []
            if c.get("required"):
                flags.append("required")
            if c.get("hidden"):
                flags.append("hidden")
            print(f"  {c['name']:<32} | display: {c.get('displayName','')}"
                  f"{' | ' + ','.join(flags) if flags else ''}")
        return 0

    if not args.input:
        fail("--input is required (or use --show-schema)")

    rows = read_rows(args.input)
    if args.limit:
        rows = rows[: args.limit]
    log(f"rows to process: {len(rows)}")

    fmap = load_field_map(args.field_map if os.path.exists(args.field_map) else None)
    resolve = build_resolver(cols, fmap)

    key_field = resolve(args.key_column) if args.key_column else None
    if args.key_column and not key_field:
        fail(f"--key-column '{args.key_column}' does not map to a writable list column")

    created = updated = skipped = failed = 0
    dropped_headers: set = set()

    for idx, row in enumerate(rows, start=1):
        fields: Dict[str, Any] = {}
        for header, value in row.items():
            target = resolve(header)
            if not target:
                dropped_headers.add(header)
                continue
            v = coerce(value, clean_amounts=not args.keep_raw_numbers)
            if v is not None:
                fields[target] = v

        if not fields:
            log(f"row {idx}: no mappable data, skipped")
            skipped += 1
            continue

        existing_id = None
        if args.mode == "upsert" and key_field and args.key_column in row:
            key_val = coerce(row[args.key_column], clean_amounts=False)
            if key_val is not None and not args.dry_run:
                existing_id = gc.find_item_id(key_field, key_val)

        if args.dry_run:
            log(f"row {idx}: DRY-RUN would {'update' if existing_id else 'create'} "
                f"with {len(fields)} fields -> {json.dumps(fields, ensure_ascii=False)[:300]}")
            skipped += 1
            continue

        if existing_id:
            resp = gc.update_item(existing_id, fields)
            ok, verb = resp.status_code in (200, 204), "updated"
        else:
            resp = gc.create_item(fields)
            ok, verb = resp.status_code in (200, 201), "created"

        if ok:
            if verb == "updated":
                updated += 1
            else:
                created += 1
            log(f"row {idx}: {verb}")
        else:
            failed += 1
            print(f"::error::row {idx} failed ({resp.status_code}): {resp.text[:400]}", flush=True)

        if args.batch_sleep:
            time.sleep(args.batch_sleep)

    if dropped_headers:
        log("columns ignored (not writable / not found in the list): "
            + ", ".join(sorted(dropped_headers)))

    log(f"summary: created={created} updated={updated} skipped={skipped} failed={failed}")

    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary_path:
        with open(summary_path, "a", encoding="utf-8") as fh:
            fh.write(f"### SharePoint upload — {cfg.list_name}\n\n"
                     f"| created | updated | skipped | failed |\n|---|---|---|---|\n"
                     f"| {created} | {updated} | {skipped} | {failed} |\n")

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
