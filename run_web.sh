#!/usr/bin/env bash
# เปิดหน้าเว็บเพิ่มข้อมูล  ใช้:  ./run_web.sh   หรือ  ./run_web.sh --dry-run
set -euo pipefail
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt --quiet || true
if [ -f .env ]; then set -a; . ./.env; set +a; echo "loaded .env"; fi
exec python3 webapp/server.py --port "${PORT:-8000}" "$@"
