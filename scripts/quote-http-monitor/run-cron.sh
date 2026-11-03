#!/usr/bin/env bash
# 文档要求：09:00–23:39 每 15 分钟检查一次
# crontab 示例（Asia/Shanghai）:
# */15 9-23 * * 1-5 /path/to/scripts/quote-http-monitor/run-cron.sh

set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

if [[ -f .env ]]; then
  set -a
  # shellcheck disable=SC1091
  source .env
  set +a
fi

python3 -m pip install -q -r requirements.txt
exec python3 monitor.py --json
