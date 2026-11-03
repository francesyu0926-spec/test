#!/usr/bin/env python3
"""
CBD 行情 HTTP 接口监控（对应文档 2.3 HTTP 接口监控）

默认监控窗口：09:00–23:39（Asia/Shanghai），每 15 分钟由 cron 触发一次。
认证：终端接口需登录，请配置环境变量（见 --help）。
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime, time as dt_time
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

try:
    import yaml
except ImportError:
    print("缺少 PyYAML，请执行: pip install -r requirements.txt", file=sys.stderr)
    sys.exit(2)

TZ = ZoneInfo("Asia/Shanghai")
DEFAULT_CONFIG = Path(__file__).resolve().parent / "endpoints.yaml"
WINDOW_START = dt_time(9, 0)
WINDOW_END = dt_time(23, 39)
SUCCESS_API_CODES = {0, 200, "0", "200"}


@dataclass
class CheckResult:
    name: str
    description: str
    url: str
    ok: bool
    skipped: bool
    http_status: int | None
    elapsed_ms: int
    message: str
    contact: str = ""


def load_config(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict) or "endpoints" not in data:
        raise ValueError(f"无效配置文件: {path}")
    return data


def parse_hhmm(value: str) -> dt_time:
    hour, minute = value.strip().split(":", 1)
    return dt_time(int(hour), int(minute))


def in_monitor_window(now: datetime, force: bool) -> bool:
    if force:
        return True
    t = now.timetz().replace(tzinfo=None)
    return WINDOW_START <= t <= WINDOW_END


def endpoint_active(now: datetime, ep: dict[str, Any], force: bool) -> tuple[bool, str]:
    if not ep.get("enabled", True):
        return False, "disabled"
    not_before = ep.get("not_before")
    if not_before and not force:
        start = parse_hhmm(str(not_before))
        if now.time() < start:
            return False, f"before {not_before}"
    return True, ""


def build_headers() -> dict[str, str]:
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "cbd-quote-http-monitor/1.0",
    }
    token = os.environ.get("QUOTE_MONITOR_TOKEN", "").strip()
    if token:
        if token.lower().startswith("bearer "):
            headers["Authorization"] = token
        else:
            headers["Authorization"] = f"Bearer {token}"

    cookie = os.environ.get("QUOTE_MONITOR_COOKIE", "").strip()
    if cookie:
        headers["Cookie"] = cookie

    extra = os.environ.get("QUOTE_MONITOR_HEADERS_JSON", "").strip()
    if extra:
        parsed = json.loads(extra)
        if isinstance(parsed, dict):
            for k, v in parsed.items():
                headers[str(k)] = str(v)
    return headers


def validate_payload(payload: Any) -> tuple[bool, str]:
    if payload is None:
        return True, "empty body"
    if not isinstance(payload, dict):
        return True, "non-object json"
    code = payload.get("code")
    msg = str(payload.get("msg") or payload.get("message") or "")
    if code in SUCCESS_API_CODES:
        return True, msg or "ok"
    if isinstance(code, int) and 200 <= code < 300:
        return True, msg or "ok"
    if code is not None:
        return False, msg or f"api code={code}"
    if msg and any(x in msg for x in ("失败", "错误", "异常", "未登录")):
        return False, msg
    if "data" in payload or "result" in payload:
        return True, msg or "ok"
    return True, msg or "ok"


def check_one(
    base_url: str,
    ep: dict[str, Any],
    headers: dict[str, str],
    timeout: float,
    now: datetime,
    force: bool,
    default_contact: str,
) -> CheckResult:
    name = str(ep["name"])
    desc = str(ep.get("description") or name)
    contact = str(ep.get("contact") or default_contact)
    url = f"{base_url.rstrip('/')}/{name}"
    active, skip_reason = endpoint_active(now, ep, force)
    if not active:
        return CheckResult(
            name=name,
            description=desc,
            url=url,
            ok=True,
            skipped=True,
            http_status=None,
            elapsed_ms=0,
            message=skip_reason,
            contact=contact,
        )

    body = ep.get("body", {})
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    start = time.perf_counter()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            http_status = resp.status
    except urllib.error.HTTPError as e:
        elapsed_ms = int((time.perf_counter() - start) * 1000)
        try:
            raw = e.read()
            payload = json.loads(raw.decode("utf-8", errors="replace"))
            ok, msg = validate_payload(payload)
            return CheckResult(
                name=name,
                description=desc,
                url=url,
                ok=ok,
                skipped=False,
                http_status=e.code,
                elapsed_ms=elapsed_ms,
                message=msg,
                contact=contact,
            )
        except Exception:
            return CheckResult(
                name=name,
                description=desc,
                url=url,
                ok=False,
                skipped=False,
                http_status=e.code,
                elapsed_ms=elapsed_ms,
                message=str(e),
                contact=contact,
            )
    except Exception as e:
        elapsed_ms = int((time.perf_counter() - start) * 1000)
        return CheckResult(
            name=name,
            description=desc,
            url=url,
            ok=False,
            skipped=False,
            http_status=None,
            elapsed_ms=elapsed_ms,
            message=str(e),
            contact=contact,
        )

    elapsed_ms = int((time.perf_counter() - start) * 1000)
    try:
        payload = json.loads(raw.decode("utf-8", errors="replace"))
    except json.JSONDecodeError as e:
        return CheckResult(
            name=name,
            description=desc,
            url=url,
            ok=False,
            skipped=False,
            http_status=http_status,
            elapsed_ms=elapsed_ms,
            message=f"invalid json: {e}",
            contact=contact,
        )

    ok, msg = validate_payload(payload)
    if http_status >= 400:
        ok = False
        msg = msg or f"http {http_status}"
    return CheckResult(
        name=name,
        description=desc,
        url=url,
        ok=ok,
        skipped=False,
        http_status=http_status,
        elapsed_ms=elapsed_ms,
        message=msg,
        contact=contact,
    )


def notify_webhook(failures: list[CheckResult], webhook_url: str) -> None:
    if not webhook_url or not failures:
        return
    text_lines = [
        f"[CBD HTTP 监控] {len(failures)} 个接口失败 @ {datetime.now(TZ).isoformat(timespec='seconds')}"
    ]
    for r in failures:
        text_lines.append(
            f"- {r.name} ({r.description}): {r.message} | 联系人: {r.contact} | {r.url}"
        )
    payload = {"text": "\n".join(text_lines), "failures": [asdict(x) for x in failures]}
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        webhook_url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=10):
            pass
    except Exception as e:
        print(f"webhook 发送失败: {e}", file=sys.stderr)


def print_human(results: list[CheckResult]) -> None:
    for r in results:
        if r.skipped:
            status = "SKIP"
        elif r.ok:
            status = "OK"
        else:
            status = "FAIL"
        http = r.http_status if r.http_status is not None else "-"
        print(
            f"{status:4}  {r.name:22}  {http:>3}  {r.elapsed_ms:4}ms  {r.message}"
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="CBD 行情 HTTP 接口监控")
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=DEFAULT_CONFIG,
        help=f"接口配置文件 (默认: {DEFAULT_CONFIG})",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="忽略监控时间窗口与 not_before 限制",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=float(os.environ.get("QUOTE_MONITOR_TIMEOUT", "15")),
        help="单接口超时秒数",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=int(os.environ.get("QUOTE_MONITOR_WORKERS", "8")),
        help="并发 worker 数",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="以 JSON 输出结果（便于 ELK / 日志采集）",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="列出配置中的接口并退出",
    )
    args = parser.parse_args()

    cfg = load_config(args.config)
    base_url = os.environ.get("QUOTE_MONITOR_BASE_URL", cfg.get("base_url", "")).strip()
    default_contact = str(cfg.get("default_contact") or "")
    endpoints: list[dict[str, Any]] = cfg.get("endpoints") or []

    if args.list:
        for ep in endpoints:
            flag = "on " if ep.get("enabled", True) else "off"
            nb = ep.get("not_before") or "-"
            print(f"[{flag}] {ep.get('name')}  not_before={nb}  {ep.get('description', '')}")
        return 0

    now = datetime.now(TZ)
    if not in_monitor_window(now, args.force):
        msg = f"当前 {now.strftime('%H:%M')} 不在监控窗口 {WINDOW_START.strftime('%H:%M')}-{WINDOW_END.strftime('%H:%M')}"
        if args.json:
            print(json.dumps({"skipped_window": True, "message": msg}, ensure_ascii=False))
        else:
            print(msg)
        return 0

    if not base_url:
        print("未配置 base_url", file=sys.stderr)
        return 2

    headers = build_headers()
    if not headers.get("Authorization") and not headers.get("Cookie"):
        print(
            "警告: 未设置 QUOTE_MONITOR_TOKEN 或 QUOTE_MONITOR_COOKIE，接口可能返回未登录(401)",
            file=sys.stderr,
        )

    results: list[CheckResult] = []
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futures = [
            pool.submit(
                check_one,
                base_url,
                ep,
                headers,
                args.timeout,
                now,
                args.force,
                default_contact,
            )
            for ep in endpoints
        ]
        for fut in as_completed(futures):
            results.append(fut.result())

    results.sort(key=lambda r: r.name)
    failures = [r for r in results if not r.skipped and not r.ok]
    webhook = os.environ.get("QUOTE_MONITOR_WEBHOOK_URL", "").strip()
    notify_webhook(failures, webhook)

    if args.json:
        out = {
            "timestamp": now.isoformat(timespec="seconds"),
            "base_url": base_url,
            "total": len(results),
            "failed": len(failures),
            "results": [asdict(r) for r in results],
        }
        print(json.dumps(out, ensure_ascii=False))
    else:
        print_human(results)
        print(
            f"\n汇总: 检测 {len([r for r in results if not r.skipped])} 项, "
            f"失败 {len(failures)} 项, 跳过 {len([r for r in results if r.skipped])} 项"
        )

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
