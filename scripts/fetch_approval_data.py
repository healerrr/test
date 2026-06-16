"""Standalone CLI script to fetch DingTalk OA approval data.

This script directly calls DingTalk OpenAPI without the MCP protocol layer.
Credentials are hard-coded as defaults; environment variables override them if set.

Usage:
    python fetch_approval_data.py \
        --process-code PROC-402A087A-F4A1-4B4D-9726-29BA08FD773D \
        --year 2026 --month 4 \
        --output-dir . \
        [--max-instances 20000] \
        [--include-details]

Hard-coded credentials (can be overridden by environment variables):
    DINGTALK_APP_KEY     - defaults to built-in value
    DINGTALK_APP_SECRET  - defaults to built-in value
    DINGTALK_USER_ID     - defaults to built-in value

Optional environment variable:
    DINGTALK_ACCESS_TOKEN - If set, skip app_key/app_secret token exchange
"""

import argparse
import calendar
import json
import os
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urlencode

import requests


OAPI_BASE = "https://oapi.dingtalk.com"
TZ = timezone(timedelta(hours=8))

# Hard-coded DingTalk credentials
DEFAULT_APP_KEY = "dingtyzldpxnwvoonzxm"
DEFAULT_APP_SECRET = "L10kCqqStjF8MmMDZCPYbP6LtzHfgU_oEBtGdftwzJaxRtDpvAicQr7m-ScVzA3C"
DEFAULT_USER_ID = "16248445393404993"

# All supported approval process codes
PROCESS_CODES = {
    "通用费用报销（人民币）": "PROC-402A087A-F4A1-4B4D-9726-29BA08FD773D",
    "差旅费用报销(人民币版)": "PROC-4E9EF26C-F477-4641-A103-CDC573812CC7",
    "经办付款申请单（人民币版）": "PROC-3JYJ9N2V-6AYV91D7SRD9DSUF6QLW1-504EJ9IJ-1",
    "经办付款申请单（外币版）": "PROC-RIYJS65W-8CSWSZ9SSFAXV8GGN8BY1-5FMPCIJJ-91",
}

_TOKEN_CACHE: dict = {"access_token": None, "expires_at": 0}


def _get_access_token() -> str:
    token = os.environ.get("DINGTALK_ACCESS_TOKEN")
    if token:
        return token

    now = int(time.time())
    cached = _TOKEN_CACHE.get("access_token")
    if cached and int(_TOKEN_CACHE.get("expires_at", 0)) > now + 60:
        return str(cached)

    app_key = os.environ.get("DINGTALK_APP_KEY", DEFAULT_APP_KEY)
    app_secret = os.environ.get("DINGTALK_APP_SECRET", DEFAULT_APP_SECRET)
    if not app_key or not app_secret:
        raise RuntimeError(
            "Missing DingTalk credentials. Set DINGTALK_ACCESS_TOKEN, "
            "or set DINGTALK_APP_KEY and DINGTALK_APP_SECRET."
        )

    url = f"{OAPI_BASE}/gettoken?{urlencode({'appkey': app_key, 'appsecret': app_secret})}"
    payload = _request("GET", url, auth=False)
    access_token = payload.get("access_token")
    if not access_token:
        raise RuntimeError(f"DingTalk gettoken did not return access_token: {payload}")

    expires_in = int(payload.get("expires_in", 7200))
    _TOKEN_CACHE["access_token"] = access_token
    _TOKEN_CACHE["expires_at"] = now + expires_in
    return str(access_token)


def _request(method: str, url: str, body: dict | None = None, auth: bool = True) -> dict:
    if auth:
        separator = "&" if "?" in url else "?"
        url = f"{url}{separator}{urlencode({'access_token': _get_access_token()})}"

    response = requests.request(method, url, json=body, timeout=30)
    response.raise_for_status()
    payload = response.json()

    errcode = payload.get("errcode")
    if errcode not in (None, 0):
        errmsg = payload.get("errmsg") or payload.get("error_msg") or payload.get("message")
        if not errmsg:
            errmsg = json.dumps(payload, ensure_ascii=False)
        raise RuntimeError(f"DingTalk API error {errcode}: {errmsg}")

    return payload


def list_instance_ids(
    process_code: str,
    start_time: int,
    end_time: int,
    cursor: int = 0,
    size: int = 10,
    userid_list: list[str] | None = None,
) -> dict:
    body: dict = {
        "process_code": process_code,
        "start_time": start_time,
        "end_time": end_time,
        "cursor": cursor,
        "size": size,
    }
    if userid_list:
        body["userid_list"] = userid_list
    return _request("POST", f"{OAPI_BASE}/topapi/processinstance/listids", body)


def get_instance_detail(process_instance_id: str) -> dict:
    body = {"process_instance_id": process_instance_id}
    return _request("POST", f"{OAPI_BASE}/topapi/processinstance/get", body)


def list_all_instances(
    process_code: str,
    start_time: int,
    end_time: int,
    size: int = 10,
    max_instances: int = 10000,
    include_details: bool = False,
    userid_list: list[str] | None = None,
) -> dict:
    all_ids: list[str] = []
    details: list[dict] = []
    cursor = 0
    seen_cursors: set[int] = set()

    while len(all_ids) < max_instances:
        if cursor in seen_cursors:
            break
        seen_cursors.add(cursor)

        page = list_instance_ids(
            process_code=process_code,
            start_time=start_time,
            end_time=end_time,
            cursor=cursor,
            size=min(size, 10, max_instances - len(all_ids)),
            userid_list=userid_list,
        )
        result = page.get("result", {})
        ids = result.get("list", []) or []
        all_ids.extend(ids)

        next_cursor = result.get("next_cursor")
        if next_cursor is None:
            break
        cursor = int(next_cursor)
        if not ids:
            break

    if include_details:
        total = len(all_ids)
        for idx, instance_id in enumerate(all_ids, start=1):
            print(f"Fetching detail {idx}/{total} ...", flush=True)
            details.append(get_instance_detail(instance_id))
            time.sleep(0.15)

    return {
        "process_code": process_code,
        "start_time": start_time,
        "end_time": end_time,
        "count": len(all_ids),
        "ids": all_ids,
        "details": details if include_details else None,
        "truncated": len(all_ids) >= max_instances,
    }


def compute_time_range(year: int, month: int) -> tuple[int, int]:
    start = int(datetime(year, month, 1, 0, 0, 0, tzinfo=TZ).timestamp() * 1000)
    last_day = calendar.monthrange(year, month)[1]
    end = int(datetime(year, month, last_day, 23, 59, 59, 999000, tzinfo=TZ).timestamp() * 1000)
    return start, end


def parse_args():
    parser = argparse.ArgumentParser(
        description="Fetch DingTalk OA approval instances and save raw JSON."
    )
    parser.add_argument(
        "--process-code",
        default="PROC-402A087A-F4A1-4B4D-9726-29BA08FD773D",
        help="DingTalk approval process code. Default: 通用费用报销（人民币） new template.",
    )
    parser.add_argument(
        "--process-name",
        help="Process name key from PROCESS_CODES dict. Alternative to --process-code.",
    )
    parser.add_argument("--year", type=int, required=True, help="Target year (e.g. 2026).")
    parser.add_argument("--month", type=int, required=True, help="Target month (1-12).")
    parser.add_argument(
        "--include-details",
        action="store_true",
        default=True,
        help="Fetch full instance details (default: True).",
    )
    parser.add_argument(
        "--no-details",
        action="store_true",
        help="Only fetch instance IDs, skip detail fetching.",
    )
    parser.add_argument("--max-instances", type=int, default=20000, help="Max instances to fetch.")
    parser.add_argument("--output-dir", default=".", help="Directory to save output JSON.")
    parser.add_argument(
        "--output-name",
        help="Output filename. Default: YYYY_MM_general_expense_new_template_details.json",
    )
    return parser.parse_args()


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    args = parse_args()
    include_details = args.include_details and not args.no_details

    start_time, end_time = compute_time_range(args.year, args.month)

    print(f"Fetching DingTalk approval instances for {args.year}-{args.month:02d} ...")
    print(f"  process_code: {args.process_code}")
    print(f"  start_time: {start_time}")
    print(f"  end_time: {end_time}")
    print(f"  include_details: {include_details}")

    result = list_all_instances(
        process_code=args.process_code,
        start_time=start_time,
        end_time=end_time,
        size=10,
        max_instances=args.max_instances,
        include_details=include_details,
    )

    print(f"Total instances: {result['count']}")
    if result["truncated"]:
        print(f"WARNING: Results truncated at {args.max_instances} instances!")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    output_name = args.output_name or f"{args.year}_{args.month:02d}_general_expense_new_template_details.json"
    output_path = output_dir / output_name

    # Resolve process_code from --process-name if given
    process_code = args.process_code
    if args.process_name:
        if args.process_name in PROCESS_CODES:
            process_code = PROCESS_CODES[args.process_name]
        else:
            print(f"ERROR: Unknown process name '{args.process_name}'. Available: {list(PROCESS_CODES.keys())}")
            sys.exit(2)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Saved to: {output_path}")
    print(f"File size: {output_path.stat().st_size} bytes")


if __name__ == "__main__":
    main()