"""Unified monthly DingTalk expense report runner.

Designed for automation platforms (e.g., OpenClaw / 小龙虾).
Automatically processes the previous month's data, builds reports,
and sends the summary Excel file to designated recipients via DingTalk.

Usage:
    # Auto mode (previous month, for OpenClaw scheduling):
    python run_monthly.py --auto

    # Manual mode (specify year/month):
    python run_monthly.py --year 2026 --month 4

    # Lookup user ID by name:
    python run_monthly.py --lookup-user "金春玲"

Exit codes:
    0 - Success
    1 - Partial failure (report built but notification failed)
    2 - Fatal failure (data fetch or report build failed)
"""

import argparse
import calendar
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urlencode

import requests


# ─── Hard-coded Configuration ───────────────────────────────────────────────────

OAPI_BASE = "https://oapi.dingtalk.com"
API_BASE = "https://api.dingtalk.com"
TZ = timezone(timedelta(hours=8))

# DingTalk app credentials
DEFAULT_APP_KEY = "dingtyzldpxnwvoonzxm"
DEFAULT_APP_SECRET = "L10kCqqStjF8MmMDZCPYbP6LtzHfgU_oEBtGdftwzJaxRtDpvAicQr7m-ScVzA3C"
DEFAULT_USER_ID = "16248445393404993"

# Default recipients for the monthly expense summary
DEFAULT_RECIPIENTS = {
    "金春玲": "16248445393404993",
    "丁红姣": "16208048923325185",
}

# Default process code for 通用费用报销（人民币）新版
DEFAULT_PROCESS_CODE = "PROC-402A087A-F4A1-4B4D-9726-29BA08FD773D"

# Script paths (relative to skill directory)
SKILL_DIR = Path(__file__).resolve().parent
FETCH_SCRIPT = SKILL_DIR / "fetch_approval_data.py"
BUILD_SCRIPT = SKILL_DIR / "build_report.py"

# ─── Token Management ───────────────────────────────────────────────────────────

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

    url = f"{OAPI_BASE}/gettoken?{urlencode({'appkey': app_key, 'appsecret': app_secret})}"
    payload = _oapi_request("GET", url, auth=False)
    access_token = payload.get("access_token")
    if not access_token:
        raise RuntimeError(f"DingTalk gettoken failed: {payload}")

    expires_in = int(payload.get("expires_in", 7200))
    _TOKEN_CACHE["access_token"] = access_token
    _TOKEN_CACHE["expires_at"] = now + expires_in
    return str(access_token)


# ─── DingTalk API Helpers ────────────────────────────────────────────────────────

def _oapi_request(method: str, url: str, body: dict | None = None, auth: bool = True) -> dict:
    """Call DingTalk OAPI (old API)."""
    if auth:
        separator = "&" if "?" in url else "?"
        url = f"{url}{separator}{urlencode({'access_token': _get_access_token()})}"

    response = requests.request(method, url, json=body, timeout=30)
    response.raise_for_status()
    payload = response.json()

    errcode = payload.get("errcode")
    if errcode not in (None, 0):
        errmsg = payload.get("errmsg") or payload.get("error_msg") or payload.get("message")
        raise RuntimeError(f"DingTalk OAPI error {errcode}: {errmsg}")

    return payload


def _api_request(method: str, path: str, body: dict | None = None) -> dict:
    """Call DingTalk new API (api.dingtalk.com) with Bearer token."""
    response = requests.request(
        method,
        f"{API_BASE}{path}",
        json=body,
        headers={"x-acs-dingtalk-access-token": _get_access_token()},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


# ─── File Upload & Notification ──────────────────────────────────────────────────

def upload_file(file_path: str | Path) -> str:
    """Upload a file to DingTalk media and return the media_id."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    with open(file_path, "rb") as f:
        result = requests.post(
            f"{OAPI_BASE}/media/upload?access_token={_get_access_token()}&type=file",
            files={"media": (file_path.name, f, "application/octet-stream")},
            timeout=120,
        ).json()

    errcode = result.get("errcode", 0)
    if errcode != 0:
        raise RuntimeError(f"DingTalk media upload failed: {result}")

    media_id = result.get("media_id")
    if not media_id:
        raise RuntimeError(f"No media_id in upload response: {result}")

    print(f"  File uploaded: {file_path.name} -> media_id={media_id}")
    return media_id


def send_file_message(recipient_ids: list[str], media_id: str, file_name: str, robot_code: str | None = None) -> dict:
    """Send a file message to specified users via DingTalk robot OTT API."""
    robot_code = robot_code or os.environ.get("DINGTALK_APP_KEY", DEFAULT_APP_KEY)

    result = _api_request("POST", "/v1.0/robot/oToMessages/batchSend", {
        "robotCode": robot_code,
        "userIds": recipient_ids,
        "msgKey": "sampleFile",
        "msgParam": json.dumps({
            "mediaId": media_id,
            "fileName": file_name,
        }),
    })

    invalid = result.get("invalidStaffIdList", [])
    if invalid:
        print(f"  WARNING: Invalid user IDs: {invalid}")

    return result


def send_markdown_message(recipient_ids: list[str], title: str, text: str, robot_code: str | None = None) -> dict:
    """Send a markdown message to specified users via DingTalk robot OTT API."""
    robot_code = robot_code or os.environ.get("DINGTALK_APP_KEY", DEFAULT_APP_KEY)

    result = _api_request("POST", "/v1.0/robot/oToMessages/batchSend", {
        "robotCode": robot_code,
        "userIds": recipient_ids,
        "msgKey": "sampleMarkdown",
        "msgParam": json.dumps({
            "title": title,
            "text": text,
        }),
    })

    return result


# ─── User Lookup ──────────────────────────────────────────────────────────────────

def lookup_user_by_name(name: str) -> str | None:
    """Look up a DingTalk user ID by name.

    Since the app may not have通讯录 permissions, this function tries:
    1. Check hard-coded DEFAULT_RECIPIENTS
    2. Search through recent approval data (if available)
    3. Traverse the org structure via API (if permissions allow)
    """
    # 1. Check hard-coded defaults
    if name in DEFAULT_RECIPIENTS:
        print(f"  Found '{name}' in default recipients: {DEFAULT_RECIPIENTS[name]}")
        return DEFAULT_RECIPIENTS[name]

    # 2. Try DingTalk API user search
    try:
        result = _oapi_request("POST", f"{OAPI_BASE}/topapi/v2/user/list", {
            "dept_id": 1, "cursor": 0, "size": 100,
        })
        for user in result.get("result", {}).get("list", []):
            if user.get("name") == name:
                userid = user.get("userid")
                print(f"  Found '{name}' via API: {userid}")
                return userid
    except RuntimeError:
        pass  # Permission denied, skip

    print(f"  WARNING: Could not find user ID for '{name}'")
    return None


# ─── Month Detection ──────────────────────────────────────────────────────────────

def previous_month() -> tuple[int, int]:
    """Return (year, month) for the previous month in Asia/Shanghai timezone."""
    now = datetime.now(TZ)
    first_of_this_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_of_prev = first_of_this_month - timedelta(days=1)
    return last_of_prev.year, last_of_prev.month


# ─── Main Workflow ────────────────────────────────────────────────────────────────

def run_fetch(year: int, month: int, output_dir: Path) -> Path:
    """Step 1: Fetch DingTalk approval data."""
    python = sys.executable
    output_name = f"{year}_{month:02d}_general_expense_new_template_details.json"
    output_path = output_dir / output_name

    cmd = [
        python, str(FETCH_SCRIPT),
        "--year", str(year),
        "--month", str(month),
        "--include-details",
        "--max-instances", "20000",
        "--output-dir", str(output_dir),
        "--output-name", output_name,
    ]

    print(f"[Step 1] Fetching approval data for {year}-{month:02d} ...")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if result.returncode != 0:
        print(f"  FAILED: {result.stderr}")
        raise RuntimeError(f"fetch_approval_data.py failed with exit code {result.returncode}")
    print(result.stdout.rstrip())

    if not output_path.exists():
        raise RuntimeError(f"Expected output file not found: {output_path}")

    return output_path


def run_build(year: int, month: int, details_json: Path, output_dir: Path) -> dict:
    """Step 2: Build report workbooks."""
    python = sys.executable

    cmd = [
        python, str(BUILD_SCRIPT),
        "--details-json", str(details_json),
        "--year", str(year),
        "--month", str(month),
        "--output-dir", str(output_dir),
    ]

    print(f"\n[Step 2] Building report workbooks ...")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if result.returncode != 0:
        print(f"  FAILED: {result.stderr}")
        raise RuntimeError(f"build_report.py failed with exit code {result.returncode}")

    # Parse the JSON output from build_report.py
    build_result = json.loads(result.stdout.strip())
    print(f"  Raw details: {build_result['raw_details']}")
    print(f"  Approved rows: {build_result['approved_rows']}")
    print(f"  Grand total: ¥{build_result['grand_total']:,.2f}")
    print(f"  Summary workbook: {build_result['summary_xlsx']}")

    return build_result


def run_notify(build_result: dict, year: int, month: int, recipient_ids: list[str]) -> bool:
    """Step 3: Upload and send the summary file via DingTalk."""
    summary_path = Path(build_result["summary_xlsx"])
    detail_path = Path(build_result["detail_xlsx"])

    if not summary_path.exists():
        print(f"  WARNING: Summary file not found: {summary_path}")
        return False

    print(f"\n[Step 3] Sending reports via DingTalk ...")
    print(f"  Recipients: {recipient_ids}")

    success = True

    # Send summary Excel
    try:
        media_id = upload_file(summary_path)
        send_file_message(recipient_ids, media_id, summary_path.name)
        print(f"  Sent: {summary_path.name}")
    except Exception as e:
        print(f"  ERROR sending summary file: {e}")
        success = False

    # Send detail Excel
    try:
        if detail_path.exists():
            media_id = upload_file(detail_path)
            send_file_message(recipient_ids, media_id, detail_path.name)
            print(f"  Sent: {detail_path.name}")
    except Exception as e:
        print(f"  ERROR sending detail file: {e}")
        # Non-fatal, detail file is secondary

    # Send markdown summary notification
    try:
        grand_total = build_result.get("grand_total", 0)
        approved_rows = build_result.get("approved_rows", 0)
        raw_details = build_result.get("raw_details", 0)
        dept_count = build_result.get("department_count", 0)
        type_count = build_result.get("expense_type_count", 0)

        md_text = (
            f"## {year}年{month:02d}月费用审批汇总\n\n"
            f"- **审批总条数**: {raw_details}\n"
            f"- **已完成/同意**: {approved_rows}\n"
            f"- **部门数**: {dept_count}\n"
            f"- **费用类型数**: {type_count}\n"
            f"- **含税金额合计**: ¥{grand_total:,.2f}\n\n"
            f"费用汇总表和审批明细表已通过文件消息发送，请查收。"
        )

        send_markdown_message(recipient_ids, f"{year}年{month:02d}月费用审批汇总", md_text)
        print(f"  Sent: markdown summary notification")
    except Exception as e:
        print(f"  ERROR sending markdown notification: {e}")
        success = False

    return success


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Unified monthly DingTalk expense report runner for automation."
    )
    mode_group = parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument("--auto", action="store_true", help="Auto-detect previous month and run full workflow.")
    mode_group.add_argument("--year", type=int, help="Target year (requires --month).")
    mode_group.add_argument("--lookup-user", type=str, help="Look up DingTalk user ID by name.")

    parser.add_argument("--month", type=int, help="Target month (1-12, requires --year).")
    parser.add_argument(
        "--output-dir",
        default=".",
        help="Directory to save output files. Default: current directory.",
    )
    parser.add_argument(
        "--recipients",
        nargs="*",
        help="Override default recipient names. Default: 金春玲 丁红姣",
    )
    parser.add_argument(
        "--skip-notify",
        action="store_true",
        help="Skip DingTalk notification step (only fetch and build).",
    )

    args = parser.parse_args()

    # ── Lookup user mode ──
    if args.lookup_user:
        userid = lookup_user_by_name(args.lookup_user)
        if userid:
            print(f"User ID for '{args.lookup_user}': {userid}")
            sys.exit(0)
        else:
            print(f"Could not find user '{args.lookup_user}'")
            sys.exit(2)

    # ── Determine year/month ──
    if args.auto:
        year, month = previous_month()
        print(f"Auto-detected previous month: {year}-{month:02d}")
    elif args.year and args.month:
        year, month = args.year, args.month
    else:
        parser.error("--year requires --month, or use --auto")
        return

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # ── Resolve recipient IDs ──
    recipient_names = args.recipients or list(DEFAULT_RECIPIENTS.keys())
    recipient_ids = []
    for name in recipient_names:
        userid = lookup_user_by_name(name)
        if userid:
            recipient_ids.append(userid)
        else:
            print(f"  WARNING: Skipping unknown recipient '{name}'")

    if not recipient_ids and not args.skip_notify:
        print("ERROR: No valid recipients found and --skip-notify not set")
        sys.exit(2)

    # ── Step 1: Fetch data ──
    try:
        details_json = run_fetch(year, month, output_dir)
    except Exception as e:
        print(f"FATAL: Data fetch failed: {e}")
        sys.exit(2)

    # ── Step 2: Build reports ──
    try:
        build_result = run_build(year, month, details_json, output_dir)
    except Exception as e:
        print(f"FATAL: Report build failed: {e}")
        sys.exit(2)

    # ── Step 3: Send notifications ──
    if args.skip_notify:
        print(f"\n[Step 3] Skipped (--skip-notify)")
        print(f"\nDone. Reports saved to: {output_dir}")
        sys.exit(0)

    notify_ok = run_notify(build_result, year, month, recipient_ids)

    # ── Final summary ──
    print(f"\n{'='*60}")
    print(f"  Month: {year}-{month:02d}")
    print(f"  Raw details: {build_result['raw_details']}")
    print(f"  Approved: {build_result['approved_rows']}")
    print(f"  Grand total: ¥{build_result['grand_total']:,.2f}")
    print(f"  Summary: {build_result['summary_xlsx']}")
    print(f"  Detail: {build_result['detail_xlsx']}")
    print(f"  Notification: {'OK' if notify_ok else 'FAILED'}")
    print(f"{'='*60}")

    sys.exit(0 if notify_ok else 1)


if __name__ == "__main__":
    main()