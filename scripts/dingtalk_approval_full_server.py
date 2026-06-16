import json
import os
import sys
import time
from typing import Any
from urllib.parse import urlencode

import requests


OAPI_BASE = "https://oapi.dingtalk.com"

if hasattr(sys.stdin, "reconfigure"):
    sys.stdin.reconfigure(encoding="utf-8")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

_TOKEN_CACHE: dict[str, Any] = {"access_token": None, "expires_at": 0}


def _write(message: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(message, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _error(request_id: Any, code: int, message: str) -> None:
    _write({"jsonrpc": "2.0", "id": request_id, "error": {"code": code, "message": message}})


def _ok(request_id: Any, result: Any) -> None:
    _write({"jsonrpc": "2.0", "id": request_id, "result": result})


def _tool_text(data: Any) -> dict[str, Any]:
    return {"content": [{"type": "text", "text": json.dumps(data, ensure_ascii=False, indent=2)}]}


def _get_access_token() -> str:
    token = os.environ.get("DINGTALK_ACCESS_TOKEN")
    if token:
        return token

    now = int(time.time())
    cached = _TOKEN_CACHE.get("access_token")
    if cached and int(_TOKEN_CACHE.get("expires_at", 0)) > now + 60:
        return str(cached)

    app_key = os.environ.get("DINGTALK_APP_KEY")
    app_secret = os.environ.get("DINGTALK_APP_SECRET")
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


def _request(method: str, url: str, body: dict[str, Any] | None = None, auth: bool = True) -> dict[str, Any]:
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


def get_manageable_templates(userid: str | None = None, app_uuid: str | None = None) -> dict[str, Any]:
    if not userid:
        userid = os.environ.get("DINGTALK_USER_ID")
    if not userid:
        raise RuntimeError("userid is required. Either pass it or set DINGTALK_USER_ID env var.")
    body: dict[str, Any] = {"userid": userid}
    if app_uuid:
        body["app_uuid"] = app_uuid
    return _request("POST", f"{OAPI_BASE}/topapi/process/template/manage/get", body)


def list_instance_ids(
    process_code: str,
    start_time: int,
    end_time: int,
    cursor: int = 0,
    size: int = 10,
    userid_list: list[str] | None = None,
) -> dict[str, Any]:
    body: dict[str, Any] = {
        "process_code": process_code,
        "start_time": start_time,
        "end_time": end_time,
        "cursor": cursor,
        "size": size,
    }
    if userid_list:
        body["userid_list"] = userid_list
    return _request("POST", f"{OAPI_BASE}/topapi/processinstance/listids", body)


def get_instance_detail(process_instance_id: str) -> dict[str, Any]:
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
) -> dict[str, Any]:
    all_ids: list[str] = []
    details: list[dict[str, Any]] = []
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
        for instance_id in all_ids:
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


TOOLS: dict[str, dict[str, Any]] = {
    "get_manageable_templates": {
        "description": "获取指定管理员 userid 在当前企业内可管理的 OA 审批模板。若不传 userid，则使用环境变量 DINGTALK_USER_ID 的值。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "userid": {"type": "string", "description": "拥有 OA 审批管理权限的钉钉 userid。不传则使用 DINGTALK_USER_ID 环境变量。"},
                "app_uuid": {"type": "string", "description": "可选，审批应用 appUuid。"},
            },
            "required": [],
        },
    },
    "list_instance_ids": {
        "description": "按审批模板和时间范围获取权限范围内的审批实例 ID 列表，支持分页。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "process_code": {"type": "string"},
                "start_time": {"type": "integer", "description": "开始时间，毫秒时间戳。"},
                "end_time": {"type": "integer", "description": "结束时间，毫秒时间戳。"},
                "cursor": {"type": "integer", "default": 0},
                "size": {"type": "integer", "default": 10, "minimum": 1, "maximum": 10},
                "userid_list": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["process_code", "start_time", "end_time"],
        },
    },
    "get_instance_detail": {
        "description": "获取单个 OA 审批实例详情，包括表单字段、发起人、抄送人、状态和流程信息。",
        "inputSchema": {
            "type": "object",
            "properties": {"process_instance_id": {"type": "string"}},
            "required": ["process_instance_id"],
        },
    },
    "list_all_instances": {
        "description": "循环分页拉取某模板在时间范围内的审批实例 ID，可选同时拉取详情。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "process_code": {"type": "string"},
                "start_time": {"type": "integer", "description": "开始时间，毫秒时间戳。"},
                "end_time": {"type": "integer", "description": "结束时间，毫秒时间戳。"},
                "size": {"type": "integer", "default": 10, "minimum": 1, "maximum": 10},
                "max_instances": {"type": "integer", "default": 10000, "minimum": 1},
                "include_details": {"type": "boolean", "default": False},
                "userid_list": {"type": "array", "items": {"type": "string"}},
            },
            "required": ["process_code", "start_time", "end_time"],
        },
    },
}


def _list_tools() -> list[dict[str, Any]]:
    return [
        {
            "name": name,
            "description": spec["description"],
            "inputSchema": spec["inputSchema"],
        }
        for name, spec in TOOLS.items()
    ]


def _call_tool(name: str, arguments: dict[str, Any]) -> Any:
    if name == "get_manageable_templates":
        return get_manageable_templates(**arguments)
    if name == "list_instance_ids":
        return list_instance_ids(**arguments)
    if name == "get_instance_detail":
        return get_instance_detail(**arguments)
    if name == "list_all_instances":
        return list_all_instances(**arguments)
    raise ValueError(f"Unknown tool: {name}")


def main() -> None:
    for line in sys.stdin:
        try:
            request = json.loads(line)
            method = request.get("method")
            request_id = request.get("id")

            if method == "initialize":
                _ok(
                    request_id,
                    {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {"tools": {}},
                        "serverInfo": {"name": "dingtalk_approval_full", "version": "0.1.0"},
                    },
                )
            elif method == "ping":
                _ok(request_id, {})
            elif method == "notifications/initialized":
                continue
            elif method == "tools/list":
                _ok(request_id, {"tools": _list_tools()})
            elif method == "tools/call":
                params = request.get("params", {})
                result = _call_tool(params.get("name"), params.get("arguments") or {})
                _ok(request_id, _tool_text(result))
            else:
                _error(request_id, -32601, f"Method not found: {method}")
        except Exception as exc:
            _error(request.get("id") if "request" in locals() else None, -32000, str(exc))


if __name__ == "__main__":
    main()
