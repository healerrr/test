---
name: monthly-dingtalk-expense-summary
description: Generate reusable monthly DingTalk OA expense approval reports from 4 approval processes (通用费用报销、差旅费用报销、经办付款人民币、经办付款外币). Use when the user asks to fetch monthly DingTalk approval details, build completed/agreed expense approval detail workbooks, produce department-by-expense-type summaries, or run the automated monthly workflow with DingTalk notification.
---

# Monthly DingTalk Expense Summary

## Overview

This skill automates the monthly DingTalk expense workflow across **4 approval processes**:

1. **通用费用报销（人民币）** — General expense reimbursement
2. **差旅费用报销(人民币版)** — Travel expense reimbursement
3. **经办付款申请单（人民币版）** — RMB payment request
4. **经办付款申请单（外币版）** — Foreign currency payment request

Workflow steps:

1. Fetch full OA approval details for each of the 4 processes via Python script.
2. Save the raw details JSON for each process separately.
3. Build **independent** report workbooks for each process type (2 xlsx per process = 8 xlsx total).
4. Upload all Excel files and send them to recipients via DingTalk robot OTT messages.

## Bundled Resources

- **`scripts/run_monthly.py`** — Unified entry point for automation (OpenClaw / 小龙虾). Auto-detects previous month, fetches data from all 4 processes, builds reports, and sends DingTalk notifications.
- **`scripts/fetch_approval_data.py`** — Standalone CLI script for fetching DingTalk approval data.
- **`scripts/build_report.py`** — Deterministic report builder with multi-process field mapping.
- **`assets/default_expense_summary_template.xlsx`** — Default Excel layout template.

## Prerequisites

The `requests` and `openpyxl` Python packages must be available. Install into the managed venv if missing:

```powershell
&C:\Users\Thinkpad User\.workbuddy\binaries\python\envs\default\Scripts\pip.exe install requests openpyxl
```

DingTalk credentials are hard-coded as defaults in the scripts (no need to set environment variables). Environment variables can still override the hard-coded values if needed:

```text
DINGTALK_APP_KEY      - overrides built-in default
DINGTALK_APP_SECRET   - overrides built-in default
DINGTALK_USER_ID      - overrides built-in default
DINGTALK_ACCESS_TOKEN - If set, skip app_key/app_secret token exchange entirely
```

## Quick Start (for OpenClaw / 小龙虾)

### One-command automation

```powershell
python "C:\Users\Thinkpad User\.workbuddy\skills\monthly-dingtalk-expense-summary\scripts\run_monthly.py" --auto --output-dir "C:\Users\Thinkpad User\WorkBuddy\expense_reports"
```

This single command:

1. **Auto-detects previous month** (e.g., running in June → processes May)
2. **Fetches** all approval instances with details from 4 DingTalk OpenAPI processes
3. **Builds** 8 independent Excel workbooks (2 per process: 费用汇总表 + 审批明细表)
4. **Sends** all 8 Excel files + a combined markdown summary notification to the default recipients via DingTalk

### OpenClaw scheduling

Configure OpenClaw to trigger this command once per month (e.g., on the 1st of each month at 09:00). The `--auto` flag ensures it always processes the previous month's data automatically.

**Exit codes:**

| Code | Meaning |
|------|---------|
| 0 | Success — reports built and notifications sent |
| 1 | Partial success — reports built but notification failed |
| 2 | Fatal failure — data fetch or report build failed |

### Supported Processes

| Process Name | Process Code |
|-------------|-------------|
| 通用费用报销（人民币） | `PROC-402A087A-F4A1-4B4D-9726-29BA08FD773D` |
| 差旅费用报销(人民币版) | `PROC-4E9EF26C-F477-4641-A103-CDC573812CC7` |
| 经办付款申请单（人民币版） | `PROC-3JYJ9N2V-6AYV91D7SRD9DSUF6QLW1-504EJ9IJ-1` |
| 经办付款申请单（外币版） | `PROC-RIYJS65W-8CSWSZ9SSFAXV8GGN8BY1-5FMPCIJJ-91` |

To fetch only specific processes, use `--processes`:

```powershell
python run_monthly.py --auto --processes "通用费用报销（人民币）" "差旅费用报销(人民币版)"
```

### Default recipients

The reports are sent to the following DingTalk users via robot OTT file messages:

| Name | User ID |
|------|---------|
| 金春玲 | `16248445393404993` |
| 丁红姣 | `16208048923325185` |

To override recipients, use `--recipients`:

```powershell
python run_monthly.py --auto --recipients "张三" "李四"
```

To skip the notification step (only fetch and build):

```powershell
python run_monthly.py --auto --skip-notify
```

### User ID lookup

To look up a DingTalk user ID by name:

```powershell
python run_monthly.py --lookup-user "金春玲"
```

## Manual Usage (Step by Step)

### 1. Fetch details

```powershell
python "C:\Users\Thinkpad User\.workbuddy\skills\monthly-dingtalk-expense-summary\scripts\fetch_approval_data.py" `
  --year 2026 --month 6 `
  --include-details `
  --output-dir "." `
  --max-instances 20000
```

The script outputs a JSON file named by default:

```text
YYYY_MM_general_expense_new_template_details.json
```

Process code defaults to `PROC-402A087A-F4A1-4B4D-9726-29BA08FD773D` (通用费用报销（人民币）新版). Override with `--process-code` or `--process-name` if needed.

### 2. Build report workbooks

```powershell
python "C:\Users\Thinkpad User\.workbuddy\skills\monthly-dingtalk-expense-summary\scripts\build_report.py" `
  --details-json ".\2026_06_general_expense_new_template_details.json" `
  --year 2026 --month 6 --output-dir "."
```

For multiple data sources, pass comma-separated JSON files:

```powershell
python build_report.py `
  --details-json "file1.json,file2.json,file3.json" `
  --year 2026 --month 6 --output-dir "."
```

Outputs per process (e.g., `YYYY_MM_通用费用报销`):

- `<YYYY_MM>_<流程名>_费用汇总表.xlsx`
- `<YYYY_MM>_<流程名>_审批明细表.xlsx`
- `<YYYY_MM>_<流程名>_审批明细表.json`
- `<YYYY_MM>_<流程名>_审批明细表.csv`

Total: 4 processes × 4 files = 16 files (8 xlsx + 4 csv + 4 json).

The flow labels used in filenames are:

| Process Name | File Label |
|-------------|-----------|
| 通用费用报销（人民币） | 通用费用报销 |
| 差旅费用报销(人民币版) | 差旅费用报销 |
| 经办付款申请单（人民币版） | 经办付款人民币 |
| 经办付款申请单（外币版） | 经办付款外币 |

Default metric is 含税金额, sourced from `付款合计金额`. Use `--metric pre-tax` only when the user explicitly wants 不含税金额.

### 3. Verify before replying

After script execution, report:

- number of raw details
- number of `已完成/同意` records
- final summary workbook path
- grand total from the workbook

Do not overwrite the raw DingTalk JSON unless the user asks.

## DingTalk Notification Details

The notification step uses:

1. **File upload**: `POST /media/upload` → gets `media_id`
2. **Robot OTT file message**: `POST /v1.0/robot/oToMessages/batchSend` with `msgKey=sampleFile`
3. **Markdown summary**: `POST /v1.0/robot/oToMessages/batchSend` with `msgKey=sampleMarkdown`

The `robotCode` used is the app's `app_key` (`dingtyzldpxnwvoonzxm`). The app must have robot OTT message permissions enabled.

## Field Rules

The approved-detail workbook columns are:

```text
序号, 审批单号, 发起日期, 发起人, 审批流程, 费用所属企业, 费用所属部门, 费用类型, 不含税金额, 付款合计金额, 费用事由, 审批状态
```

### Multi-Process Field Mapping

Each process type has different form field names. The build script auto-detects the process type and maps fields accordingly:

| Canonical Field | 通用费用报销 | 差旅费用报销 | 经办付款(人民币) | 经办付款(外币) |
|----------------|------------|------------|---------------|--------------|
| 费用所属企业 | 费用所属企业 | 费用所属企业 | 付款所属企业 | (Remmit from)付款所属企业 |
| 费用所属部门 | 费用所属部门 | 资金预算和费用所属部门 | 资金预算(和费用)所属部门 | (Department)资金预算(和费用)所属部门 |
| 不含税金额 | 不含税金额（普票填含税） | 不含税金额（普票填含税） | — | — |
| 付款合计金额 | 费用付款合计金额 | 合计金额 | 付款合计金额 | (Total amount)付款合计金额 |
| 费用类型 | 费用类型 | 科目 | 付款类型 | (Payment type)付款类型 |
| 费用事由 | 费用事由 | 备注 | 付款明细 | (Remarks)备注 |

Process type detection rules (applied in order):

1. Fields contain `(Remmit from)付款所属企业` or `(Payment type)付款类型` → 经办付款外币
2. Fields contain `付款类型` and `付款明细` → 经办付款人民币
3. Fields contain `科目` and `合计金额` → 差旅费用报销
4. Default → 通用费用报销

### Common mapping:

- `审批单号`: `process_instance.business_id`
- `发起日期`: `process_instance.create_time`
- `发起人`: parse `process_instance.title` as `(.+?)提交的`; fallback to start operation userid
- `审批流程`: auto-detected process type label
- `审批状态`: combine DingTalk status/result, e.g. `COMPLETED + agree` => `已完成/同意`

Filter detail rows to:

```text
审批状态 == 已完成/同意
```

## Template Layout

By default, copy the bundled template workbook's first sheet and preserve:

- title merge and formatting
- column widths and row heights
- header, total-row, total-column styles
- department order and expense-type order

Replace title text from `不含税金额` to `含税金额` when using `--metric tax-included`.

If a later month contains a new department or expense type, append it before the total row/column and copy nearby styles.
