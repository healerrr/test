# DingTalk Expense Approval Reference

## Process Codes

Use by default:

| Name | process_code |
| --- | --- |
| 通用费用报销（人民币） | `PROC-402A087A-F4A1-4B4D-9726-29BA08FD773D` |

Old template, only when explicitly requested:

| Name | process_code |
| --- | --- |
| 通用费用报销（人民币）旧版 | `PROC-62A2FF1F-0159-4450-AD2B-2E58ECC02747` |

## Known May 2026 Baseline

These values are useful sanity checks, not hard requirements for other months:

- May 2026 new-template full instance IDs: `279`
- May 2026 full details fetched: `279`
- May 2026 `已完成/同意` rows: `236`
- May 2026 tax-included grand total: `6,940,053.31`

## Timestamp Helper

Use Asia/Shanghai calendar boundaries. Example for June 2026:

- Start: `2026-06-01 00:00:00.000 +08:00`
- End: `2026-06-30 23:59:59.999 +08:00`

In Python:

```python
from datetime import datetime, timezone, timedelta
tz = timezone(timedelta(hours=8))
start = int(datetime(2026, 6, 1, tzinfo=tz).timestamp() * 1000)
end = int(datetime(2026, 7, 1, tzinfo=tz).timestamp() * 1000) - 1
```
