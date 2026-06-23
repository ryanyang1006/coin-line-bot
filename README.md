# LINE Excel quote automation framework

This framework updates an existing Excel workbook, recalculates formulas through Microsoft Excel, screenshots a fixed range, and exposes both a CLI and an optional HTTP API for n8n / LINE Bot integration.

## Workbook assumptions

- Quote input: `報價單!B3`
- History sheet: `歷史報價`
- History date column: `A`
- History price column: `C`
- History formulas to pull down: `D:E`
- Screenshot source: `對外報價單2!B2:E38`

These are configurable in `config.example.json`.

## Setup

Use a Windows machine with Microsoft Excel installed.

```powershell
cd C:\Users\cc\Documents\Codex\2026-06-05\line-ex-t-sp-2500-today\outputs\line_excel_bot_framework
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item config.example.json config.json
```

Edit `config.json` if the workbook path or screenshot range changes.

## Run from command line

```powershell
python excel_quote_bot.py --price 2500 --date 2026-06-05
```

The command prints JSON:

```json
{
  "ok": true,
  "price": 2500.0,
  "quote_cell": "報價單!B3",
  "history_cell": "歷史報價!C93",
  "image_path": "..."
}
```

## Run as an API for n8n

```powershell
uvicorn excel_quote_bot:app --host 127.0.0.1 --port 8088
```

POST JSON to:

```text
http://127.0.0.1:8088/quote
```

Example body:

```json
{
  "price": 2500,
  "date": "2026-06-05"
}
```

To return the latest existing quote image for today without recalculating Excel, send:

```json
{
  "command": "tsp"
}
```

The API searches `outputs/generated_quotes` for files like `quote_YYYYMMDD_HHMMSS.png` and returns the newest matching `image_filename`.

In n8n, LINE webhook can parse `t sp 2500`, call this API, then upload or serve the returned `image_path` and reply through LINE Messaging API.

## Link LINE user to review account

The same API service also exposes:

```text
POST http://127.0.0.1:8088/line-user/link
```

Example body:

```json
{
  "body": {
    "text": "Line名稱 : 楊仁薰\n送評帳號ID : 1234567890",
    "line_user_id": "Line_UserId"
  }
}
```

The flat body and older expanded body are also accepted:

```json
{
  "line_name": "LINE顯示名稱",
  "review_account_id": "0919695288",
  "line_user_id": "Uxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
}
```

It checks `網頁資料_設計版.xlsx` sheet `客戶資料` column `B`. If the review account ID is found, it writes the LINE display name to column `H` and LINE userId to column `I` on the same row.

The command-line version is:

```powershell
python line_user_linker.py --line-name "LINE顯示名稱" --review-account-id "0919695288" --line-user-id "Uxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
```

## Export PCGS history for a LINE user

The same API service also exposes:

```text
POST http://127.0.0.1:8088/pcgs-history
```

Example body from n8n:

```json
{
  "body": {
    "command": "pcgs_history U35412679ed3d78ab44c7c2fd61350452"
  }
}
```

You can also pass the LINE user id directly:

```json
{
  "body": {
    "Line_UserId": "U35412679ed3d78ab44c7c2fd61350452"
  }
}
```

The API looks up `客戶資料` column `I`, reads the matching customer name from column `A`, filters `送評紀錄` column `D` by that customer name and column `I` by `未取回`, then exports the visible result table to `outputs/pcgs_history`.

Use the returned `image_filename` to fetch the PNG:

```text
GET http://127.0.0.1:8088/images/pcgs-history/{image_filename}
```

If there are no matching unreturned rows, the response contains:

```json
{
  "ok": true,
  "has_history": false,
  "message": "no history found"
}
```

## Notes

- The script uses Microsoft Excel via COM automation, so it is intended for Windows + Excel, not Linux Docker.
- If multiple requests may arrive at the same time, keep this API as a single local service. It has an in-process lock to prevent two jobs from controlling Excel at once.
- LINE image replies need public HTTPS URLs. The local `image_path` should be uploaded to S3, Cloudflare R2, Google Drive public link, or served behind a tunnel before LINE replies with it.
