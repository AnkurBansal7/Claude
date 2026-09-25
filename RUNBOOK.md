# Petpooja Daily Discount Dashboard — Runbook

Every night at midnight, Petpooja emails two reports per outlet from
`support@petpooja.com`:

1. **"Report Notification: Item Wise Report With Bill No. : <Outlet>"** — one
   row per line item, with `Invoice No.`, `Sub Total`, `Discount`, `Tax`,
   `Final Total`. Arrives as a real `.xlsx` (OOXML) attachment.
2. **"Report Notification: Payment Wise Summary : <Outlet>"** — one row per
   bill, with `Invoice No.`, `Payment Type` (Online/Cash/Card), `Order Type`
   (`Dine In` or `Delivery(Parcel)`), and `Area` (`Zomato`/`Swiggy`/blank).
   This is what lets us tell Dine-in apart from Swiggy and Zomato. **Gotcha:**
   despite the `.xls` filename, this attachment is actually an HTML table
   (Excel's old "save as HTML, name it .xls" export), not a real binary/OOXML
   file — `scripts/build_dashboard.py` auto-detects this from content (checks
   for the `PK` zip magic bytes) and parses either format. It also has a
   trailing `Total` footer row that must be filtered out (both parsers skip
   any row whose invoice value isn't purely numeric).

This repo holds the scripts that turn those two attachments into the daily
**Petpooja Daily Discount Dashboard** workbook. There is no cron job that
runs headlessly — a Claude session wakes up on a schedule (see "Scheduling"
below) and executes these steps directly using its Gmail/Drive/Sheets tools.

## Daily steps (run at 9am, for the previous day's midnight emails)

1. **Find the emails.** Search Gmail (`from:support@petpooja.com newer_than:1d`)
   for each outlet's Item Wise Report and Payment Wise Summary for yesterday's
   date. There may be more than one outlet — process each. Subject lines seen
   so far: "Report Notification: Item Wise Report With Bill No. : <Outlet>"
   and "Report Notification: Payment Wise Summary : <Outlet>" — search
   broadly (e.g. `subject:"Report Notification"`) since exact wording could
   vary by account/report configuration.

2. **Extract attachments.** Gmail's `get_message`/`get_thread` tools only
   expose attachment *metadata* (id/filename/mime type), not the bytes. Fetch
   the message with `messageFormat: RAW`. For most messages this is large
   enough that the harness auto-saves it to a tool-result file, and you run:

   ```bash
   python3 scripts/extract_gmail_attachment.py <raw_message.json> <out_dir>
   ```

   This decodes the RFC822 message and writes each attachment to `<out_dir>`.

   **Gotcha: borderline-sized messages return inline instead of saving to a
   file.** If a RAW response comes back inline in your own context rather
   than as a "result saved to file" pointer, do **not** try to save that
   inline JSON to a file yourself by retyping/copying it into a Write call —
   this was tried (confirmed with three independent attempts on one message)
   and it silently truncated or corrupted the base64 `raw` field every time,
   at a consistent length, producing a `.xlsx` that fails to open
   (`zipfile.BadZipFile`). This is the same binary-transcription problem as
   step 5, just triggered earlier in the pipeline. Instead, run:

   ```bash
   python3 scripts/extract_from_transcript.py <your_own_transcript.jsonl> <out_dir>
   ```

   This finds the `get_message` tool_result inside the calling agent's own
   `.jsonl` transcript (the harness wrote the exact bytes there — the file
   path is `~/.claude/projects/<project>/subagents/agent-<id>.jsonl` for a
   subagent transcript, or check the session's own log location if this is
   the main session) and extracts the attachment mechanically from that, with
   zero model-generated text in the path. If you delegated the fetch to a
   subagent specifically because of this issue, its `.jsonl` transcript path
   is under `~/.claude/projects/.../subagents/agent-<agentId>.jsonl` — the
   `agentId` is given when you spawn it. Confirmed working (2026-09-24, Tuskin
   Coffee Bandra's item-wise report).

3. **Build a manifest** (`manifest.json`) listing every outlet found and the
   paths to its two attachments for that date:

   ```json
   {
     "report_date": "2026-09-22",
     "outlets": [
       {"name": "Tuskin Coffee Andheri West",
        "item_wise_xlsx": "/path/Item_bill_report_....xlsx",
        "payment_wise_xlsx": "/path/payment_wise_summary_....xls"}
     ]
   }
   ```

   Set `"payment_wise_xlsx": null` for an outlet whose payment-wise report
   hasn't arrived that day — the dashboard still builds, just without the
   platform split for that outlet (orders show as
   "Pending (awaiting Payment Wise report)").

4. **Build the dashboard:**

   ```bash
   python3 scripts/build_dashboard.py manifest.json Petpooja_Discount_Dashboard.xlsx
   ```

   This produces the formula-driven `.xlsx` (one sheet per outlet with the
   raw per-order table, summary formulas, and conditional-formatting
   highlights, plus a `Dashboard` summary sheet and a `PaymentMap` sheet) —
   see "What the dashboard shows" below — **and** a plain-text
   `Petpooja_Discount_Dashboard.csv` snapshot of the same numbers (same base
   name, `.csv` extension). Always use both outputs; see step 5 for why.

5. **Publish it — READ THIS BEFORE ATTACHING ANYTHING:**

   Do not pass the `.xlsx` file's bytes as a `base64Content` (Drive) or
   attachment `content` (Gmail) tool parameter. Verified the hard way while
   building this pipeline: an agent must reproduce that value as a literal
   string in its own output, and verbatim reproduction of an opaque
   high-entropy blob like base64 is **not reliable at any practical size** —
   corruption was observed even in an isolated ~8KB chunk. A file that looks
   fine when spot-checked can still be silently corrupted elsewhere (e.g. in
   unused theme/style XML), so "it opened and looked right" is not proof the
   bytes were transmitted correctly. This will bite every single daily run
   unless you use the reliable channels below:

   - **Google Sheet** — upload the **CSV**, not the xlsx: call
     `mcp__Google_Drive__create_file` with `contentMimeType: "text/csv"` and
     `textContent` set to the *literal* CSV text (plain UTF-8 — safe to
     reproduce, unlike base64). Do not set `base64Content`. Drive
     auto-converts it to a native Sheet. `create_file` cannot update an
     existing file's content (only rename via `update_file`), so create a
     new Sheet each day and `trash_file` the previous day's — keep the title
     identical ("Petpooja Daily Discount Dashboard") so the user finds it the
     same way each time.
   - **Give the human the real workbook** — use the `SendUserFile` tool with
     the local `.xlsx` path. It transfers the file directly through the
     harness, not through the model's text output, so it has none of the
     base64-transcription risk and preserves formulas/formatting/chart. This
     is the right way to hand over the full xlsx — never Gmail.
   - **Email** — send a plain-text summary (platform split, flagged invoice
     numbers, the Google Sheet link) in the message body. Do not attach the
     xlsx (or any binary file) via `mcp__Gmail__send_message` — its
     `attachments[].content` field requires base64, which has the same
     unreliability. If you must email something, attach the CSV's
     already-verified `textContent` re-encoded to base64 only when small
     (a few KB) — always verify by reading back whatever you upload to Drive
     before trusting it, since Drive's `read_file_content` gives a cheap
     correctness check that Gmail's attachment path doesn't.

6. **Recalculation note.** This sandbox's LibreOffice install currently fails
   to load *any* file (`soffice --convert-to` errors with "source file could
   not be loaded" even for a plain `.txt`), so `scripts/recalc.py` from the
   xlsx skill cannot be used to pre-verify formulas here. This doesn't affect
   the deliverable: Excel and Google Sheets both recalculate on open. The CSV
   snapshot's numbers are computed directly in Python (not via the xlsx
   formulas), so cross-check a few against the xlsx's formula cells if you
   change the calculation logic.

## What the dashboard shows (per outlet)

- **Avg discount %** by platform (Dine-in / Swiggy / Zomato).
- **Orders >52% discount** on Swiggy or Zomato — highlighted orange. (Raised
  from >50% on 2026-09-25 per user request; see `ONLINE_HIGH_DISCOUNT_PCT` in
  `scripts/build_dashboard.py`.)
- **100% discount Dine-in orders** — flagged as staff orders, highlighted red.
- **Dine-in discount tiers** — counts of dine-in orders above 15%, 30%, and
  50% discount, highlighted yellow/amber/red on the order table.

Platform is resolved by matching each Item Wise Report's `Invoice No.`
against the Payment Wise Summary's `Order Type` + `Area` columns:
`Order Type` containing "dine" → Dine-in; otherwise `Area` = Zomato/Swiggy
gives the platform directly. Until that second report arrives for a given
day/outlet, orders show as "Pending (awaiting Payment Wise report)".

## Scheduling

A daily trigger fires this session at 9am IST (03:30 UTC) with a prompt to
run the steps above for the previous day. Update the trigger
(`update_trigger`) if the wake time or recipients need to change.
