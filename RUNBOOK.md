# Petpooja Daily Discount Dashboard — Runbook

Every night at midnight, Petpooja emails two reports per outlet from
`support@petpooja.com`:

1. **"Report Notification: Item Wise Report With Bill No. : <Outlet>"** — one
   row per line item, with `Invoice No.`, `Sub Total`, `Discount`, `Tax`,
   `Final Total`. Always arrives as an `.xlsx` attachment.
2. **"Report Notification: Payment Wise Order Summary"** (or similarly named,
   exact subject may vary) — one row per bill with a payment-mode / order-type
   column. This is what lets us tell Dine-in apart from Swiggy and Zomato.
   *As of 2026-09-23 this report was not yet enabled in the Petpooja
   Notification tab for the one outlet on this account (Tuskin Coffee Andheri
   West); the dashboard runs without it and marks every order
   "Pending (awaiting Payment Wise report)" until it's turned on.*

This repo holds the scripts that turn those two attachments into the daily
**Petpooja Daily Discount Dashboard** workbook. There is no cron job that
runs headlessly — a Claude session wakes up on a schedule (see "Scheduling"
below) and executes these steps directly using its Gmail/Drive/Sheets tools.

## Daily steps (run at 9am, for the previous day's midnight emails)

1. **Find the emails.** Search Gmail (`from:support@petpooja.com newer_than:1d`)
   for each outlet's Item Wise Report and Payment Wise Order Summary for
   yesterday's date. There may be more than one outlet — process each.

2. **Extract attachments.** Gmail's `get_message`/`get_thread` tools only
   expose attachment *metadata* (id/filename/mime type), not the bytes. Fetch
   the message with `messageFormat: RAW` (this returns a large base64 MIME
   blob and gets saved to a tool-result file automatically) and run:

   ```bash
   python3 scripts/extract_gmail_attachment.py <raw_message.json> <out_dir>
   ```

   This decodes the RFC822 message and writes each attachment to `<out_dir>`.

3. **Build a manifest** (`manifest.json`) listing every outlet found and the
   paths to its two attachments for that date:

   ```json
   {
     "report_date": "2026-09-22",
     "outlets": [
       {"name": "Tuskin Coffee Andheri West",
        "item_wise_xlsx": "/path/Item_bill_report_....xlsx",
        "payment_wise_xlsx": "/path/Payment_wise_....xlsx"}
     ]
   }
   ```

   Set `"payment_wise_xlsx": null` for an outlet whose payment-wise report
   hasn't arrived yet — the dashboard still builds, just without the platform
   split for that outlet.

4. **Build the dashboard:**

   ```bash
   python3 scripts/build_dashboard.py manifest.json Petpooja_Discount_Dashboard.xlsx
   ```

   This produces one sheet per outlet (raw per-order table + formula-driven
   summary + conditional-formatting highlights) plus a `Dashboard` summary
   sheet and a `PaymentMap` sheet holding the raw payment-mode rows. See
   "What the dashboard shows" below.

5. **Publish it:**
   - **Google Sheet** — upload via the Drive connector
     (`contentMimeType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`,
     do *not* set `disableConversionToGoogleType`) so it converts to a native
     Sheet and recalculates. Update the existing daily sheet in place if one
     already exists, don't create a new file every day.
   - **Email** — attach the same `.xlsx` and send it to the outlet mailbox
     (or whichever address the user prefers) with the report date in the
     subject.

6. **Recalculation note.** This sandbox's LibreOffice install currently fails
   to load *any* file (`soffice --convert-to` errors with "source file could
   not be loaded" even for a plain `.txt`), so `scripts/recalc.py` from the
   xlsx skill cannot be used to pre-verify formulas here. This doesn't affect
   the deliverable: Excel and Google Sheets both recalculate on open. If you
   ever need to verify formula output before sending, cross-check the numbers
   independently in Python (aggregate the item-wise rows by invoice and
   compare) rather than relying on a local LO recalc in this environment.

## What the dashboard shows (per outlet)

- **Avg discount %** by platform (Dine-in / Swiggy / Zomato).
- **Orders >50% discount** on Swiggy or Zomato — highlighted orange.
- **100% discount Dine-in orders** — flagged as staff orders, highlighted red.
- **Dine-in discount tiers** — counts of dine-in orders above 15%, 30%, and
  50% discount, highlighted yellow/amber/red on the order table.

Platform is resolved by matching each Item Wise Report's `Invoice No.`
against the Payment Wise Order Summary's bill/invoice + payment-mode columns
(`Zomato` in the mode → Zomato, `Swiggy` → Swiggy, everything else → Dine-in).
Until that second report is enabled, orders show as
"Pending (awaiting Payment Wise report)".

## Scheduling

A daily trigger fires this session at 9am IST (03:30 UTC) with a prompt to
run the steps above for the previous day. Update the trigger
(`update_trigger`) if the wake time or recipients need to change.
