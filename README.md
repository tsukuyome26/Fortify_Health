# Fortify Health -- Compliance Email System v3.1

## Production Guide

---

## Table of Contents

1. [Overview](#overview)
2. [Installation](#installation)
3. [First Run](#first-run)
4. [Menu Reference](#menu-reference)
5. [Sheet Reference](#sheet-reference)
6. [Settings Reference](#settings-reference)
7. [Scheduling Triggers](#scheduling-triggers)
8. [Operational Procedures](#operational-procedures)
9. [Troubleshooting](#troubleshooting)
10. [Technical Notes](#technical-notes)

---

## Overview

This system automates compliance email communication for the Fortify Health wheat flour fortification programme. It reads mill compliance data from a Google Sheet, validates it, generates tier-appropriate emails (HIGH/MEDIUM/LOW), and either creates Gmail drafts or sends them directly.

**What it does:**

- Reads mill data from the Mill_Compliance_Data sheet
- Validates every row (production volume, premix usage, lab results, compliance percentages)
- Looks up contacts from Mill_Contacts and field managers from Field_Managers
- Generates customised emails based on risk tier (HIGH = urgent, MEDIUM = advisory, LOW = check-in or recognition)
- Creates Gmail drafts or sends emails directly depending on configuration
- Logs every action to Draft_Log
- Supports scheduled automatic runs via time-based triggers
- Sends escalation alerts and weekly digest reports

---

## Installation

### Step 1: Upload the Spreadsheet

1. Download the `Fortify_Health_System_v31.xlsx` file.
2. Go to Google Drive.
3. Right-click the file and select **Open with > Google Sheets**.
4. Alternatively, open Google Sheets and go to File > Import > Upload.

### Step 2: Install the Apps Script

1. In the Google Sheet, go to **Extensions > Apps Script**.
2. Delete all default code in the editor (select all, delete).
3. Open the `Code.gs.js` file in a text editor.
4. Copy the entire contents.
5. Paste into the Apps Script editor.
6. Click the floppy disk icon or press Ctrl+S to save.
7. Name the project "Fortify Health v3.1" when prompted.

### Step 3: Reload and Authorize

1. Close the Apps Script editor tab.
2. Close the Google Sheet tab.
3. Reopen the Google Sheet.
4. You should see a **Fortify Health** menu in the toolbar (may take a few seconds).
5. Click **Fortify Health > Draft compliance emails**.
6. A permissions dialog will appear. Click **Review Permissions**.
7. Select your Google account.
8. Click **Advanced > Go to Fortify Health v3.1 (unsafe)**.
9. Click **Allow**.

The system is now installed and authorised.

---

## First Run

### Before your first run, verify:

1. **Mill_Compliance_Data** has data with all required columns.
2. **Mill_Contacts** has a matching entry for every Mill ID in the data sheet.
3. **Field_Managers** has an entry for every State referenced in the data.
4. **Email Flag column** contains "REVIEW" (or is blank) for rows you want to process.

### Running:

1. Click **Fortify Health > Draft compliance emails**.
2. The script processes each row and creates Gmail drafts.
3. When complete, a dialog shows: `Drafted: X | Sent: 0 | Skipped: Y`.
4. Open Gmail to review the drafts.
5. When satisfied, click **Fortify Health > Send all pending drafts** to send them.

### Important: Understanding the Email Flag

The Email Flag column controls whether a row gets processed:

| Flag Value | What Happens |
|---|---|
| REVIEW | Row is processed (email created) |
| (blank/empty) | Row is processed (email created) |
| Any other text | Row is processed (email created) |
| DRAFTED - date | Row is SKIPPED (already drafted) |
| SENT - date | Row is SKIPPED (already sent) |
| DATA ERROR - reason | Row is SKIPPED (has a validation error) |

**If you want to re-process a row**, change its Email Flag back to "REVIEW" or clear it.

**If you want to re-process ALL rows**, use **Fortify Health > Data Tools > Reset all email flags to REVIEW**.

---

## Menu Reference

### Fortify Health (top-level)

| Menu Item | What It Does |
|---|---|
| Draft compliance emails | Creates Gmail drafts for all unprocessed rows |
| Auto-send compliance emails | Sends emails directly (no drafts) for all unprocessed rows |
| Send all pending drafts | Finds all Fortify Health drafts in Gmail and sends them |

### Schedule Triggers

| Menu Item | What It Does |
|---|---|
| Custom date and time trigger... | Opens a dialog to create any trigger with exact date/time/frequency |
| Weekly HIGH tier (Mon 09:00) | Creates a weekly trigger that drafts HIGH tier emails every Monday at 09:00 |
| Bi-weekly MEDIUM (Wed 14:00) | Creates a weekly trigger for MEDIUM tier every Wednesday at 14:00 |
| Monthly LOW auto-send (1st 10:00) | Auto-sends LOW tier emails on the 1st of each month at 10:00 |
| Daily auto-send (09:00) | Auto-sends ALL tiers daily at 09:00 |
| Escalation check (Fri 17:00) | Checks for mills that have been HIGH 3+ times and alerts admin |
| Weekly digest (Fri 18:00) | Sends a programme summary email to admin every Friday |
| View all active triggers | Shows a list of all currently scheduled triggers |
| Remove ALL triggers | Deletes all scheduled triggers |

### Reports

| Menu Item | What It Does |
|---|---|
| Check escalations now | Immediately checks for persistent HIGH-tier mills and sends alert |
| Send weekly digest now | Immediately sends the weekly summary email |
| Generate mill report | Creates a text report in a new _Report sheet |
| Export log to CSV | Exports the Draft_Log sheet as a CSV file to Google Drive |

### Data Tools

| Menu Item | What It Does |
|---|---|
| Recalculate all drift labels | Updates the Drift Label column for every row based on M1/M2/M3 |
| Reset all email flags to REVIEW | Sets ALL Email Flag values back to REVIEW (allows re-processing) |
| Validate all rows (dry run) | Checks every row for data errors without creating any emails |
| Count mills by tier | Shows a count of HIGH/MEDIUM/LOW/Other mills |
| Find missing contacts | Lists mills that have no matching entry in Mill_Contacts |

### One-Time Send

| Menu Item | What It Does |
|---|---|
| Send to specific mill... | Prompts for a Mill ID, then sends or drafts for that single mill |
| Send to specific tier... | Prompts for a tier, then processes all mills of that tier |
| Schedule one-time at date/time... | Prompts for a future date/time and schedules a one-time run |

### Other

| Menu Item | What It Does |
|---|---|
| Reset batch checkpoint | Clears the saved progress so the next run starts from the beginning |

---

## Sheet Reference

### Mill_Compliance_Data

The main data sheet. Required columns:

| Column | Description | Required |
|---|---|---|
| Mill ID | Unique identifier (e.g. FH-MH-001) | Yes |
| Mill Name | Full name of the mill | Yes |
| State | Indian state (must match Field_Managers) | Yes |
| Reporting Month | e.g. "March 2026" | Yes |
| Risk Tier | HIGH, MEDIUM, or LOW | Yes |
| M1 Compliance pct | Most recent month compliance % | Yes |
| M2 Compliance pct | Previous month | Yes |
| M3 Compliance pct | Two months ago | Yes |
| Lab Result ppm Iron | Lab test result in ppm | Yes |
| Lab Variance pct | Variance from FSSAI standard | Yes |
| Production Volume MT | Monthly production in metric tonnes | Yes |
| Premix Used kg | Actual premix used in kg | Yes |
| Expected Premix kg | Expected premix based on production | Yes |
| Premix Deviation pct | Percentage deviation | Yes |
| Days Since Audit | Days since last verified audit | Yes |
| Email Flag | Processing status (REVIEW/DRAFTED/SENT/DATA ERROR) | Yes |
| Drift Label | Calculated trend label (optional, auto-filled) | No |
| Follow-up Due | Auto-calculated follow-up date | No |

### Mill_Contacts

| Column | Description |
|---|---|
| Mill ID | Must match Mill_Compliance_Data exactly |
| Contact Name | Name used in email greeting |
| Email | Recipient email address |
| Phone | Phone number (for reference) |
| Role | Contact's role at the mill |

### Field_Managers

| Column | Description |
|---|---|
| State | Must match State values in Mill_Compliance_Data exactly |
| Manager Name | Used to sign emails |
| Email | CC'd on emails if CC_MANAGER is TRUE |

### Draft_Log

Automatically populated. Each row records one email action with timestamp, mill details, tier, subject preview, field manager, contact email, follow-up date, drift trend, and status.

### System_Settings

Configuration values the script reads at runtime. Edit column B to change behaviour without modifying code.

---

## Settings Reference

All settings are in the System_Settings sheet, column B. The script reads them at the start of each run.

| Setting | Default | Description |
|---|---|---|
| GMAIL_DAILY_LIMIT | 1400 | Stop if remaining quota falls below 10. Google Workspace allows 1500/day; we use 1400 as a safety margin. |
| BATCH_SIZE | 20 | Number of emails per batch. After this many, the script pauses and saves a checkpoint. Run again to continue. |
| FSSAI_PPM_STD | 40 | The iron ppm standard referenced in email text. |
| AUTO_SEND | FALSE | When TRUE, "Draft compliance emails" sends immediately instead of creating drafts. |
| CC_MANAGER | TRUE | When TRUE, the field manager for the mill's state is CC'd on every email. |
| ADMIN_BCC | FALSE | When TRUE, the admin email is BCC'd on every email. |
| PREMIX_RATIO | 0.15 | Expected premix-to-production ratio. Rows with ratio > 3x this value fail validation. |
| MAX_RETRIES | 3 | Number of retry attempts for transient Gmail errors. |
| SEND_JITTER_MS | 1500 | Maximum random delay (ms) between emails to avoid rate limiting. |
| ADMIN_EMAIL | admin@fortifyhealth.in | Receives escalation alerts and weekly digests. |
| HIGH_FOLLOWUP_DAYS | 3 | Follow-up deadline for HIGH tier mills. |
| MEDIUM_FOLLOWUP_DAYS | 7 | Follow-up deadline for MEDIUM tier mills. |
| LOW_FOLLOWUP_DAYS | 30 | Follow-up deadline for LOW tier mills. |

To change a setting: edit the value in column B of the System_Settings sheet. Changes take effect on the next run (no code changes or reloads needed).

---

## Scheduling Triggers

### Using the Custom Trigger Dialog

1. Click **Fortify Health > Schedule Triggers > Custom date and time trigger...**
2. A dialog opens with these fields:
   - **Action**: Which function to run (draft all, auto-send all, draft/send by tier, escalation check, digest)
   - **Frequency**: One-time, daily, weekly, or monthly
   - **Date**: For one-time triggers, select the exact date
   - **Day of week**: For weekly triggers, select the day
   - **Day of month**: For monthly triggers, select 1-28
   - **Hour**: 0-23 (in your Apps Script timezone)
   - **Minute**: 0-59
3. Click **Create Trigger**.
4. The dialog confirms success or shows an error.

### Using Preset Triggers

Click any preset in the Schedule Triggers submenu. Each one creates a single trigger. If a trigger for the same function already exists, it is replaced (not duplicated).

### Viewing and Managing Triggers

- **View all active triggers**: Shows function name and event type for every trigger.
- **Remove ALL triggers**: Deletes every trigger. Use this before reinstalling or if triggers are misbehaving.

### Timezone

Triggers run in the timezone set in your Apps Script project:

1. Open Extensions > Apps Script.
2. Click the gear icon (Project Settings) in the left sidebar.
3. Check "Script timezone" and set it to your local timezone (e.g. Asia/Kolkata).

---

## Operational Procedures

### Daily Operation (Manual)

1. Update Mill_Compliance_Data with the latest monthly data.
2. Set Email Flag to "REVIEW" for new or updated rows.
3. Click **Fortify Health > Draft compliance emails**.
4. Open Gmail, review the drafts.
5. Click **Fortify Health > Send all pending drafts**.
6. Check the Draft_Log sheet for a record of what was sent.

### Daily Operation (Automated)

1. Set up a daily trigger: **Fortify Health > Schedule Triggers > Daily auto-send (09:00)**.
2. Update Mill_Compliance_Data with new data and set flags to "REVIEW".
3. The system sends emails automatically at 09:00 each day.
4. Check Draft_Log periodically for errors.

### Monthly Cycle

1. At month-end, update all mill data for the new reporting month.
2. Run **Data Tools > Reset all email flags to REVIEW** (this clears all flags for re-processing).
3. Run **Data Tools > Recalculate all drift labels** (updates trends).
4. Run **Data Tools > Validate all rows (dry run)** to catch data errors before emailing.
5. Fix any errors flagged in the validation report.
6. Run **Draft compliance emails** or let the scheduled trigger handle it.

### Handling Re-runs

If you need to re-send emails for specific mills:

- For one mill: **One-Time Send > Send to specific mill...** (enter the Mill ID)
- For one tier: **One-Time Send > Send to specific tier...** (enter HIGH/MEDIUM/LOW)
- For all mills: **Data Tools > Reset all email flags to REVIEW**, then run again

### Handling Errors

If the summary shows skipped rows or errors:

1. Check the Email Flag column in Mill_Compliance_Data for "DATA ERROR" entries.
2. The error reason is written in the flag (e.g. "DATA ERROR - Production volume missing/zero").
3. Fix the data in the row.
4. Change the Email Flag back to "REVIEW".
5. Run again.

---

## Troubleshooting

### "Drafted: 0 | Sent: 0 | Skipped: N"

**Cause**: All rows have flags starting with DRAFTED, SENT, or DATA ERROR from a previous run.

**Fix**: Run **Data Tools > Reset all email flags to REVIEW**, then run again.

### "Date must be in the future"

**Cause**: The date/time entered in the trigger dialog is in the past relative to the server timezone.

**Fix**: Ensure your selected date and time are in the future. Check your Apps Script project timezone matches your local timezone (Project Settings > Script timezone).

### "Column not found" error

**Cause**: A required column header in Mill_Compliance_Data does not match the expected name exactly.

**Fix**: Check that all column headers match the names listed in the Sheet Reference section above. Spelling, spaces, and capitalisation must be exact.

### "No contact found" for a mill

**Cause**: The Mill ID in Mill_Compliance_Data does not have a matching entry in Mill_Contacts.

**Fix**: Add the mill's contact to the Mill_Contacts sheet. The Mill ID must match exactly (including hyphens and case). Use **Data Tools > Find missing contacts** to see all missing entries.

### Gmail quota errors

**Cause**: Google Workspace limits to 1500 emails per day. The script stops when quota is below 10.

**Fix**: Wait until the next day. The checkpoint system remembers where it stopped. Run again to continue from where it left off.

### Batch paused message

**Cause**: The script processed BATCH_SIZE emails and paused to avoid timeout.

**Fix**: This is normal. Run the same command again. The script continues from the checkpoint. Repeat until you see "Run complete."

### Script takes too long

**Cause**: Google Apps Script has a 6-minute execution limit.

**Fix**: Reduce BATCH_SIZE in System_Settings (e.g. from 20 to 10). The script will pause more frequently but avoid timeouts.

### Triggers not firing

**Cause**: Various. Common issues include incorrect timezone, expired authorization, or too many triggers.

**Fix**:
1. Check timezone in Project Settings.
2. Run any menu item manually to re-authorize if needed.
3. Use **View all active triggers** to confirm triggers exist.
4. Use **Remove ALL triggers** and recreate them if needed.

---

## Technical Notes

### Index Logic

The script carefully maps between array indices and sheet rows:

- `data[0]` = header row = sheet row 1
- `data[i]` = record = sheet row (i + 1)
- To read sheet cell for data[i]: `getRange(i + 1, columnIndex + 1)`

This is documented at the top of the script to prevent off-by-one bugs.

### Email Flag Live Read

The Email Flag is read directly from the sheet cell (not from the cached data array) to prevent duplicate processing if the script is run concurrently or if flags change during a run.

### Concurrency Lock

The script uses `LockService.getScriptLock()` to prevent two instances from running simultaneously. If a trigger fires while a manual run is in progress, the trigger run will exit gracefully.

### Checkpoint System

After processing BATCH_SIZE emails, the script saves the current position to `ScriptProperties`. On the next run, it resumes from that position. The checkpoint is cleared when a run completes all rows. Use **Reset batch checkpoint** to force a fresh start.

### Settings Override

The script loads hardcoded defaults from the CONFIG object, then overrides them with any matching values from the System_Settings sheet. This means the script works even if the System_Settings sheet is missing (it uses defaults).

---

## File List

| File | Purpose |
|---|---|
| Fortify_Health_System_v31.xlsx | The complete spreadsheet with all sheets, sample data, and formatting. Upload to Google Sheets. |
| Code.gs.js | The Apps Script code. Paste into Extensions > Apps Script. |
| README.md | This guide. |
