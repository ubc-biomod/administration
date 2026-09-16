# Engineering Scope Ticket: Automated Individual Email Sending via Gmail API

## Summary
Build a Python-based tool that lets `ubcbiomod@gmail.com` and/or `wetlab.ubcbiomod@gmail.com` (standard, unpaid personal Gmail accounts) send individually personalized emails via the Gmail API, using a CSV recipient list and a **Word (.docx)** template — up to 500 emails/day per account. The tool must be runnable by a non-coder team member with minimal setup, and any failures must be clearly explained, not just logged as raw errors.

## Background / Context
Team members sending these emails are not necessarily developers. The script needs to be triggerable with a simple action, pull recipient data from a CSV, pull email content from a .docx template, and just work — with clear, human-readable instructions/errors if something's misconfigured or a send fails.

## Requirements
- Support sending from either `ubcbiomod@gmail.com` or `wetlab.ubcbiomod@gmail.com` (configurable via a simple config file — not a code edit).
- Each email sent individually, personalized per recipient (name + any other merge fields).
- Recipient list: CSV file with columns such as name, email, and any merge fields.
- Email template: a **Word (.docx) file only**, with placeholders (e.g., `{{first_name}}`) replaced per recipient using `python-docx`.
- Up to 500 emails/day per account.
- Built in Python using `google-api-python-client` with Gmail API (`gmail.send` scope).
- **Non-coder usability is a hard requirement.**
- **All failures must be descriptive** — every error (setup-time or send-time) must clearly state what went wrong, which record/recipient it affected (if applicable), and ideally how to fix it. No raw stack traces or generic exceptions surfaced to the user.

## Technical Notes

**Auth**
- Standard OAuth 2.0 consent flow per personal Gmail account (no Workspace domain-wide delegation available).
- One-time setup per account: run script, log in via browser popup, token saved locally for reuse.
- Refresh tokens stored locally so the non-coder doesn't need to re-authenticate every run.

**Template handling (.docx only)**
- Use `python-docx` to open the .docx file and extract paragraph text (and basic formatting — bold/italic — if feasible for v1).
- Placeholder syntax (e.g., `{{first_name}}`) scanned and replaced per-recipient using CSV row data.
- Validate the template contains only placeholders that exist as CSV columns; descriptive error if not (e.g., "Template uses {{event_date}} but CSV has no 'event_date' column — check your CSV headers or template placeholders.").
- Decide v1 scope: plain-text-equivalent output vs. preserving .docx formatting into HTML email body. Recommend plain text / minimal formatting for v1.

**CSV handling**
- Use Python's built-in `csv` or `pandas` to read recipient list.
- Validate required columns exist before running; descriptive error if not (e.g., "Missing 'email' column in CSV — found columns: name, note. Please add an 'email' column.").
- Validate each row has a non-empty, plausibly formatted email address before attempting to send; skip and log descriptively if not (e.g., "Row 14 skipped: 'email' field is blank.").

**Non-coder usability**
- No code editing required to run day-to-day. Approach:
  - A config file (e.g., `config.yaml`) or interactive CLI prompts for: which account to send from, path to CSV, path to .docx template.
  - Consider a `.bat`/`.command` double-click launcher.
- Dry run / preview mode: show 1-2 generated emails before sending the full batch.
- Confirmation prompt before sending ("You are about to send 342 emails from ubcbiomod@gmail.com. Continue? [y/n]").

**Descriptive failure handling (expanded)**
- Every failure category gets its own clear, plain-English message template, e.g.:
  - **Auth failure:** "Could not log in to ubcbiomod@gmail.com — your saved login has expired. Run the script again and log in when the browser opens."
  - **Missing file:** "Can't find the CSV file at [path]. Check the file exists and the path is correct."
  - **Bad CSV column:** "Missing 'email' column in CSV — found columns: [x, y, z]."
  - **Template/CSV mismatch:** "Template placeholder {{x}} has no matching CSV column."
  - **Per-recipient send failure:** "Failed to send to jane@example.com (row 22): [reason, e.g., invalid address / quota exceeded / network error]."
  - **Quota exceeded:** "Daily send limit (500) reached for ubcbiomod@gmail.com. Remaining recipients will need to be sent tomorrow or from the other account."
- All descriptive errors/warnings written both to console/terminal (for immediate visibility) and to the log file (for later review).
- Underlying technical exception details still captured in the log file (for a developer to debug later) but not shown front-and-center to the non-coder user.

**Rate limiting & quotas**
- Gmail's own cap (~500/day for personal accounts) aligns with this requirement; no quota increase expected, but worth a quick check against current Google policy before build.
- Add delay between sends to avoid tripping short-window spam detection.

**Reliability**
- Log every send attempt (success/failure) to a local file (e.g., `sent_log.csv`) with timestamp and descriptive status.
- Skip recipients already marked as sent in the log, so re-running after an interruption doesn't duplicate-send.
- Per-recipient failures logged descriptively and skipped, not fatal to the whole run.

**Credential storage**
- OAuth client secret + per-account tokens stored locally, excluded from version control (`.gitignore`).

## Acceptance Criteria
- [ ] Non-coder can run the tool via a single command or double-click launcher, no code edits needed.
- [ ] Tool authenticates via OAuth per account, with token reused after first login.
- [ ] Tool reads recipient CSV and validates required columns, with a descriptive error if malformed.
- [ ] Tool reads a .docx template and correctly substitutes placeholders per recipient using `python-docx`.
- [ ] Descriptive error shown if a template placeholder has no matching CSV column.
- [ ] Descriptive, per-row error/skip if a CSV row has missing/invalid email data.
- [ ] Preview/dry-run of at least one generated email shown before sending.
- [ ] Confirmation step before sending the full batch.
- [ ] Sends personalized, individual emails (not CC/BCC) up to 500/day successfully.
- [ ] All failures (auth, file, template, per-recipient send) produce a clear, human-readable message — verified by intentionally triggering each failure type during testing.
- [ ] No raw Python tracebacks are shown to the end user in normal operation.
- [ ] Duplicate-send protection confirmed by re-running script after partial send.
- [ ] Written setup guide (README or one-pager) walking a non-coder through: first-time login, where to put the CSV, where to put the .docx template, how to run it, and how to read the log file.

## Open Questions
1. What merge fields beyond name are needed (e.g., team name, event date)?
2. Should this run on someone's personal laptop ad hoc, or does it need to live somewhere shared?
3. Who will be the "non-coder" running this day-to-day — worth a quick walkthrough once built?
4. Any formatting requirements in the email body (bold, links) needed for v1, or is plain text acceptable?

## Estimated Effort
TBD — likely 4-5 days: OAuth setup, `.docx` template parsing + placeholder substitution, CSV validation, non-coder-friendly CLI/config layer, descriptive error handling across all failure points, and setup guide are the main drivers.

---

Want me to draft the actual Python script now, or keep iterating on scope first?

**Title:** Enable automated, personalized sending of up to 500 emails/day from BioMod Gmail accounts (non-technical user friendly)

**Type:** Feature / Integration

**Priority:** TBD

## Summary
Build a Python-based tool that lets `ubcbiomod@gmail.com` and/or `wetlab.ubcbiomod@gmail.com` (standard, unpaid personal Gmail accounts) send individually personalized emails via the Gmail API, using a CSV recipient list and a **Word (.docx)** template — up to 500 emails/day per account. The tool must be runnable by a non-coder team member with minimal setup, and any failures must be clearly explained, not just logged as raw errors.

## Background / Context
Team members sending these emails are not necessarily developers. The script needs to be triggerable with a simple action, pull recipient data from a CSV, pull email content from a .docx template, and just work — with clear, human-readable instructions/errors if something's misconfigured or a send fails.

## Requirements
- Support sending from either `ubcbiomod@gmail.com` or `wetlab.ubcbiomod@gmail.com` (configurable via a simple config file — not a code edit).
- Each email sent individually, personalized per recipient (name + any other merge fields).
- Recipient list: CSV file with columns such as name, email, and any merge fields.
- Email template: a **Word (.docx) file only**, with placeholders (e.g., `{{first_name}}`) replaced per recipient using `python-docx`.
- Up to 500 emails/day per account.
- Built in Python using `google-api-python-client` with Gmail API (`gmail.send` scope).
- **Non-coder usability is a hard requirement.**
- **All failures must be descriptive** — every error (setup-time or send-time) must clearly state what went wrong, which record/recipient it affected (if applicable), and ideally how to fix it. No raw stack traces or generic exceptions surfaced to the user.

## Technical Notes

**Auth**
- Standard OAuth 2.0 consent flow per personal Gmail account (no Workspace domain-wide delegation available).
- One-time setup per account: run script, log in via browser popup, token saved locally for reuse.
- Refresh tokens stored locally so the non-coder doesn't need to re-authenticate every run.

**Template handling (.docx only)**
- Use `python-docx` to open the .docx file and extract paragraph text (and basic formatting — bold/italic — if feasible for v1).
- Placeholder syntax (e.g., `{{first_name}}`) scanned and replaced per-recipient using CSV row data.
- Validate the template contains only placeholders that exist as CSV columns; descriptive error if not (e.g., "Template uses {{event_date}} but CSV has no 'event_date' column — check your CSV headers or template placeholders.").
- Decide v1 scope: plain-text-equivalent output vs. preserving .docx formatting into HTML email body. Recommend plain text / minimal formatting for v1.

**CSV handling**
- Use Python's built-in `csv` or `pandas` to read recipient list.
- Validate required columns exist before running; descriptive error if not (e.g., "Missing 'email' column in CSV — found columns: name, note. Please add an 'email' column.").
- Validate each row has a non-empty, plausibly formatted email address before attempting to send; skip and log descriptively if not (e.g., "Row 14 skipped: 'email' field is blank.").

**Non-coder usability**
- No code editing required to run day-to-day. Approach:
  - A config file (e.g., `config.yaml`) or interactive CLI prompts for: which account to send from, path to CSV, path to .docx template.
  - Consider a `.bat`/`.command` double-click launcher.
- Dry run / preview mode: show 1-2 generated emails before sending the full batch.
- Confirmation prompt before sending ("You are about to send 342 emails from ubcbiomod@gmail.com. Continue? [y/n]").

**Descriptive failure handling (expanded)**
- Every failure category gets its own clear, plain-English message template, e.g.:
  - **Auth failure:** "Could not log in to ubcbiomod@gmail.com — your saved login has expired. Run the script again and log in when the browser opens."
  - **Missing file:** "Can't find the CSV file at [path]. Check the file exists and the path is correct."
  - **Bad CSV column:** "Missing 'email' column in CSV — found columns: [x, y, z]."
  - **Template/CSV mismatch:** "Template placeholder {{x}} has no matching CSV column."
  - **Per-recipient send failure:** "Failed to send to jane@example.com (row 22): [reason, e.g., invalid address / quota exceeded / network error]."
  - **Quota exceeded:** "Daily send limit (500) reached for ubcbiomod@gmail.com. Remaining recipients will need to be sent tomorrow or from the other account."
- All descriptive errors/warnings written both to console/terminal (for immediate visibility) and to the log file (for later review).
- Underlying technical exception details still captured in the log file (for a developer to debug later) but not shown front-and-center to the non-coder user.

**Rate limiting & quotas**
- Gmail's own cap (~500/day for personal accounts) aligns with this requirement; no quota increase expected, but worth a quick check against current Google policy before build.
- Add delay between sends to avoid tripping short-window spam detection.

**Reliability**
- Log every send attempt (success/failure) to a local file (e.g., `sent_log.csv`) with timestamp and descriptive status.
- Skip recipients already marked as sent in the log, so re-running after an interruption doesn't duplicate-send.
- Per-recipient failures logged descriptively and skipped, not fatal to the whole run.

**Credential storage**
- OAuth client secret + per-account tokens stored locally, excluded from version control (`.gitignore`).

## Acceptance Criteria
- [ ] Non-coder can run the tool via a single command or double-click launcher, no code edits needed.
- [ ] Tool authenticates via OAuth per account, with token reused after first login.
- [ ] Tool reads recipient CSV and validates required columns, with a descriptive error if malformed.
- [ ] Tool reads a .docx template and correctly substitutes placeholders per recipient using `python-docx`.
- [ ] Descriptive error shown if a template placeholder has no matching CSV column.
- [ ] Descriptive, per-row error/skip if a CSV row has missing/invalid email data.
- [ ] Preview/dry-run of at least one generated email shown before sending.
- [ ] Confirmation step before sending the full batch.
- [ ] Sends personalized, individual emails (not CC/BCC) up to 500/day successfully.
- [ ] All failures (auth, file, template, per-recipient send) produce a clear, human-readable message — verified by intentionally triggering each failure type during testing.
- [ ] No raw Python tracebacks are shown to the end user in normal operation.
- [ ] Duplicate-send protection confirmed by re-running script after partial send.
- [ ] Written setup guide (README or one-pager) walking a non-coder through: first-time login, where to put the CSV, where to put the .docx template, how to run it, and how to read the log file.

## Open Questions
1. What merge fields beyond name are needed (e.g., team name, event date)?
2. Should this run on someone's personal laptop ad hoc, or does it need to live somewhere shared?
3. Who will be the "non-coder" running this day-to-day — worth a quick walkthrough once built?
4. Any formatting requirements in the email body (bold, links) needed for v1, or is plain text acceptable?

## Estimated Effort
TBD — likely 4-5 days: OAuth setup, `.docx` template parsing + placeholder substitution, CSV validation, non-coder-friendly CLI/config layer, descriptive error handling across all failure points, and setup guide are the main drivers.

