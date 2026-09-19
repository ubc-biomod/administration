# UBC BioMod email sender

This tool sends **one personalized email per person**

You do **not** need to edit any Python code. You will:

1. Install Python once
2. Connect Gmail once
3. Put your recipient list (CSV) and email wording (Word file) in this folder
4. Double-click **Send Emails**

Personal Gmail accounts can send about **500 emails per day** per account. The tool stops at that limit.

---

## What you need before you start

- A Windows computer (Mac works too; see the last section)
- Access to either BioMod Gmail account
- About 10 minutes for the one-time Gmail connection
- Your recipient list and the email you want to send

---



## 1. Install Python (one time)

1. Open [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Download Python and run the installer
3. On the first installer screen, tick **Add python.exe to PATH**
4. Click **Install Now**
5. When it finishes, close the installer

To check it worked: press the Windows key, type `cmd`, open **Command Prompt**, and type:

```text
python --version
```

You should see a version number, such as `Python 3.12.4`. If you see “not recognized”, install Python again and make sure **Add python.exe to PATH** is ticked.

---



## 2. Install the extra pieces this tool needs (one time)

1. Open the `email-automator` folder (the folder that contains this README)
2. Click the address bar at the top of File Explorer, type `cmd`, and press Enter
  A black window should open already inside this folder
3. Copy and paste this line, then press Enter:

```text
python -m pip install -r requirements.txt
```

Wait until it finishes. You only need to do this once (or again if a teammate tells you the tool was updated).

---



## 3. Connect Gmail (one time per account)

Gmail will not let a random program send mail until you create a small “app password file” in Google Cloud. This is tedious once; after that, logging in is a normal Google popup.

Use a browser while signed into the Gmail account you will send from (`ubcbiomod@gmail.com` or `wetlab.ubcbiomod@gmail.com`).

### 3a. Create a Google Cloud project

1. Open [https://console.cloud.google.com/](https://console.cloud.google.com/)
2. If Google asks you to accept terms, accept them
3. At the top of the page, open the project picker and choose **New Project**
4. Name it something like `biomod-email-sender` and click **Create**
5. Make sure that new project is selected at the top of the page



### 3b. Turn on the Gmail API

1. Open [https://console.cloud.google.com/apis/library/gmail.googleapis.com](https://console.cloud.google.com/apis/library/gmail.googleapis.com)
2. Click **Enable**



### 3c. Fill in the consent screen and add test users

1. Open [https://console.cloud.google.com/auth/overview](https://console.cloud.google.com/auth/overview) (or search for **Google Auth Platform** / **OAuth consent screen**)
2. If asked for User Type, choose **External**, then **Create**
3. App name: `UBC BioMod Email Sender`
4. User support email: pick the BioMod Gmail address
5. Developer contact email: the same address is fine
6. Save and continue through the screens until you can finish
7. Find **Test users** / **Audience** and add:
  - `ubcbiomod@gmail.com`
  -  any other email you wish to send automated emails from
  Only accounts listed here can log in while the app is in testing mode. That is what we want.



### 3d. Create a Desktop login and download `credentials.json`

1. Open [https://console.cloud.google.com/auth/clients](https://console.cloud.google.com/auth/clients) (or **Clients** / **Credentials**)
2. Click **Create client** / **Create credentials** → **OAuth client ID**
3. Application type: **Desktop app**
4. Name: `BioMod email sender desktop`
5. Create it, then **Download** the JSON file
6. Rename the downloaded file to exactly `credentials.json`
7. Move `credentials.json` into this `email-automator` folder (next to `email.py`)

Do **not** email `credentials.json` to a public list or commit it to GitHub. Treat it like a password.

---



## 4. Prepare your recipient list (CSV)

Your list must be a **CSV** file (Excel: **File → Save As → CSV UTF-8**).

Rules:

- The **first row** is column names
- There must be a column named `email`
- Any other columns can be used as fill-in fields in the email

Example:

```text
email,first_name,team_name
jane@example.com,Jane,UBC iGEM
alex@example.com,Alex,Wetlab
```

There is a starter file named `recipients.example.csv`. You can:

1. Copy it
2. Rename the copy to `recipients.csv`
3. Open it in Excel and replace the sample people with your real list

Save `recipients.csv` in this same folder.

---



## 5. Prepare your Word template

1. Open Microsoft Word
2. Write the email **body** (not the subject — the subject goes in `config.yaml`)
3. Wherever something should change per person, type a placeholder that matches a CSV column, with double curly braces:

```text
Hi {{first_name}},

Thanks for being part of {{team_name}}. We hope to see you soon.

UBC BioMod
```

1. Save the file as **Word Document (*.docx)** named `template.docx` in this folder
  (not `.doc`, not PDF, not Google Docs unless you download it as `.docx`)

Placeholder names must match your CSV headers. Capitalization and spaces vs underscores are OK (`{{first_name}}` matches a column named `First Name`).

To create starter files automatically, open Command Prompt in this folder and run:

```text
python email.py --create-examples
```

---



## 6. Fill in `config.yaml` (no code)

Open `config.yaml`. Change these lines as needed:


| Setting          | What it means                                                                                                                                                      |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `sender_account` | `ubcbiomod@gmail.com` or `wetlab.ubcbiomod@gmail.com`                                                                                                              |
| `csv_file`       | Usually `recipients.csv`                                                                                                                                           |
| `template_file`  | Usually `template.docx`                                                                                                                                            |
| `subject`        | Email subject. You can use placeholders, e.g. `Hello {{first_name}}`                                                                                               |
| `campaign`       | A name for this blast. Keep it the same if you re-run after a crash (so people are not emailed twice). **Change it** when you send a new email to the same people. |
| `preview_count`  | How many sample emails to show before sending                                                                                                                      |
| `delay_seconds`  | Pause between sends. Leave at `1.5` unless a teammate says otherwise                                                                                               |
| `daily_limit`    | Leave at `500`                                                                                                                                                     |
| `dry_run`        | `true` = preview only, send nothing. `false` = actually send after you confirm                                                                                     |


Save the file.

---



## 7. Send the emails

**Windows:** double-click `Send Emails.bat`

The first time, a browser window should open:

1. Choose the same Gmail address as `sender_account`
2. If Google says the app is not verified, click **Continue** / **Advanced** → **Go to UBC BioMod Email Sender (unsafe)** — this is expected for a private club tool in testing mode
3. Allow sending email

Then the program will:

1. Check your CSV and Word file
2. Show 1–2 preview emails
3. Ask: `You are about to send N emails from … Continue? [y/n]`
4. Type `y` and press Enter to send, or `n` to cancel

A pause between emails is normal. Leave the window open until it says **Finished**.

### Preview without sending

In `config.yaml` set `dry_run: true`, or in Command Prompt run:

```text
python email.py --dry-run
```

---



## 8. After a send: how to read the log

Two files may appear in this folder:


| File           | Who it is for                                                                                        |
| -------------- | ---------------------------------------------------------------------------------------------------- |
| `sent_log.csv` | You. Open in Excel. Each row is one person: `sent`, `skipped`, or `failed`, plus a short explanation |
| `debug.log`    | A teammate who can read technical details. You do not need to open this unless something goes wrong  |


If the program stops halfway, **run it again with the same** `campaign` **name**. Anyone already marked `sent` in `sent_log.csv` is skipped, so they should not get a duplicate.

To email the same people with a **new** message later, change `campaign` in `config.yaml` (for example `sponsors-sept-2026`).

---



## Common problems (plain English)

**“Can't find the CSV file…”**  
The name in `config.yaml` does not match the file on disk, or the file is in a different folder. Put the CSV in this folder and check the spelling, including `.csv`.

**“Missing 'email' column…”**  
The first row of the spreadsheet must include `email`. If you exported from Excel, make sure you did not leave a blank extra header.

**“Template uses {{something}} but CSV has no matching column”**  
Add that column to the CSV, or change the placeholder in Word so it matches a real column name.

**“Row 14 skipped: 'email' field is blank”**  
That person has no address. Fill it in or ignore the skip.

**“Could not log in…” / browser never opens**  
`credentials.json` is missing or the Google Cloud steps were skipped. Repeat section 3. If an old login is stuck, delete the file that starts with `token_` in this folder and try again.

**“You logged in as X, but config.yaml says to send from Y”**  
You picked the wrong Google account in the browser. Delete the `token_…json` file and run again, choosing the address in `config.yaml`.

**“Daily send limit (500) reached”**  
Wait until tomorrow, or switch `sender_account` to the other BioMod Gmail (you will need to log in to that account once too).

**Google says the app is blocked / user is not a test user**  
Add that Gmail address as a test user on the Google Cloud consent screen (section 3c).

**Python is not recognized when I double-click the sender**  
Python is not installed, or it was installed without **Add to PATH**. Repeat section 1, then open a new Command Prompt.

---



## Mac notes

1. Install Python 3 from [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. In Terminal, `cd` into this folder and run `python3 -m pip install -r requirements.txt`
3. First time only: `chmod +x "Send Emails.command"`
4. Double-click `Send Emails.command`, or run `python3 email.py`

---



## What this folder should contain when you are ready

- `email.py` — the program (do not edit)
- `config.yaml` — your settings (do edit)
- `credentials.json` — downloaded from Google (never share publicly)
- `recipients.csv` — your people
- `template.docx` — your wording
- `Send Emails.bat` — double-click this on Windows

