#!/usr/bin/env python3
"""Send personalized emails from BioMod Gmail accounts using a CSV and a Word template."""

from __future__ import annotations

import argparse
import base64
import csv
import logging
import mimetypes
import re
import sys
import time
import traceback
from datetime import datetime, timezone
from html import escape as html_escape
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote

# This file is named email.py, which would otherwise hide Python's built-in email package.
def _import_stdlib_email_mime():
    script_dir = str(Path(__file__).resolve().parent)
    saved_path = sys.path[:]
    sys.path = [entry for entry in sys.path if entry not in ("", script_dir)]
    try:
        from email.mime.multipart import MIMEMultipart
        from email.mime.image import MIMEImage
        from email.mime.text import MIMEText

        return MIMEMultipart, MIMEImage, MIMEText
    finally:
        sys.path = saved_path


MIMEMultipart, MIMEImage, MIMEText = _import_stdlib_email_mime()

SCRIPT_DIR = Path(__file__).resolve().parent
ALLOWED_SENDERS = (
    "ubcbiomod@gmail.com",
    "wetlab.ubcbiomod@gmail.com",
)
GMAIL_SCOPES = (
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/userinfo.email",
    "openid",
)
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")
HTML_IMAGE_RE = re.compile(
    r'(<img\b[^>]*?\bsrc\s*=\s*)(["\'])([^"\']+)(\2)', re.IGNORECASE
)
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
SENT_LOG_FIELDS = (
    "timestamp",
    "campaign",
    "sender",
    "row",
    "email",
    "status",
    "message",
)


class UserFacingError(Exception):
    """An error that should be shown in plain English, without a traceback."""


def configure_stdio() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def setup_debug_log(path: Path) -> None:
    logging.basicConfig(
        filename=path,
        filemode="a",
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )


def say(message: str) -> None:
    print(message)
    logging.info(message)


def warn(message: str) -> None:
    print(message)
    logging.warning(message)


def fail(message: str) -> None:
    logging.error(message)
    raise UserFacingError(message)


def token_path_for(account: str) -> Path:
    safe = account.strip().lower().replace("@", "_at_").replace(".", "_")
    return SCRIPT_DIR / f"token_{safe}.json"


def normalize_header(value: str) -> str:
    cleaned = value.strip().lower().replace("_", " ")
    return re.sub(r"\s+", " ", cleaned)


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send personalized emails from a CSV and a Word (.docx) template."
    )
    parser.add_argument(
        "--config",
        default=str(SCRIPT_DIR / "config.yaml"),
        help="Path to config.yaml",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview emails without sending",
    )
    parser.add_argument(
        "--create-examples",
        action="store_true",
        help="Create a sample recipients.csv and template.docx in this folder",
    )
    return parser.parse_args()


def load_config(path: Path) -> dict:
    try:
        import yaml
    except ImportError:
        fail(
            "The PyYAML package is missing. Open a terminal in this folder and run:\n"
            "  python -m pip install -r requirements.txt"
        )

    if not path.exists():
        fail(
            f"Can't find the settings file at {path}. "
            "Keep config.yaml in the same folder as email.py, or pass --config with the correct path."
        )

    try:
        with path.open(encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
    except Exception as exc:
        logging.exception("Failed to parse config")
        fail(
            f"Could not read {path.name}. Make sure it is valid YAML "
            f"(check quotes and indentation). Technical detail saved in debug.log ({exc})."
        )

    if not isinstance(data, dict):
        fail(f"{path.name} should contain settings like sender_account and csv_file, not a list.")
    return data


def required_setting(config: dict, key: str) -> str:
    value = config.get(key)
    if value is None or str(value).strip() == "":
        fail(
            f"config.yaml is missing '{key}'. Open config.yaml in Notepad and fill it in. "
            "See README.md for what each setting means."
        )
    return str(value).strip()


def resolve_path(raw: str) -> Path:
    path = Path(raw)
    if not path.is_absolute():
        path = SCRIPT_DIR / path
    return path.resolve()


def create_example_files() -> None:
    csv_path = SCRIPT_DIR / "recipients.csv"
    template_path = SCRIPT_DIR / "template.docx"
    example_csv = SCRIPT_DIR / "recipients.example.csv"

    if not csv_path.exists():
        if example_csv.exists():
            csv_path.write_text(example_csv.read_text(encoding="utf-8"), encoding="utf-8")
        else:
            csv_path.write_text(
                "email,first_name,team_name\n"
                "jane@example.com,Jane,UBC iGEM\n"
                "alex@example.com,Alex,Wetlab\n",
                encoding="utf-8",
            )
        say(f"Created sample recipient list: {csv_path.name}")
    else:
        say(f"{csv_path.name} already exists — left it unchanged.")

    if not template_path.exists():
        try:
            from docx import Document
        except ImportError:
            fail(
                "python-docx is missing, so a sample Word file could not be created. "
                "Run: python -m pip install -r requirements.txt"
            )
        document = Document()
        document.add_paragraph("Hi {{first_name}},")
        document.add_paragraph("")
        document.add_paragraph(
            "This is a sample email for {{team_name}}. Replace this text with your real message. "
            "Names inside double curly braces are filled in from your spreadsheet."
        )
        document.add_paragraph("")
        document.add_paragraph("Thanks,")
        document.add_paragraph("UBC BioMod")
        document.save(template_path)
        say(f"Created sample Word template: {template_path.name}")
    else:
        say(f"{template_path.name} already exists — left it unchanged.")

    say("You can now edit those two files, then run the sender as usual.")


class PlainTextHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"br", "div", "li", "p"}:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"div", "li", "p"}:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def text(self) -> str:
        lines = (line.strip() for line in "".join(self.parts).splitlines())
        return "\n".join(line for line in lines if line).strip()


def extract_html_file_text(template_path: Path) -> str:
    try:
        html = read_html_template(template_path)
        parser = PlainTextHTMLParser()
        parser.feed(html)
        text = parser.text()
    except Exception as exc:
        logging.exception("Failed to read HTML template")
        fail(
            f"Could not read {template_path.name} as HTML. Save it as UTF-8 and try again. "
            f"Technical detail saved in debug.log ({exc})."
        )
    if not text:
        fail(f"{template_path.name} appears to be empty. Add an HTML email body and try again.")
    return text


def read_html_template(template_path: Path) -> str:
    raw = template_path.read_bytes()
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1252")


def read_recipients(csv_path: Path) -> tuple[list[str], list[dict[str, str]]]:
    if not csv_path.exists():
        fail(
            f"Can't find the CSV file at {csv_path}. "
            "Check that the file exists and that csv_file in config.yaml matches its name."
        )
    if csv_path.suffix.lower() != ".csv":
        fail(
            f"The recipient list must be a .csv file. You pointed csv_file at {csv_path.name}."
        )

    try:
        with csv_path.open(newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            if not reader.fieldnames:
                fail(
                    f"{csv_path.name} has no header row. The first row should be column names "
                    "such as email,first_name,team_name."
                )
            original_headers = [h for h in reader.fieldnames if h is not None]
            normalized_headers = [normalize_header(h) for h in original_headers]
            if len(set(normalized_headers)) != len(normalized_headers):
                fail(
                    f"{csv_path.name} has duplicate column names (ignoring capitalization). "
                    f"Found columns: {', '.join(original_headers)}."
                )
            header_map = {
                normalize_header(original): original for original in original_headers
            }
            if "email" not in header_map:
                fail(
                    f"Missing 'email' column in CSV — found columns: {', '.join(original_headers)}. "
                    "Please add an 'email' column."
                )

            rows: list[dict[str, str]] = []
            for index, raw_row in enumerate(reader, start=2):
                row = {
                    normalize_header(key): (value or "").strip()
                    for key, value in raw_row.items()
                    if key is not None
                }
                row["_row_number"] = str(index)
                rows.append(row)
    except UserFacingError:
        raise
    except Exception as exc:
        logging.exception("Failed to read CSV")
        fail(
            f"Could not read {csv_path.name}. Make sure it is a CSV saved from Excel "
            f"(File → Save As → CSV UTF-8). Technical detail saved in debug.log ({exc})."
        )

    if not rows:
        fail(f"{csv_path.name} has a header row but no recipients. Add at least one person.")
    return original_headers, rows


def extract_template_text(template_path: Path) -> str:
    if not template_path.exists():
        fail(
            f"Can't find the Word template at {template_path}. "
            "Check that the file exists and that template_file in config.yaml matches its name."
        )
    if template_path.suffix.lower() in {".htm", ".html"}:
        return extract_html_file_text(template_path)
    if template_path.suffix.lower() != ".docx":
        fail(
            f"The email template must be a Word .docx or HTML .htm/.html file. "
            f"You pointed template_file at {template_path.name}."
        )

    try:
        from docx import Document
    except ImportError:
        fail(
            "The python-docx package is missing. Open a terminal in this folder and run:\n"
            "  python -m pip install -r requirements.txt"
        )

    try:
        document = Document(template_path)
        paragraphs = [paragraph.text for paragraph in document.paragraphs]
        for table in document.tables:
            for table_row in table.rows:
                for cell in table_row.cells:
                    paragraphs.extend(cell.paragraphs and [p.text for p in cell.paragraphs] or [])
    except Exception as exc:
        logging.exception("Failed to read Word template")
        fail(
            f"Could not open {template_path.name}. Save it again from Microsoft Word as "
            f"a .docx file. Technical detail saved in debug.log ({exc})."
        )

    text = "\n".join(paragraphs).strip()
    if not text:
        fail(
            f"{template_path.name} appears to be empty. Type your email in Word, using "
            "placeholders like {{first_name}}, then save and try again."
        )
    return text


def extract_template_html(template_path: Path) -> str:
    if template_path.suffix.lower() in {".htm", ".html"}:
        try:
            html = read_html_template(template_path).strip()
        except Exception as exc:
            logging.exception("Failed to read HTML template")
            fail(
                f"Could not read {template_path.name} as HTML. Save it as UTF-8 and try again. "
                f"Technical detail saved in debug.log ({exc})."
            )
        if not html:
            fail(f"{template_path.name} appears to be empty. Add an HTML email body and try again.")
        return html

    try:
        import mammoth
    except ImportError:
        fail(
            "The mammoth package is missing. Open a terminal in this folder and run:\n"
            "  python -m pip install -r requirements.txt"
        )

    try:
        result = mammoth.convert_to_html(str(template_path))
        for message in result.messages:
            logging.warning("DOCX conversion: %s", message)
        html = result.value.strip()
    except Exception as exc:
        logging.exception("Failed to convert Word template to HTML")
        fail(
            f"Could not preserve the formatting in {template_path.name}. Technical detail saved in "
            f"debug.log ({exc})."
        )

    if not html:
        fail(f"{template_path.name} did not contain any content that could be formatted for email.")
    # Email clients handle inline paragraph margins more consistently than DOCX spacing rules.
    return re.sub(r"<p(\s[^>]*)?>", r'<p\1 style="margin: 0 0 10pt;">', html)


def embed_local_images(html: str, template_dir: Path) -> tuple[str, list[tuple[str, bytes, str, str]]]:
    images: list[tuple[str, bytes, str, str]] = []
    image_index = 0

    def replace(match: re.Match[str]) -> str:
        nonlocal image_index
        source = unquote(match.group(3))
        if source.startswith(("data:", "http://", "https://", "cid:")):
            return match.group(0)

        image_path = (template_dir / source).resolve()
        try:
            image_path.relative_to(template_dir.resolve())
        except ValueError:
            logging.warning("Skipping image outside template folder: %s", source)
            return match.group(0)
        if not image_path.is_file():
            logging.warning("Could not find HTML image: %s", image_path)
            return match.group(0)

        content_type, _ = mimetypes.guess_type(image_path.name)
        if not content_type or not content_type.startswith("image/"):
            logging.warning("Skipping unsupported HTML image: %s", image_path)
            return match.group(0)
        image_index += 1
        content_id = f"image-{image_index}@biomod"
        subtype = content_type.split("/", 1)[1]
        images.append((content_id, image_path.read_bytes(), subtype, image_path.name))
        return f'{match.group(1)}{match.group(2)}cid:{content_id}{match.group(4)}'

    return HTML_IMAGE_RE.sub(replace, html), images


def find_placeholders(*texts: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    seen: set[str] = set()
    for text in texts:
        for match in PLACEHOLDER_RE.finditer(text):
            original = match.group(1).strip()
            name = normalize_header(original)
            if name and name not in seen:
                seen.add(name)
                found.append((name, original))
    return found


def substitute(text: str, row: dict[str, str]) -> str:
    def replace(match: re.Match[str]) -> str:
        key = normalize_header(match.group(1))
        return row.get(key, "")

    return PLACEHOLDER_RE.sub(replace, text)


def substitute_html(text: str, row: dict[str, str]) -> str:
    def replace(match: re.Match[str]) -> str:
        key = normalize_header(match.group(1))
        return html_escape(row.get(key, ""))

    return PLACEHOLDER_RE.sub(replace, text)


def validate_placeholders(placeholders: list[tuple[str, str]], csv_headers: list[str]) -> None:
    available = {normalize_header(header) for header in csv_headers}
    missing = [original for name, original in placeholders if name not in available]
    if missing:
        pretty = ", ".join(f"{{{{{name}}}}}" for name in missing)
        fail(
            f"Template uses {pretty} but CSV has no matching column — "
            f"found columns: {', '.join(csv_headers)}. "
            "Check your CSV headers or template placeholders."
        )


def is_plausible_email(value: str) -> bool:
    return bool(EMAIL_RE.match(value))


def load_sent_log(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    try:
        with path.open(newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            return [dict(row) for row in reader]
    except Exception as exc:
        logging.exception("Failed to read sent log")
        fail(
            f"Could not read {path.name}. If this file looks damaged, rename it "
            f"(for example to sent_log.bak.csv) and run again. Technical detail saved in debug.log ({exc})."
        )
    return []


def append_sent_log(path: Path, record: dict[str, str]) -> None:
    new_file = not path.exists()
    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=SENT_LOG_FIELDS)
        if new_file:
            writer.writeheader()
        writer.writerow({field: record.get(field, "") for field in SENT_LOG_FIELDS})


def already_sent_emails(log_rows: list[dict[str, str]], campaign: str, sender: str) -> set[str]:
    sent: set[str] = set()
    for row in log_rows:
        if (
            row.get("campaign", "").strip() == campaign
            and row.get("sender", "").strip().lower() == sender
            and row.get("status", "").strip().lower() == "sent"
        ):
            sent.add(row.get("email", "").strip().lower())
    return sent


def sent_today_count(log_rows: list[dict[str, str]], sender: str, today: datetime) -> int:
    count = 0
    today_str = today.date().isoformat()
    for row in log_rows:
        if row.get("sender", "").strip().lower() != sender:
            continue
        if row.get("status", "").strip().lower() != "sent":
            continue
        timestamp = row.get("timestamp", "")
        if timestamp.startswith(today_str):
            count += 1
    return count


def preview_email(index: int, recipient: dict[str, str], subject: str, body: str) -> None:
    say("")
    say(f"----- Preview {index} -----")
    say(f"To: {recipient['email']}")
    say(f"Subject: {subject}")
    say("")
    say(body)
    say("----- End preview -----")


def confirm_send(count: int, sender: str) -> bool:
    say("")
    prompt = f"You are about to send {count} emails from {sender}. Continue? [y/n] "
    try:
        answer = input(prompt).strip().lower()
    except EOFError:
        fail("No confirmation was provided, so nothing was sent.")
    return answer in {"y", "yes"}


def import_google() -> tuple:
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
        from googleapiclient.errors import HttpError
    except ImportError:
        fail(
            "Google Gmail libraries are missing. Open a terminal in this folder and run:\n"
            "  python -m pip install -r requirements.txt"
        )
    return Request, Credentials, InstalledAppFlow, build, HttpError


def authenticate(sender: str, credentials_path: Path, token_path: Path):
    Request, Credentials, InstalledAppFlow, build, _HttpError = import_google()

    if not credentials_path.exists():
        fail(
            f"Can't find Google login file credentials.json at {credentials_path}. "
            "A one-time Google Cloud setup is required. Follow the 'Connect Gmail' section in README.md, "
            "then put the downloaded credentials.json in this folder."
        )

    creds = None
    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), GMAIL_SCOPES)
        except Exception:
            logging.exception("Could not load saved token")
            warn(
                f"The saved login file {token_path.name} could not be read, so you will need to log in again."
            )
            creds = None

    try:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        if not creds or not creds.valid:
            say("A browser window will open so you can log in to Gmail. Use the account listed in config.yaml.")
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    except UserFacingError:
        raise
    except Exception:
        logging.exception("OAuth failed")
        fail(
            f"Could not log in to {sender} — your saved login may have expired or the Google setup is incomplete. "
            "Run the script again and log in when the browser opens. If a browser never appears, "
            "follow the 'Connect Gmail' steps in README.md. Technical detail is in debug.log."
        )

    try:
        oauth = build("oauth2", "v2", credentials=creds, cache_discovery=False)
        profile = oauth.userinfo().get().execute()
        logged_in = (profile.get("email") or "").strip().lower()
    except Exception:
        logging.exception("Could not read logged-in email")
        fail(
            f"Logged in, but could not confirm which Gmail account was used. "
            f"Delete {token_path.name} and run again, then choose {sender} in the browser."
        )

    if logged_in != sender:
        fail(
            f"You logged in as {logged_in}, but config.yaml says to send from {sender}. "
            f"Delete {token_path.name} in this folder, then run again and choose {sender} in the browser."
        )

    gmail = build("gmail", "v1", credentials=creds, cache_discovery=False)
    say(f"Logged in as {sender}.")
    return gmail


def build_message(
    sender: str,
    to_address: str,
    subject: str,
    body: str,
    html_body: str | None = None,
    inline_images: list[tuple[str, bytes, str, str]] | None = None,
) -> dict:
    message = MIMEMultipart("mixed")
    alternative = MIMEMultipart("alternative")
    message.attach(alternative)
    message["To"] = to_address
    message["From"] = sender
    message["Subject"] = subject
    alternative.attach(MIMEText(body, "plain", "utf-8"))
    if html_body is None:
        html_body = "".join(
            f"<p>{html_escape(line) if line else '&nbsp;'}</p>" for line in body.split("\n")
        )
    if inline_images:
        related = MIMEMultipart("related")
        related.attach(MIMEText(html_body, "html", "utf-8"))
        for content_id, image_data, subtype, filename in inline_images:
            image = MIMEImage(image_data, _subtype=subtype)
            image.add_header("Content-ID", f"<{content_id}>")
            image.add_header("Content-Disposition", "inline", filename=filename)
            related.attach(image)
        alternative.attach(related)
    else:
        alternative.attach(MIMEText(html_body, "html", "utf-8"))
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")
    return {"raw": raw}


def describe_send_error(exc: Exception, HttpError) -> tuple[str, bool, bool]:
    """Return (human message, daily quota hit, short-window rate limit)."""
    if HttpError is not None and isinstance(exc, HttpError):
        status = getattr(exc.resp, "status", None)
        content = ""
        try:
            content = exc.content.decode("utf-8", errors="replace") if exc.content else ""
        except Exception:
            content = str(exc)
        lowered = content.lower()
        daily_markers = ("dailylimitexceeded", "daily limit", "quotaexceeded")
        rate_markers = ("user_rate_limit_exceeded", "ratelimitexceeded", "rate limit")
        if status in (429, 403) and any(marker in lowered for marker in daily_markers):
            return "Gmail daily send limit reached", True, False
        if status in (429, 403) and any(marker in lowered for marker in rate_markers):
            return "Gmail asked us to slow down (rate limit)", False, True
        if status == 400:
            return "Gmail rejected the address or message as invalid", False, False
        return f"Gmail returned an error (code {status})", False, False
    return "network or unexpected Gmail error", False, False


def send_one(
    gmail,
    sender: str,
    to_address: str,
    subject: str,
    body: str,
    html_body: str,
    inline_images: list[tuple[str, bytes, str, str]],
) -> None:
    gmail.users().messages().send(
        userId="me",
        body=build_message(sender, to_address, subject, body, html_body, inline_images),
    ).execute()


def prepare_queue(
    rows: list[dict[str, str]],
    already_sent: set[str],
    sent_log_path: Path,
    campaign: str,
    sender: str,
) -> list[dict[str, str]]:
    queue: list[dict[str, str]] = []
    for row in rows:
        row_number = row["_row_number"]
        address = row.get("email", "").strip()
        if not address:
            message = f"Row {row_number} skipped: 'email' field is blank."
            warn(message)
            append_sent_log(
                sent_log_path,
                {
                    "timestamp": utc_now().isoformat(timespec="seconds"),
                    "campaign": campaign,
                    "sender": sender,
                    "row": row_number,
                    "email": "",
                    "status": "skipped",
                    "message": message,
                },
            )
            continue
        if not is_plausible_email(address):
            message = (
                f"Row {row_number} skipped: '{address}' does not look like a valid email address."
            )
            warn(message)
            append_sent_log(
                sent_log_path,
                {
                    "timestamp": utc_now().isoformat(timespec="seconds"),
                    "campaign": campaign,
                    "sender": sender,
                    "row": row_number,
                    "email": address,
                    "status": "skipped",
                    "message": message,
                },
            )
            continue
        if address.lower() in already_sent:
            message = (
                f"Row {row_number} skipped: {address} already has a successful send "
                f"logged for campaign '{campaign}'."
            )
            say(message)
            continue
        queue.append(row)
    return queue


def run(args: argparse.Namespace) -> int:
    debug_path = SCRIPT_DIR / "debug.log"
    setup_debug_log(debug_path)
    logging.info("Starting email sender")

    if args.create_examples:
        create_example_files()
        return 0

    config = load_config(Path(args.config))
    sender = required_setting(config, "sender_account").lower()
    # if sender not in ALLOWED_SENDERS:
    #     fail(
    #         f"sender_account in config.yaml is '{sender}', which is not allowed. "
    #         f"Use one of: {', '.join(ALLOWED_SENDERS)}."
    #     )

    csv_path = resolve_path(required_setting(config, "csv_file"))
    template_path = resolve_path(required_setting(config, "template_file"))
    subject_template = required_setting(config, "subject")
    campaign = str(config.get("campaign") or "default").strip() or "default"
    preview_count = int(config.get("preview_count") or 2)
    delay_seconds = float(config.get("delay_seconds") or 1.5)
    daily_limit = int(config.get("daily_limit") or 500)
    dry_run = bool(args.dry_run or config.get("dry_run"))
    credentials_path = SCRIPT_DIR / "credentials.json"
    sent_log_path = SCRIPT_DIR / "sent_log.csv"

    csv_headers, rows = read_recipients(csv_path)
    template_text = extract_template_text(template_path)
    template_html = extract_template_html(template_path)
    template_html, inline_images = embed_local_images(template_html, template_path.parent)
    placeholders = find_placeholders(template_text, subject_template)
    validate_placeholders(placeholders, csv_headers)

    log_rows = load_sent_log(sent_log_path)
    already_sent = already_sent_emails(log_rows, campaign, sender)
    used_today = sent_today_count(log_rows, sender, utc_now())
    remaining_today = max(0, daily_limit - used_today)

    queue = prepare_queue(rows, already_sent, sent_log_path, campaign, sender)
    if remaining_today <= 0:
        fail(
            f"Daily send limit ({daily_limit}) reached for {sender}. "
            "Remaining recipients will need to be sent tomorrow or from the other account."
        )
    if len(queue) > remaining_today:
        warn(
            f"{sender} can send {remaining_today} more email(s) today (limit {daily_limit}). "
            f"Only the first {remaining_today} of {len(queue)} remaining recipients will be sent."
        )
        queue = queue[:remaining_today]

    if not queue:
        say(
            "There is nobody left to email. Either the CSV is empty after skipping invalid rows, "
            "or everyone was already marked as sent in sent_log.csv for this campaign."
        )
        return 0

    say(f"Ready to send from: {sender}")
    say(f"Campaign: {campaign}")
    say(f"Recipients in this run: {len(queue)}")
    say(f"Already sent today from this account (from the log): {used_today} of {daily_limit}")

    samples = queue[: max(1, preview_count)]
    for index, recipient in enumerate(samples, start=1):
        preview_email(
            index,
            recipient,
            substitute(subject_template, recipient),
            substitute(template_text, recipient),
        )

    if dry_run:
        say("")
        say("Dry run only — no emails were sent. Set dry_run: false in config.yaml when you are ready.")
        return 0

    if not confirm_send(len(queue), sender):
        say("Cancelled. No emails were sent.")
        return 0

    gmail = authenticate(sender, credentials_path, token_path_for(sender))
    _, _, _, _, HttpError = import_google()

    sent_ok = 0
    failed = 0
    for index, recipient in enumerate(queue, start=1):
        address = recipient["email"]
        row_number = recipient["_row_number"]
        subject = substitute(subject_template, recipient)
        body = substitute(template_text, recipient)
        html_body = substitute_html(template_html, recipient)
        try:
            send_one(gmail, sender, address, subject, body, html_body, inline_images)
            message = f"Sent to {address} (row {row_number})."
            say(f"[{index}/{len(queue)}] {message}")
            append_sent_log(
                sent_log_path,
                {
                    "timestamp": utc_now().isoformat(timespec="seconds"),
                    "campaign": campaign,
                    "sender": sender,
                    "row": row_number,
                    "email": address,
                    "status": "sent",
                    "message": message,
                },
            )
            sent_ok += 1
        except Exception as exc:
            logging.exception("Send failed for %s row %s", address, row_number)
            reason, quota_hit, rate_limited = describe_send_error(exc, HttpError)
            if rate_limited:
                warn(
                    f"Failed to send to {address} (row {row_number}): {reason}. "
                    "Waiting 30 seconds, then continuing with the rest of the list."
                )
                append_sent_log(
                    sent_log_path,
                    {
                        "timestamp": utc_now().isoformat(timespec="seconds"),
                        "campaign": campaign,
                        "sender": sender,
                        "row": row_number,
                        "email": address,
                        "status": "failed",
                        "message": reason,
                    },
                )
                failed += 1
                time.sleep(30)
                continue
            if quota_hit:
                message = (
                    f"Daily send limit ({daily_limit}) reached for {sender}. "
                    "Remaining recipients will need to be sent tomorrow or from the other account."
                )
                warn(
                    f"Failed to send to {address} (row {row_number}): {reason}. {message}"
                )
                append_sent_log(
                    sent_log_path,
                    {
                        "timestamp": utc_now().isoformat(timespec="seconds"),
                        "campaign": campaign,
                        "sender": sender,
                        "row": row_number,
                        "email": address,
                        "status": "failed",
                        "message": reason,
                    },
                )
                failed += 1
                say(message)
                break
            message = f"Failed to send to {address} (row {row_number}): {reason}."
            warn(message)
            append_sent_log(
                sent_log_path,
                {
                    "timestamp": utc_now().isoformat(timespec="seconds"),
                    "campaign": campaign,
                    "sender": sender,
                    "row": row_number,
                    "email": address,
                    "status": "failed",
                    "message": reason,
                },
            )
            failed += 1
        if index < len(queue):
            time.sleep(max(0.0, delay_seconds))

    say("")
    say(f"Finished. Sent: {sent_ok}. Failed or skipped during send: {failed}.")
    say(f"A full record is in {sent_log_path.name}. Technical details (if any) are in {debug_path.name}.")
    return 0 if failed == 0 else 1


def main() -> int:
    configure_stdio()
    args = parse_args()
    try:
        return run(args)
    except UserFacingError as exc:
        print()
        print(str(exc))
        print()
        print("Nothing else was changed. If you need help, share debug.log and sent_log.csv with a teammate.")
        return 1
    except KeyboardInterrupt:
        print()
        print("Stopped. Emails already sent stay sent; check sent_log.csv before running again.")
        return 1
    except Exception:
        logging.exception("Unhandled error")
        print()
        print(
            "Something unexpected went wrong. The technical details were saved in debug.log "
            "in this folder — you can send that file to a teammate. "
            f"({traceback.format_exc().splitlines()[-1]})"
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
