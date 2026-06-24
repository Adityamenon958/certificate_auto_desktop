"""
Certificate Auto - Desktop app with UI.
Add entries (Name, Course, Month, Email), optionally via Paste rows (TSV/CSV from Sheets),
then Generate & Send certificates. Save/Load list to JSON / CSV / Excel. No Google Sheet.
"""
import base64
import csv
import io
import json
import os
import re
import shlex
import sys
import smtplib
from datetime import datetime, date, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from jinja2 import Environment, FileSystemLoader
import pdfkit
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Paths: frozen (PyInstaller) vs script
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    BASE_PATH = sys._MEIPASS
    EXE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR = BASE_PATH

_env_path = os.path.join(EXE_DIR, ".env")
load_dotenv(_env_path)

TEMPLATE_DIR = os.path.join(BASE_PATH, "templates")
STATIC_DIR = os.path.join(BASE_PATH, "static")
if not os.path.isdir(TEMPLATE_DIR):
    TEMPLATE_DIR = os.path.join(os.path.dirname(BASE_PATH), "templates")
if not os.path.isdir(STATIC_DIR):
    STATIC_DIR = os.path.join(os.path.dirname(BASE_PATH), "static")

OUTPUT_DIR = os.getenv("OUTPUT_DIR") or os.path.join(EXE_DIR, "certificates")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# SMTP (from .env)
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
TEMPLATE_NAME = os.getenv("TEMPLATE_NAME", "certificate_template.html")
TEMPLATE_NAME_2 = os.getenv("TEMPLATE_NAME_2", "certificate_template_2.html")
TEMPLATE_NAME_MARKSHEET = os.getenv("TEMPLATE_NAME_MARKSHEET", "marksheet_template.html")

MARKSHEET_SUBJECTS = (
    {"key": "theory_1", "section": "theory", "name": "OVERVIEW OF EARLY CHILDHOOD EDUCATION"},
    {"key": "theory_2", "section": "theory", "name": "SPECIAL CHILDREN"},
    {"key": "theory_3", "section": "theory", "name": "KNOWLEDGE ABOUT CHILD DEVELOPMENT"},
    {"key": "theory_4", "section": "theory", "name": "NUTRITION IN CHILD DEVELOPMENT"},
    {"key": "theory_5", "section": "theory", "name": "CURRICULUM"},
    {"key": "theory_6", "section": "theory", "name": "CLASSROOM MANAGEMENT"},
    {"key": "practical_1", "section": "practical", "name": "CREATIVE JOURNALS"},
    {"key": "practical_2", "section": "practical", "name": "LESSON PLAN / WORKSHEET TECHNIQUE"},
    {"key": "practical_3", "section": "practical", "name": "LEARNING GAMES HOLISTIC DEVELOPMENT OF PRESCHOOLER"},
    {"key": "practical_4", "section": "practical", "name": "INTERVIEW / RESUME TECHNIQUE"},
)
MARK_KEYS = tuple(s["key"] for s in MARKSHEET_SUBJECTS)
MARKSHEET_MIN_TOTAL = 750
MARKSHEET_OUT_OF = 1000

CERTIFICATE_TYPES = {
    "1": {
        "label": "Certificate of Achievement",
        "template": TEMPLATE_NAME,
        "page_width": "215mm",
        "page_height": "158mm",
        "required": ("name", "course", "month", "email"),
        "tree_columns": ("Name", "Course", "Month", "Email"),
        "entry_keys": ("name", "course", "month", "email"),
    },
    "2": {
        "label": "Certificate of Completion",
        "template": TEMPLATE_NAME_2,
        "page_width": "210mm",
        "page_height": "297mm",
        "required": ("name", "gr_no", "course", "year", "grade", "email"),
        "tree_columns": ("Name", "Gr.No.", "Course", "Year", "Grade", "Email"),
        "entry_keys": ("name", "gr_no", "course", "year", "grade", "email"),
        "program_title_variants": {
            "post_grad": {
                "label": "Post Graduation Diploma",
                "heading_line1": "Certificate of Completion in Post Graduation Diploma in",
            },
            "diploma": {
                "label": "Diploma",
                "heading_line1": "Certificate of Completion in Diploma in",
            },
        },
    },
    "3": {
        "label": "Marksheet",
        "template": TEMPLATE_NAME_MARKSHEET,
        "page_width": "810px",
        "page_height": "1440px",
        "required": ("name", "gr_no", "year", "email") + MARK_KEYS,
        "tree_columns": ("Name", "Gr.No.", "Total", "Grade", "Email"),
        "entry_keys": ("name", "gr_no", "year", "email") + MARK_KEYS,
        "subtitle": "DIPLOMA IN EARLY CHILDHOOD CARE AND EDUCATION",
    },
}

MARKSHEET_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "marksheets")
os.makedirs(MARKSHEET_OUTPUT_DIR, exist_ok=True)


def _type2_variant_config(variant_key):
    """Return variant config dict for Type 2, or None if invalid."""
    variants = CERTIFICATE_TYPES["2"].get("program_title_variants", {})
    return variants.get(variant_key)


def _type2_heading_line1(variant_key):
    """Return first header line for Type 2 (Post Grad vs Diploma)."""
    cfg = _type2_variant_config(variant_key)
    if not cfg:
        return None
    return cfg["heading_line1"]


def _parse_mark_value(value):
    """Return int 0–100 or raise ValueError."""
    s = str(value).strip()
    if not s:
        raise ValueError("empty")
    try:
        n = int(s)
    except ValueError:
        raise ValueError("not an integer")
    if n < 0 or n > 100:
        raise ValueError("out of range")
    return n


def marks_from_entry(entry):
    """Extract marks dict from entry; raises ValueError on invalid mark."""
    marks = {}
    for key in MARK_KEYS:
        marks[key] = _parse_mark_value(entry.get(key, ""))
    return marks


def calc_marksheet_totals(entry_or_marks):
    """Compute total, percentage (out of 1000), and grade B / A+."""
    if isinstance(entry_or_marks, dict) and any(k in entry_or_marks for k in MARK_KEYS):
        marks = marks_from_entry(entry_or_marks)
    else:
        marks = entry_or_marks
    total = sum(marks[k] for k in MARK_KEYS)
    percentage = round(total / 10.0, 1)
    grade = "A+" if total >= 800 else "B"
    return {"total": total, "percentage": percentage, "grade": grade, "marks": marks}


def validate_marksheet_batch(entries):
    """
    Pre-check all marksheet entries before generating any PDF.
    Returns (ok, failures) where failures is a list of dicts with name, gr_no, total.
    """
    failures = []
    for entry in entries:
        try:
            totals = calc_marksheet_totals(entry)
        except ValueError:
            failures.append(
                {
                    "name": entry.get("name", ""),
                    "gr_no": entry.get("gr_no", ""),
                    "total": None,
                    "error": "invalid marks",
                }
            )
            continue
        if totals["total"] < MARKSHEET_MIN_TOTAL:
            failures.append(
                {
                    "name": entry.get("name", ""),
                    "gr_no": entry.get("gr_no", ""),
                    "total": totals["total"],
                }
            )
    return len(failures) == 0, failures


def format_marksheet_batch_failure_message(failures):
    lines = [
        "Cannot generate marksheets. The following students have total marks below 750:",
        "",
    ]
    for f in failures:
        name = f.get("name", "")
        gr = f.get("gr_no", "")
        total = f.get("total")
        if total is None:
            lines.append(f"  • {name} ({gr}) — invalid marks")
        else:
            lines.append(f"  • {name} ({gr}) — {total}")
    lines.append("")
    lines.append("Fix marks and try again. No marksheets were generated or sent.")
    return "\n".join(lines)


def _marksheet_subject_rows(marks):
    theory = []
    practical = []
    for subj in MARKSHEET_SUBJECTS:
        row = {
            "name": subj["name"],
            "max": 100,
            "min": 35,
            "obtained": marks[subj["key"]],
        }
        if subj["section"] == "theory":
            theory.append(row)
        else:
            practical.append(row)
    return theory, practical


ALL_ENTRY_KEYS = (
    "name",
    "course",
    "month",
    "email",
    "gr_no",
    "year",
    "grade",
    "total",
    "percentage",
    "date_of_completion",
    "scheduled_time",
) + MARK_KEYS

# Save/Load list path (next to exe or in project root)
ENTRIES_JSON = os.path.join(EXE_DIR, "certificate_entries.json")
# History of sent certificates (written after each successful send)
HISTORY_JSON = os.path.join(EXE_DIR, "certificate_history.json")


def load_history():
    """Load history from certificate_history.json. Returns list of record dicts."""
    if not os.path.isfile(HISTORY_JSON):
        return []
    try:
        with open(HISTORY_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def append_to_history(record):
    """Append one sent record to history file. record: dict with sent_at, name, email, course, month, etc."""
    history = load_history()
    history.append(record)
    try:
        with open(HISTORY_JSON, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2)
    except Exception:
        pass


def save_history(records):
    """Overwrite history file with full list (e.g. for Clear history)."""
    try:
        with open(HISTORY_JSON, "w", encoding="utf-8") as f:
            json.dump(records, f, indent=2)
    except Exception:
        pass


# Map normalized CSV/Excel header names to our entry keys
_HEADER_TO_KEY = {
    "name": "name",
    "email": "email",
    "e mail": "email",
    "course": "course",
    "month": "month",
    "date of completion": "date_of_completion",
    "date_of_completion": "date_of_completion",
    "scheduled time": "scheduled_time",
    "scheduled_time": "scheduled_time",
    "gr no": "gr_no",
    "gr.no.": "gr_no",
    "gr no.": "gr_no",
    "gr_no": "gr_no",
    "gr number": "gr_no",
    "year": "year",
    "grade": "grade",
    "program title": "program_title",
    "program_title": "program_title",
}
for _mi, _mk in enumerate(MARK_KEYS, start=1):
    _HEADER_TO_KEY[f"m{_mi}"] = _mk


def _normalize_header(h):
    """Strip, lower, and normalize spaces/underscores for column matching."""
    if not h:
        return ""
    s = str(h).strip().lower()
    s = s.replace("_", " ").replace("-", " ")
    return " ".join(s.split())


def _format_cell_for_display(key, val):
    """
    When loading from Excel, cells may be date/datetime objects. Format them
    so Month shows as 'March 2026' and Date of Completion / Scheduled Time stay readable.
    """
    if val is None or (isinstance(val, str) and not val.strip()):
        return ""
    if isinstance(val, str):
        return val.strip()
    # time-only (e.g. Excel time column)
    if hasattr(val, "hour") and not hasattr(val, "year"):
        if key == "scheduled_time":
            return val.strftime("%H:%M")
        return str(val).strip()
    # datetime or date
    if isinstance(val, datetime):
        if key == "month":
            return val.strftime("%B %Y")  # e.g. March 2026
        if key == "date_of_completion":
            return val.strftime("%m/%d/%Y")  # e.g. 12/04/2025
        if key == "scheduled_time":
            return val.strftime("%H:%M")
    if hasattr(val, "year") and hasattr(val, "month"):  # date but not datetime
        if key == "month":
            return val.strftime("%B %Y")
        if key == "date_of_completion":
            return val.strftime("%m/%d/%Y")
    return str(val).strip()


def _normalize_entry_dict(raw):
    """Normalize a raw dict from JSON/CSV/Excel into all known entry keys."""
    entry = {k: "" for k in ALL_ENTRY_KEYS}
    for k in ALL_ENTRY_KEYS:
        if k in raw and raw[k] is not None:
            entry[k] = str(raw[k]).strip()
    return entry


def load_entries_from_file(path):
    """
    Load list of entries from a file. Supports:
    - .json: array of objects with entry fields (type 1 and/or type 2 columns)
    - .csv: first row = headers, same column names (flexible: Name, Email, etc.)
    - .xlsx: first row = headers, same column names
    Returns list of dicts with all ALL_ENTRY_KEYS (empty string if missing).
    """
    path = path.strip()
    if not path:
        return []
    ext = os.path.splitext(path)[1].lower()
    if ext == ".json":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            raise ValueError("JSON file must contain a list of entries.")
        return [_normalize_entry_dict(e) for e in data]
    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            raw_headers = reader.fieldnames or []
            header_to_key = {}
            for h in raw_headers:
                norm = _normalize_header(h)
                if norm in _HEADER_TO_KEY:
                    header_to_key[h] = _HEADER_TO_KEY[norm]
            rows = []
            for row in reader:
                entry = {k: "" for k in ALL_ENTRY_KEYS}
                for file_col, key in header_to_key.items():
                    if key in entry:
                        val = row.get(file_col, "")
                        entry[key] = str(val).strip() if val is not None else ""
                rows.append(entry)
            return rows
    if ext == ".xlsx":
        try:
            import openpyxl
        except ImportError:
            raise ValueError("Excel support requires the 'openpyxl' package. Install it with: pip install openpyxl")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        if ws is None:
            wb.close()
            return []
        rows_iter = ws.iter_rows(values_only=True)
        header_row = next(rows_iter, None)
        if not header_row:
            wb.close()
            return []
        header_to_col = {}
        for col_idx, cell in enumerate(header_row):
            norm = _normalize_header(cell)
            if norm in _HEADER_TO_KEY:
                header_to_col[_HEADER_TO_KEY[norm]] = col_idx
        rows = []
        for row in rows_iter:
            if not any(cell is not None and str(cell).strip() for cell in row):
                continue
            entry = {k: "" for k in ALL_ENTRY_KEYS}
            for k in ALL_ENTRY_KEYS:
                idx = header_to_col.get(k)
                val = None if idx is None else row[idx]
                entry[k] = _format_cell_for_display(k, val)
            rows.append(entry)
        wb.close()
        return rows
    raise ValueError(f"Unsupported file type: {ext}. Use .json, .csv, or .xlsx (Excel).")


_EMAIL_IN_ROW_RE = re.compile(r"[\w.+-]+@[\w.-]+\.\w+")


def _row_looks_complete(line):
    """True when pasted line likely has all columns (email is last field)."""
    return bool(_EMAIL_IN_ROW_RE.search(line or ""))


def _merge_wrapped_paste_lines(lines):
    """
    Excel/Sheets sometimes wrap one row across multiple lines in the paste box
    (e.g. course title with a line break). Join until we see an email address.
    """
    merged = []
    buf = ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        buf = f"{buf} {ln}".strip() if buf else ln
        if _row_looks_complete(buf):
            merged.append(buf)
            buf = ""
    if buf.strip():
        merged.append(buf.strip())
    return merged


def _paste_raw_rows(text):
    """
    Parse pasted clipboard text into rows of cell strings.
    Handles TAB-separated data (incl. multiline cells) and space-separated rows
    with quoted fields that span lines.
    """
    text = (text or "").strip()
    if not text:
        return []

    if "\t" in text:
        rows = []
        for row in csv.reader(io.StringIO(text), delimiter="\t"):
            cells = [str(c).strip() for c in row]
            if any(cells):
                rows.append(cells)
        return rows

    physical = [ln for ln in text.splitlines() if ln.strip()]
    if not physical:
        return []

    merged_lines = _merge_wrapped_paste_lines(physical)
    rows = []
    for line in merged_lines:
        if '"' in line:
            try:
                cells = shlex.split(line, posix=True)
            except ValueError:
                cells = line.split()
        else:
            try:
                cells = next(csv.reader(io.StringIO(line)))
            except Exception:
                cells = line.split()
        rows.append([str(c).strip() for c in cells if c is not None])
    return rows


def parse_bulk_paste(text, certificate_type="1", first_row_is_header=False):
    """
    Parse pasted text from Google Sheets / Excel / Notepad.
    Type 1 without header: Name, Course, Month, Email
    Type 2 without header: Name, Gr.No., Course, Year, Grade, Email
    Type 3 without header: Name, Gr.No., Year, Email, then 10 marks (M1..M10 order)
    With header: first row is column titles; uses same names as CSV import (M1..M10 for marks).

    Returns (entries, warnings) where entries match the active certificate type keys.
    """
    cert_type = certificate_type if certificate_type in CERTIFICATE_TYPES else "1"
    type_cfg = CERTIFICATE_TYPES[cert_type]
    entry_keys = type_cfg["entry_keys"]
    required_no_email = tuple(k for k in type_cfg["required"] if k != "email")

    text = (text or "").strip()
    if not text:
        return [], ["Nothing to paste."]

    raw_rows = _paste_raw_rows(text)
    if not raw_rows:
        return [], ["No non-empty lines."]

    header_allowed = set(ALL_ENTRY_KEYS) | {"program_title"}

    col_to_key = {}
    start_idx = 0
    if first_row_is_header:
        header_cells = raw_rows[0]
        for i, cell in enumerate(header_cells):
            norm = _normalize_header(cell)
            key = _HEADER_TO_KEY.get(norm)
            if key in header_allowed:
                col_to_key[i] = key
        mapped = set(col_to_key.values())
        if "name" not in mapped or "email" not in mapped:
            first_row_is_header = False
            col_to_key = {}
            start_idx = 0
        else:
            start_idx = 1

    col_count = len(entry_keys)
    entries_out = []
    warnings = []

    for row_idx, cells in enumerate(raw_rows[start_idx:], start=1):
        d = {k: "" for k in entry_keys}
        if col_to_key:
            for i, key in col_to_key.items():
                if key in d and i < len(cells) and cells[i] is not None:
                    d[key] = _pdf_text(cells[i])
        else:
            padded = list(cells) + [""] * max(0, col_count - len(cells))
            if len(padded) > col_count:
                padded = padded[:col_count]
            for idx, key in enumerate(entry_keys):
                val = padded[idx] if idx < len(padded) and padded[idx] else ""
                d[key] = _pdf_text(val)

        name = d.get("name", "")
        email = d.get("email", "")

        if not name and not email:
            continue
        if not name or not email:
            warnings.append(f"Row {row_idx}: skipped — Name and Email are required.")
            continue
        missing = [k for k in required_no_email if not d.get(k)]
        if missing:
            warnings.append(
                f"Row {row_idx}: skipped — missing {', '.join(missing)} ({name!r})."
            )
            continue
        if cert_type == "3":
            try:
                totals = calc_marksheet_totals(d)
                d["total"] = str(totals["total"])
                d["percentage"] = str(totals["percentage"])
                d["grade"] = totals["grade"]
            except ValueError:
                warnings.append(
                    f"Row {row_idx}: skipped — invalid marks ({name!r})."
                )
                continue
        entries_out.append(d)

    return entries_out, warnings


def _image_path(filename):
    return os.path.join(STATIC_DIR, "images", filename)


def _font_css_url():
    """Local @font-face CSS for embedded PDF fonts (Cormorant, Montserrat, Cinzel, Poppins)."""
    path = os.path.join(STATIC_DIR, "fonts", "certificate-fonts.css")
    if os.path.isfile(path):
        return "file:///" + os.path.normpath(path).replace("\\", "/")
    return ""


def _merriweather_font_css_url():
    path = os.path.join(STATIC_DIR, "fonts", "merriweather.css")
    if os.path.isfile(path):
        return "file:///" + os.path.normpath(path).replace("\\", "/")
    return ""


def _pdf_text(value):
    """Normalize text for PDF templates (no stray newlines/tabs from paste or Excel)."""
    if value is None:
        return ""
    s = str(value).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    s = s.replace("\t", " ")
    return " ".join(s.split())


def _image_data_url(path):
    with open(path, "rb") as f:
        return f"data:image/png;base64,{base64.b64encode(f.read()).decode('utf-8')}"


def _load_certificate_images(certificate_type="1"):
    """Load logo, certify badge, signature (and border frame for Type 2 / 3) as base64 data URLs."""
    urls = {}
    for key, filename in (
        ("logo_url", "logo.png"),
        ("certify_url", "certify.png"),
        ("signature_url", "sign.png"),
    ):
        path = _image_path(filename)
        urls[key] = _image_data_url(path)
    border_files = {
        "2": "border_frame.png",
        "3": "border_frame_marksheet.png",
    }
    border_file = border_files.get(certificate_type)
    if border_file:
        border_path = os.path.join(STATIC_DIR, border_file)
        if not os.path.isfile(border_path):
            raise FileNotFoundError(border_path)
        urls["border_frame_url"] = _image_data_url(border_path)
    return urls


def send_email(
    receiver_email,
    certificate_path,
    name,
    course,
    certificate_type="1",
    month=None,
    year=None,
    grade=None,
    total_marks=None,
    percentage=None,
):
    unsubscribe_link = os.getenv("UNSUBSCRIBE_LINK", "https://leveluponline.shop/")
    if certificate_type == "3":
        total_marks = total_marks if total_marks is not None else ""
        percentage = percentage if percentage is not None else ""
        body = f"""
    <html><body>
    <p>Dear {name},</p>
    <p>Please find your marksheet for {year} attached.</p>
    <p>Total marks: {total_marks} / 1000 &nbsp;|&nbsp; Percentage: {percentage}% &nbsp;|&nbsp; Grade: {grade}</p>
    <br><br>
    <p style="font-size:12px;color:gray;">
    If you no longer wish to receive emails, you can <a href="{unsubscribe_link}">unsubscribe here</a>.
    </p>
    </body></html>
    """
        subject = f"Marksheet: {name} — {year}"
    elif certificate_type == "2":
        body = f"""
    <html><body>
    <p>Dear {name},</p>
    <p>Congratulations! You have successfully completed the {course} in the year {year} with grade {grade}.</p>
    <p>Please find your certificate of completion attached.</p>
    <br><br>
    <p style="font-size:12px;color:gray;">
    If you no longer wish to receive emails, you can <a href="{unsubscribe_link}">unsubscribe here</a>.
    </p>
    </body></html>
    """
        subject = f"Certificate of Completion: {course}"
    else:
        body = f"""
    <html><body>
    <p>Dear {name},</p>
    <p>Congratulations! You have successfully completed the {course} course on {month}.</p>
    <p>Please find your certificate attached.</p>
    <br><br>
    <p style="font-size:12px;color:gray;">
    If you no longer wish to receive emails, you can <a href="{unsubscribe_link}">unsubscribe here</a>.
    </p>
    </body></html>
    """
        subject = f"Certificate of Achievement: {course}"
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html"))
    with open(certificate_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(certificate_path)}")
        msg.attach(part)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SENDER_EMAIL, receiver_email, msg.as_string())


def get_pdfkit_config():
    import platform
    if platform.system() == "Windows":
        wk = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
        if os.path.isfile(wk):
            return pdfkit.configuration(wkhtmltopdf=wk)
    return pdfkit.configuration()


def generate_and_send_certificate(
    certificate_type,
    entry,
    program_title=None,
    program_title_variant=None,
    log_callback=None,
):
    """
    Generate PDF certificate and send email. Returns (success: bool, message: str).
    entry: dict with fields required for the given certificate_type.
    program_title_variant: Type 2 only — "post_grad" or "diploma" (required for Type 2).
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    cert_type = certificate_type if certificate_type in CERTIFICATE_TYPES else "1"
    type_cfg = CERTIFICATE_TYPES[cert_type]
    entry = entry or {}

    missing = [k for k in type_cfg["required"] if not str(entry.get(k, "")).strip()]
    if missing:
        return False, f"Missing required field(s): {', '.join(missing)}"

    name = _pdf_text(entry.get("name", ""))
    course = _pdf_text(entry.get("course", ""))
    email = str(entry.get("email", "")).strip()

    try:
        config = get_pdfkit_config()
    except Exception as e:
        return False, f"wkhtmltopdf config failed: {e}"

    try:
        image_urls = _load_certificate_images(cert_type)
    except FileNotFoundError as e:
        return False, f"Missing image: {e}"

    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(type_cfg["template"])
    render_ctx = {
        "name": name,
        "course": course,
        "font_css_url": _font_css_url(),
        **image_urls,
    }

    if cert_type == "1":
        render_ctx["month"] = str(entry.get("month", "")).strip()
    else:
        line1 = _type2_heading_line1(program_title_variant)
        if not line1:
            return False, (
                "Select certificate title: Post Graduation Diploma or Diploma (Type 2)."
            )
        variant_cfg = _type2_variant_config(program_title_variant)
        render_ctx.update(
            {
                "gr_no": _pdf_text(entry.get("gr_no", "")),
                "heading_line1": line1,
                "heading_line2": course,
                "diploma_type_label": variant_cfg["label"],
                "year": _pdf_text(entry.get("year", "")),
                "grade": _pdf_text(entry.get("grade", "")),
                "merriweather_font_css_url": _merriweather_font_css_url(),
            }
        )

    html = template.render(**render_ctx)

    safe_name = name.replace(" ", "_").replace("/", "_")
    if cert_type == "1":
        month = str(entry.get("month", "")).strip()
        output_path = os.path.join(
            OUTPUT_DIR, f"{safe_name}_{course}_{month}.pdf"
        )
    else:
        year = str(entry.get("year", "")).strip()
        safe_course = course.replace("/", "_")
        output_path = os.path.join(
            OUTPUT_DIR, f"{safe_name}_{safe_course}_{year}.pdf"
        )

    options = {
        "enable-local-file-access": None,
        "no-stop-slow-scripts": "",
        "quiet": "",
        "margin-top": "0mm",
        "margin-bottom": "0mm",
        "margin-left": "0mm",
        "margin-right": "0mm",
        "page-width": type_cfg["page_width"],
        "page-height": type_cfg["page_height"],
    }
    if cert_type == "2":
        # Type 2: fill A4 edge-to-edge (no smart shrink)
        options.update(
            {
                "disable-smart-shrinking": None,
                "print-media-type": None,
                "dpi": "96",
                "page-size": "A4",
            }
        )
    else:
        # Type 1: original pdfkit settings (zoom 1.25 in template + smart shrink)
        options["dpi"] = "300"

    try:
        pdfkit.from_string(html, output_path, configuration=config, options=options)
    except Exception as e:
        return False, f"PDF failed: {e}"

    try:
        if cert_type == "2":
            send_email(
                email,
                output_path,
                name,
                course,
                certificate_type="2",
                year=str(entry.get("year", "")).strip(),
                grade=str(entry.get("grade", "")).strip(),
            )
        else:
            send_email(
                email,
                output_path,
                name,
                course,
                certificate_type="1",
                month=str(entry.get("month", "")).strip(),
            )
        log(f"[SENT] {name} ({email})")
        return True, "Sent"
    except Exception as e:
        log(f"[ERROR] {name}: {e}")
        return False, str(e)


def generate_and_send_marksheet(entry, log_callback=None):
    """
    Generate marksheet PDF and send email. Batch must pass validate_marksheet_batch first.
    Returns (success: bool, message: str).
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    entry = entry or {}
    type_cfg = CERTIFICATE_TYPES["3"]

    missing = [k for k in type_cfg["required"] if not str(entry.get(k, "")).strip()]
    if missing:
        return False, f"Missing required field(s): {', '.join(missing)}"

    try:
        totals = calc_marksheet_totals(entry)
    except ValueError as e:
        return False, f"Invalid marks for {entry.get('name', '')}: {e}"

    name = _pdf_text(entry.get("name", ""))
    email = str(entry.get("email", "")).strip()
    year = _pdf_text(entry.get("year", ""))
    gr_no = _pdf_text(entry.get("gr_no", ""))
    marks = totals["marks"]

    try:
        config = get_pdfkit_config()
    except Exception as e:
        return False, f"wkhtmltopdf config failed: {e}"

    try:
        image_urls = _load_certificate_images("3")
    except FileNotFoundError as e:
        return False, f"Missing image: {e}"

    theory_rows, practical_rows = _marksheet_subject_rows(marks)
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(type_cfg["template"])
    html = template.render(
        name=name,
        gr_no=gr_no,
        year=year,
        subtitle=type_cfg["subtitle"],
        theory_subjects=theory_rows,
        practical_subjects=practical_rows,
        total_marks=totals["total"],
        out_of=MARKSHEET_OUT_OF,
        percentage=totals["percentage"],
        grade=totals["grade"],
        font_css_url=_font_css_url(),
        **image_urls,
    )

    safe_name = name.replace(" ", "_").replace("/", "_")
    safe_gr = gr_no.replace(" ", "_").replace("/", "_")
    output_path = os.path.join(
        MARKSHEET_OUTPUT_DIR,
        f"{safe_name}_{safe_gr}_Marksheet_{year}.pdf",
    )

    options = {
        "enable-local-file-access": None,
        "no-stop-slow-scripts": "",
        "quiet": "",
        "margin-top": "0mm",
        "margin-bottom": "0mm",
        "margin-left": "0mm",
        "margin-right": "0mm",
        "page-width": type_cfg["page_width"],
        "page-height": type_cfg["page_height"],
        "disable-smart-shrinking": None,
        "print-media-type": None,
        "dpi": "96",
    }

    try:
        pdfkit.from_string(html, output_path, configuration=config, options=options)
    except Exception as e:
        return False, f"PDF failed: {e}"

    try:
        send_email(
            email,
            output_path,
            name,
            "",
            certificate_type="3",
            year=year,
            grade=totals["grade"],
            total_marks=totals["total"],
            percentage=totals["percentage"],
        )
        log(f"[SENT] {name} ({email})")
        return True, "Sent"
    except Exception as e:
        log(f"[ERROR] {name}: {e}")
        return False, str(e)


# ---------------------------------------------------------------------------
# GUI — LevelUp branding
# ---------------------------------------------------------------------------
LEVELUP_COLORS = {
    "teal": "#0a6b6b",
    "teal_dark": "#085656",
    "gold": "#c9a227",
    "gold_light": "#e8d5a3",
    "navy": "#1D3557",
    "bg": "#f5f9f9",
    "white": "#ffffff",
}


def _setup_gui_theme(root):
  import tkinter as tk
  from tkinter import ttk

  c = LEVELUP_COLORS
  style = ttk.Style(root)
  try:
    style.theme_use("clam")
  except tk.TclError:
    pass

  bg = c["bg"]
  style.configure(".", background=bg, foreground=c["navy"])
  style.configure("TFrame", background=bg)
  style.configure("TLabel", background=bg, foreground=c["navy"])
  style.configure("TNotebook", background=bg, borderwidth=0)
  style.configure("TNotebook.Tab", padding=(14, 8), font=("Segoe UI", 10, "bold"))
  style.map(
    "TNotebook.Tab",
    background=[("selected", c["teal"]), ("!selected", bg)],
    foreground=[("selected", c["white"]), ("!selected", c["navy"])],
  )
  style.configure("TLabelframe", background=bg)
  style.configure(
    "TLabelframe.Label",
    background=bg,
    foreground=c["teal_dark"],
    font=("Segoe UI", 10, "bold"),
  )
  style.configure("TRadiobutton", background=bg, foreground=c["navy"])
  style.configure("TCheckbutton", background=bg, foreground=c["navy"])
  style.configure("TButton", font=("Segoe UI", 9), padding=(10, 5))
  style.configure(
    "Accent.TButton",
    background=c["teal"],
    foreground=c["white"],
    font=("Segoe UI", 10, "bold"),
    padding=(12, 6),
  )
  style.map(
    "Accent.TButton",
    background=[("active", c["teal_dark"]), ("pressed", c["teal_dark"])],
    foreground=[("active", c["white"]), ("pressed", c["white"])],
  )
  style.configure(
    "Treeview",
    background=c["white"],
    fieldbackground=c["white"],
    foreground=c["navy"],
    rowheight=24,
  )
  style.configure(
    "Treeview.Heading",
    background=c["teal"],
    foreground=c["white"],
    font=("Segoe UI", 9, "bold"),
  )
  style.map(
    "Treeview",
    background=[("selected", c["teal"])],
    foreground=[("selected", c["white"])],
  )
  root.configure(bg=bg)


def _load_header_logo(root, max_height=58):
  import tkinter as tk

  path = _image_path("logo.png")
  if not os.path.isfile(path):
    return None
  try:
    img = tk.PhotoImage(file=path)
    h = img.height()
    if h > max_height:
      factor = max(1, round(h / max_height))
      img = img.subsample(factor, factor)
    root._header_logo_img = img
    return img
  except tk.TclError:
    return None


def _build_app_header(root):
  import tkinter as tk

  c = LEVELUP_COLORS
  wrapper = tk.Frame(root, bg=c["teal_dark"], highlightthickness=0)
  wrapper.pack(side=tk.TOP, fill=tk.X)

  header = tk.Frame(wrapper, bg=c["teal"], height=72)
  header.pack(fill=tk.X)
  header.pack_propagate(False)

  logo = _load_header_logo(root)
  if logo:
    tk.Label(header, image=logo, bg=c["teal"]).pack(side=tk.LEFT, padx=(14, 10), pady=6)

  text_col = tk.Frame(header, bg=c["teal"])
  text_col.pack(side=tk.LEFT, fill=tk.Y, pady=10)
  tk.Label(
    text_col,
    text="LEVELUP ONLINE EDUCATION",
    bg=c["teal"],
    fg=c["white"],
    font=("Segoe UI", 13, "bold"),
  ).pack(anchor=tk.W)
  tk.Label(
    text_col,
    text="Certificate Auto",
    bg=c["teal"],
    fg=c["gold_light"],
    font=("Segoe UI", 10),
  ).pack(anchor=tk.W, pady=(2, 0))

  tk.Frame(wrapper, bg=c["gold"], height=3).pack(fill=tk.X)


def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    form_widgets = {}
    tree = None
    list_frame = None
    entry_store = {}
    last_cert_type = ["1"]

    def get_certificate_type():
        t = certificate_type_var.get()
        return t if t in CERTIFICATE_TYPES else "1"

    def entry_keys_for_type(cert_type=None):
        return CERTIFICATE_TYPES[cert_type or get_certificate_type()]["entry_keys"]

    def tree_values_from_entry(entry, cert_type=None):
        cert_type = cert_type or get_certificate_type()
        if cert_type == "3":
            return (
                entry.get("name", ""),
                entry.get("gr_no", ""),
                str(entry.get("total", "")),
                entry.get("grade", ""),
                entry.get("email", ""),
            )
        keys = entry_keys_for_type(cert_type)
        return tuple(entry.get(k, "") for k in keys)

    def insert_tree_row(entry, cert_type=None):
        cert_type = cert_type or get_certificate_type()
        iid = tree.insert("", tk.END, values=tree_values_from_entry(entry, cert_type))
        if cert_type == "3":
            entry_store[iid] = dict(entry)
        return iid

    def entry_from_tree_values(vals, cert_type=None):
        keys = entry_keys_for_type(cert_type)
        entry = {k: "" for k in keys}
        for i, k in enumerate(keys):
            if i < len(vals):
                entry[k] = vals[i] if vals[i] is not None else ""
        return entry

    def rebuild_tree_columns():
        nonlocal tree
        cert_type = get_certificate_type()
        cols = CERTIFICATE_TYPES[cert_type]["tree_columns"]
        saved = []
        for row_id in tree.get_children():
            vals = tree.item(row_id, "values")
            saved.append(vals)
        tree.configure(columns=cols)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100)
        tree.column("Name", width=120)
        if "Email" in cols:
            tree.column("Email", width=180)
        if "Gr.No." in cols:
            tree.column("Gr.No.", width=90)
        if "Total" in cols:
            tree.column("Total", width=70)
        entry_store.clear()
        for row_id in tree.get_children():
            tree.delete(row_id)

    def update_form_visibility():
        cert_type = get_certificate_type()
        for key, widgets in form_widgets.items():
            visible = key in entry_keys_for_type(cert_type)
            for w in widgets:
                if visible:
                    w.grid()
                else:
                    w.grid_remove()
        if cert_type == "2":
            if not program_title_option_frame.winfo_ismapped():
                program_title_option_frame.pack(
                    fill=tk.X, padx=4, pady=4, before=entry_row_frame
                )
        else:
            program_title_option_frame.pack_forget()
        if cert_type == "3":
            marks_panel_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
            form.grid(row=0, column=1, sticky="nsew")
            root.after_idle(_refresh_marks_scrollregion)
        else:
            marks_panel_frame.grid_remove()
            form.grid(row=0, column=0, columnspan=2, sticky="ew")

    def on_certificate_type_changed():
        if tree is None:
            return
        new_type = certificate_type_var.get()
        if tree.get_children():
            if not messagebox.askyesno(
                "Switch document type",
                "Switching type will clear the entry list. Continue?",
            ):
                certificate_type_var.set(last_cert_type[0])
                return
            entry_store.clear()
            for i in tree.get_children():
                tree.delete(i)
        last_cert_type[0] = new_type
        rebuild_tree_columns()
        if get_certificate_type() != "2":
            program_title_variant_var.set("")
        update_form_visibility()

    def toggle_console(show: bool):
        if sys.platform != "win32":
            return
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            con = kernel32.GetConsoleWindow()
            if con:
                kernel32.ShowWindow(con, 1 if show else 0)
        except Exception:
            pass

    def add_entry():
        cert_type = get_certificate_type()
        if cert_type == "3":
            entry = {
                "name": field_vars["name"].get().strip(),
                "gr_no": field_vars["gr_no"].get().strip(),
                "year": field_vars["year"].get().strip(),
                "email": field_vars["email"].get().strip(),
            }
            if not entry.get("name") or not entry.get("email"):
                messagebox.showwarning("Add entry", "Name and Email are required.")
                return
            if not entry.get("gr_no") or not entry.get("year"):
                messagebox.showwarning("Add entry", "Gr.No. and Year are required.")
                return
            try:
                for key in MARK_KEYS:
                    entry[key] = str(_parse_mark_value(mark_vars[key].get()))
            except ValueError:
                messagebox.showwarning(
                    "Add entry",
                    "Each subject mark must be a whole number from 0 to 100.",
                )
                return
            totals = calc_marksheet_totals(entry)
            entry["total"] = str(totals["total"])
            entry["percentage"] = str(totals["percentage"])
            entry["grade"] = totals["grade"]
            insert_tree_row(entry, cert_type)
            for k in ("name", "gr_no", "year", "email"):
                field_vars[k].set("")
            for key in MARK_KEYS:
                mark_vars[key].set("")
            return

        keys = entry_keys_for_type(cert_type)
        entry = {k: field_vars[k].get().strip() for k in keys}
        if not entry.get("name") or not entry.get("email"):
            messagebox.showwarning("Add entry", "Name and Email are required.")
            return
        required = [k for k in CERTIFICATE_TYPES[cert_type]["required"] if k not in ("email",)]
        missing = [k for k in required if not entry.get(k)]
        if missing:
            messagebox.showwarning(
                "Add entry",
                f"Please fill: {', '.join(missing)}",
            )
            return
        tree.insert("", tk.END, values=tree_values_from_entry(entry, cert_type))
        for k in keys:
            field_vars[k].set("")

    def remove_selected():
        sel = tree.selection()
        for i in sel:
            entry_store.pop(i, None)
            tree.delete(i)

    def clear_list():
        if messagebox.askyesno("Clear list", "Remove all entries from the list?"):
            entry_store.clear()
            for i in tree.get_children():
                tree.delete(i)

    def get_entries():
        cert_type = get_certificate_type()
        if cert_type == "3":
            return [entry_store[iid].copy() for iid in tree.get_children() if iid in entry_store]
        out = []
        for row in tree.get_children():
            vals = tree.item(row, "values")
            if vals:
                out.append(entry_from_tree_values(vals, cert_type))
        return out

    def save_list():
        data = get_entries()
        for row in data:
            row["certificate_type"] = get_certificate_type()
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir=EXE_DIR,
            initialfile="certificate_entries.json",
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            messagebox.showinfo("Save", f"Saved {len(data)} entries to:\n{path}")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

    def load_list():
        path = filedialog.askopenfilename(
            filetypes=[
                ("All supported", "*.json;*.csv;*.xlsx"),
                ("JSON", "*.json"),
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx"),
            ],
            initialdir=EXE_DIR,
        )
        if not path:
            return
        try:
            data = load_entries_from_file(path)
            cert_type = get_certificate_type()
            keys = entry_keys_for_type(cert_type)
            entry_store.clear()
            for i in tree.get_children():
                tree.delete(i)
            loaded = 0
            for e in data:
                entry = {k: e.get(k, "") for k in keys}
                if not entry.get("name") and not entry.get("email"):
                    continue
                missing = [
                    k
                    for k in CERTIFICATE_TYPES[cert_type]["required"]
                    if k != "email" and not entry.get(k)
                ]
                if missing:
                    continue
                if cert_type == "3":
                    try:
                        for key in MARK_KEYS:
                            entry[key] = str(_parse_mark_value(entry.get(key, "")))
                        totals = calc_marksheet_totals(entry)
                        entry["total"] = str(totals["total"])
                        entry["percentage"] = str(totals["percentage"])
                        entry["grade"] = totals["grade"]
                    except ValueError:
                        continue
                    insert_tree_row(entry, cert_type)
                else:
                    tree.insert("", tk.END, values=tree_values_from_entry(entry, cert_type))
                loaded += 1
            messagebox.showinfo(
                "Load",
                f"Loaded {loaded} entries from {os.path.basename(path)}.",
            )
        except Exception as e:
            messagebox.showerror("Load failed", str(e))

    def generate_and_send():
        entries = get_entries()
        if not entries:
            messagebox.showwarning("Generate & Send", "Add at least one entry.")
            return
        status_var.set("Sending...")
        root.update()
        sent = 0
        failed = 0
        log_lines = []

        def log(msg):
            log_lines.append(msg)
            log_text.insert(tk.END, msg + "\n")
            log_text.see(tk.END)
            root.update()

        cert_type = get_certificate_type()
        title_variant = None
        if cert_type == "2":
            title_variant = program_title_variant_var.get().strip()
            if title_variant not in CERTIFICATE_TYPES["2"]["program_title_variants"]:
                messagebox.showwarning(
                    "Generate & Send",
                    "Please select a certificate title for Type 2:\n"
                    "Post Graduation Diploma or Diploma.",
                )
                status_var.set("Ready")
                return

        if cert_type == "3":
            ok_batch, failures = validate_marksheet_batch(entries)
            if not ok_batch:
                messagebox.showwarning(
                    "Marksheet batch blocked",
                    format_marksheet_batch_failure_message(failures),
                )
                status_var.set("Ready")
                return
            for e in entries:
                ok, msg = generate_and_send_marksheet(e, log_callback=log)
                if ok:
                    sent += 1
                    record = {
                        "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "certificate_type": "3",
                        "name": e.get("name", ""),
                        "email": e.get("email", ""),
                        "course": "Marksheet",
                        "gr_no": e.get("gr_no", ""),
                        "year": e.get("year", ""),
                        "grade": e.get("grade", ""),
                        "total_marks": e.get("total", ""),
                        "percentage": e.get("percentage", ""),
                    }
                    append_to_history(record)
                else:
                    failed += 1
        else:
            for e in entries:
                ok, msg = generate_and_send_certificate(
                    cert_type,
                    e,
                    program_title_variant=title_variant,
                    log_callback=log,
                )
                if ok:
                    sent += 1
                    record = {
                        "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "certificate_type": cert_type,
                        "name": e.get("name", ""),
                        "email": e.get("email", ""),
                        "course": e.get("course", ""),
                    }
                    if cert_type == "1":
                        record["month"] = e.get("month", "")
                    else:
                        record["gr_no"] = e.get("gr_no", "")
                        record["year"] = e.get("year", "")
                        record["grade"] = e.get("grade", "")
                        h1 = _type2_heading_line1(title_variant)
                        record["program_title"] = f"{h1} {e.get('course', '')}".strip()
                    append_to_history(record)
                else:
                    failed += 1

        status_var.set(f"Done. Sent: {sent}, Failed: {failed}")
        if refresh_history_ui[0]:
            refresh_history_ui[0]()

        messagebox.showinfo("Generate & Send", f"Sent: {sent}\nFailed: {failed}")

    # Mutable ref so History tab can register its refresh (called after send)
    refresh_history_ui = [None]

    root = tk.Tk()
    root.title("Certificate Auto — LevelUp")
    root.minsize(700, 550)
    root.geometry("800x600")
    _setup_gui_theme(root)
    _build_app_header(root)

    certificate_type_var = tk.StringVar(master=root, value="1")
    program_title_variant_var = tk.StringVar(master=root, value="")
    field_vars = {
        "name": tk.StringVar(master=root),
        "course": tk.StringVar(master=root),
        "month": tk.StringVar(master=root),
        "email": tk.StringVar(master=root),
        "gr_no": tk.StringVar(master=root),
        "year": tk.StringVar(master=root),
        "grade": tk.StringVar(master=root),
    }

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

    # ---------- Tab 1: Entries ----------
    tab_entries = ttk.Frame(notebook, padding=4)
    notebook.add(tab_entries, text="Entries")

    # Pinned footer: action buttons + status + log always visible on small screens
    entries_footer = ttk.Frame(tab_entries)
    entries_footer.pack(side=tk.BOTTOM, fill=tk.X)

    log_frame = ttk.LabelFrame(tab_entries, text="Log", padding=4)
    log_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=4, pady=(0, 4))
    log_text = tk.Text(
        log_frame,
        height=3,
        wrap=tk.WORD,
        state=tk.NORMAL,
        bg=LEVELUP_COLORS["white"],
        fg=LEVELUP_COLORS["navy"],
        relief=tk.FLAT,
        padx=6,
        pady=4,
    )
    log_text.pack(fill=tk.X)

    status_var = tk.StringVar(value="Ready")
    ttk.Label(entries_footer, textvariable=status_var).pack(anchor=tk.W, padx=8, pady=(0, 2))

    btn_frame = ttk.Frame(entries_footer, padding=8)
    btn_frame.pack(fill=tk.X)

    entries_main = ttk.Frame(tab_entries)
    entries_main.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    type_frame = ttk.LabelFrame(entries_main, text="Certificate type (whole batch)", padding=8)
    type_frame.pack(fill=tk.X, padx=4, pady=4)
    ttk.Radiobutton(
        type_frame,
        text="Type 1",
        variable=certificate_type_var,
        value="1",
        command=on_certificate_type_changed,
    ).pack(side=tk.LEFT, padx=(0, 12))
    ttk.Radiobutton(
        type_frame,
        text="Type 2",
        variable=certificate_type_var,
        value="2",
        command=on_certificate_type_changed,
    ).pack(side=tk.LEFT, padx=(0, 12))
    ttk.Radiobutton(
        type_frame,
        text="Type 3",
        variable=certificate_type_var,
        value="3",
        command=on_certificate_type_changed,
    ).pack(side=tk.LEFT, padx=(0, 12))

    entry_row_frame = ttk.Frame(entries_main)
    entry_row_frame.pack(fill=tk.X, padx=4, pady=4)
    entry_row_frame.columnconfigure(0, weight=3)
    entry_row_frame.columnconfigure(1, weight=2)
    entry_row_frame.rowconfigure(0, weight=1)

    mark_vars = {key: tk.StringVar(master=root) for key in MARK_KEYS}

    marks_panel_frame = ttk.LabelFrame(
        entry_row_frame,
        text="Subject marks (Type 3) * — enter 0 to 100 for each subject",
        padding=8,
    )
    marks_scroll_wrap = ttk.Frame(marks_panel_frame)
    marks_scroll_wrap.pack(fill=tk.BOTH, expand=True)
    marks_canvas = tk.Canvas(marks_scroll_wrap, height=180, highlightthickness=0)
    marks_scrollbar = ttk.Scrollbar(
        marks_scroll_wrap, orient=tk.VERTICAL, command=marks_canvas.yview
    )
    marks_inner = ttk.Frame(marks_canvas)
    marks_canvas.configure(yscrollcommand=marks_scrollbar.set)
    marks_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    marks_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    marks_inner.columnconfigure(0, weight=1)

    marks_window_id = marks_canvas.create_window((0, 0), window=marks_inner, anchor=tk.NW)

    def _refresh_marks_scrollregion(_event=None):
        marks_canvas.update_idletasks()
        bbox = marks_canvas.bbox("all")
        if bbox:
            marks_canvas.configure(scrollregion=bbox)

    def _on_marks_canvas_configure(event):
        marks_canvas.itemconfigure(marks_window_id, width=event.width)

    def _on_marks_mousewheel(event):
        if event.delta:
            marks_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            marks_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            marks_canvas.yview_scroll(1, "units")
        return "break"

    def _bind_marks_mousewheel_recursive(widget):
        widget.bind("<MouseWheel>", _on_marks_mousewheel)
        widget.bind("<Button-4>", _on_marks_mousewheel)
        widget.bind("<Button-5>", _on_marks_mousewheel)
        for child in widget.winfo_children():
            _bind_marks_mousewheel_recursive(child)

    marks_inner.bind("<Configure>", _refresh_marks_scrollregion)
    marks_canvas.bind("<Configure>", _on_marks_canvas_configure)
    marks_canvas.bind("<MouseWheel>", _on_marks_mousewheel)
    marks_scrollbar.bind("<MouseWheel>", _on_marks_mousewheel)

    def _add_mark_rows(section_title, subjects, start_row):
        ttk.Label(
            marks_inner,
            text=section_title,
            font=("Segoe UI", 9, "bold"),
            foreground=LEVELUP_COLORS["teal_dark"],
        ).grid(row=start_row, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))
        row = start_row + 1
        for subj in subjects:
            ttk.Label(
                marks_inner,
                text=subj["name"],
                wraplength=300,
                justify=tk.LEFT,
            ).grid(row=row, column=0, sticky=tk.W, padx=(0, 12), pady=3)
            ttk.Entry(
                marks_inner, textvariable=mark_vars[subj["key"]], width=8
            ).grid(row=row, column=1, sticky=tk.E, pady=3)
            row += 1
        return row

    theory_subjects = [s for s in MARKSHEET_SUBJECTS if s["section"] == "theory"]
    practical_subjects = [s for s in MARKSHEET_SUBJECTS if s["section"] == "practical"]
    next_row = _add_mark_rows("THEORY SUBJECTS", theory_subjects, 0)
    _add_mark_rows("PRACTICAL SUBJECTS", practical_subjects, next_row + 1)
    _bind_marks_mousewheel_recursive(marks_inner)
    root.after_idle(_refresh_marks_scrollregion)

    program_title_option_frame = ttk.LabelFrame(
        entries_main,
        text="Certificate title (Type 2) * — select before Generate & Send",
        padding=8,
    )
    ttk.Radiobutton(
        program_title_option_frame,
        text="Certificate of Completion in Post Graduation Diploma",
        variable=program_title_variant_var,
        value="post_grad",
    ).pack(anchor=tk.W, pady=2)
    ttk.Radiobutton(
        program_title_option_frame,
        text="Certificate of Completion in Diploma",
        variable=program_title_variant_var,
        value="diploma",
    ).pack(anchor=tk.W, pady=2)

    form = ttk.LabelFrame(entry_row_frame, text="Add entry", padding=8)

    def _add_field(parent, row, col, label, key, width=18, colspan=1):
        lbl = ttk.Label(parent, text=label)
        ent = ttk.Entry(parent, textvariable=field_vars[key], width=width)
        lbl.grid(row=row, column=col, sticky=tk.W, padx=(0, 4))
        ent.grid(row=row, column=col + 1, columnspan=colspan, padx=(0, 12), sticky=tk.EW)
        form_widgets[key] = [lbl, ent]
        return ent

    _add_field(form, 0, 0, "Name *", "name", width=22)
    _add_field(form, 0, 2, "Course *", "course", width=18)
    _add_field(form, 0, 4, "Month *", "month", width=14)
    _add_field(form, 1, 0, "Gr.No. *", "gr_no", width=14)
    _add_field(form, 1, 2, "Year *", "year", width=10)
    _add_field(form, 1, 4, "Grade *", "grade", width=10)
    email_lbl = ttk.Label(form, text="Email *")
    email_ent = ttk.Entry(form, textvariable=field_vars["email"], width=40)
    email_lbl.grid(row=2, column=0, sticky=tk.W, padx=(0, 4))
    email_ent.grid(row=2, column=1, columnspan=5, padx=(0, 12), sticky=tk.EW)
    form_widgets["email"] = [email_lbl, email_ent]
    ttk.Button(form, text="Add to list", command=add_entry).grid(row=2, column=6, padx=(0, 6))
    paste_btn = ttk.Button(form, text="Paste rows…")
    paste_btn.grid(row=2, column=7, padx=(0, 0))

    list_frame = ttk.LabelFrame(entries_main, text="Entries to send", padding=8)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    init_cols = CERTIFICATE_TYPES["1"]["tree_columns"]
    tree = ttk.Treeview(list_frame, columns=init_cols, show="headings", height=5, selectmode="extended")
    for c in init_cols:
        tree.heading(c, text=c)
        tree.column(c, width=100)
    tree.column("Name", width=120)
    tree.column("Email", width=180)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=scroll.set)
    update_form_visibility()

    def open_paste_rows_dialog():
        # ✅ Bulk-add: paste TSV/CSV from Sheets, Excel, or Notepad
        win = tk.Toplevel(root)
        win.title("Paste rows")
        win.minsize(520, 420)
        win.geometry("640x480")

        header_var = tk.BooleanVar(value=False)

        paste_type = get_certificate_type()
        if paste_type == "3":
            help_txt = (
                "Paste from Google Sheets or Excel (select cells → Copy): columns are TAB-separated.\n"
                "Or type one person per line. Without a header row, use this order:\n"
                "Name — Gr.No. — Year — Email — M1 — M2 — … — M10 (TAB-separated from Sheets)\n"
                "M1–M6 = theory subjects, M7–M10 = practical (each mark 0–100)\n"
                "Tick “First row is column names” if row 1 has headers like Name, GR No, M1, …"
            )
        elif paste_type == "2":
            help_txt = (
                "Paste from Google Sheets or Excel (select cells → Copy): columns are TAB-separated.\n"
                "Or type one person per line. Without a header row, use this order:\n"
                "Name — Gr.No. — Course — Year — Grade — Email\n"
                "A long course name may wrap to a second line — that is OK.\n"
                "Tick “First row is column names” if the first line is titles like Name, Email, …"
            )
        else:
            help_txt = (
                "Paste from Google Sheets or Excel (select cells → Copy): columns are TAB-separated.\n"
                "Or type one person per line. Without a header row, use this order:\n"
                "Name — Course — Month — Email\n"
                "Wrapped text inside one row is OK.\n"
                "Tick “First row is column names” if the first line is titles like Name, Email, …"
            )
        ttk.Label(win, text=help_txt, wraplength=600, justify=tk.LEFT).pack(
            anchor=tk.W, padx=10, pady=(10, 4)
        )

        body = ttk.Frame(win, padding=(10, 4))
        body.pack(fill=tk.BOTH, expand=True)
        paste_text = tk.Text(body, height=16, width=80, wrap=tk.NONE)
        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=paste_text.yview)
        hsb = ttk.Scrollbar(body, orient=tk.HORIZONTAL, command=paste_text.xview)
        paste_text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        paste_text.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            win,
            text="First row is column names (header)",
            variable=header_var,
        ).pack(anchor=tk.W, padx=10, pady=4)

        btn_row = ttk.Frame(win, padding=10)
        btn_row.pack(fill=tk.X)

        def paste_from_clipboard():
            try:
                clip = root.clipboard_get()
            except tk.TclError:
                messagebox.showwarning(
                    "Clipboard",
                    "Could not read the clipboard. Copy cells in Sheets/Excel first, then try again.",
                    parent=win,
                )
                return
            # ✅ Replace box so a second “Paste from clipboard” doesn’t duplicate data
            paste_text.delete("1.0", tk.END)
            paste_text.insert(tk.END, clip)

        def clear_paste_box():
            paste_text.delete("1.0", tk.END)

        def do_add_to_list():
            raw = paste_text.get("1.0", tk.END)
            entries_parsed, warns = parse_bulk_paste(
                raw,
                certificate_type=get_certificate_type(),
                first_row_is_header=header_var.get(),
            )
            cert_type = get_certificate_type()
            for e in entries_parsed:
                if cert_type == "3":
                    try:
                        for key in MARK_KEYS:
                            e[key] = str(_parse_mark_value(e.get(key, "")))
                        totals = calc_marksheet_totals(e)
                        e["total"] = str(totals["total"])
                        e["percentage"] = str(totals["percentage"])
                        e["grade"] = totals["grade"]
                    except ValueError:
                        warns.append(f"Skipped {e.get('name', '')!r}: invalid marks.")
                        continue
                    insert_tree_row(e, cert_type)
                else:
                    tree.insert(
                        "",
                        tk.END,
                        values=tree_values_from_entry(e, cert_type),
                    )
            parts = [f"Added {len(entries_parsed)} row(s) to the list."]
            if warns:
                show_warns = warns[:25]
                parts.append("")
                parts.extend(show_warns)
                if len(warns) > 25:
                    parts.append(f"… and {len(warns) - 25} more warning(s).")
            msg = "\n".join(parts)
            if warns:
                messagebox.showwarning("Paste rows", msg, parent=win)
            else:
                messagebox.showinfo("Paste rows", msg, parent=win)
            if entries_parsed:
                win.destroy()

        ttk.Button(
            btn_row, text="Paste from clipboard", command=paste_from_clipboard
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Clear box", command=clear_paste_box).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Add to list", command=do_add_to_list).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Close", command=win.destroy).pack(side=tk.LEFT)

    paste_btn.configure(command=open_paste_rows_dialog)

    ttk.Button(btn_frame, text="Remove selected", command=remove_selected).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="Clear list", command=clear_list).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="Save list...", command=save_list).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="Load list...", command=load_list).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(
        btn_frame,
        text="Generate & Send",
        command=generate_and_send,
        style="Accent.TButton",
    ).pack(side=tk.LEFT, padx=(0, 6))

    # Only show console toggle when running from Python (not from built exe); exe is built windowed so no console
    if sys.platform == "win32" and not getattr(sys, "frozen", False):
        ttk.Button(btn_frame, text="Hide console", command=lambda: toggle_console(False)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text="Show console", command=lambda: toggle_console(True)).pack(side=tk.LEFT, padx=(0, 6))

    # ---------- Tab 2: History ----------
    tab_history = ttk.Frame(notebook, padding=4)
    notebook.add(tab_history, text="History")

    history_summary_var = tk.StringVar(value="Total certificates sent: 0")
    history_filter_var = tk.StringVar(value="all")  # all, today, week, month
    history_search_var = tk.StringVar()

    def _parse_sent_at(sent_at_str):
        """Parse 'YYYY-MM-DD HH:MM:SS' to date, or None if invalid."""
        if not sent_at_str:
            return None
        s = str(sent_at_str).strip()[:10]
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except ValueError:
            return None

    def _apply_history_filters():
        """Load history, apply date filter + search, refresh tree and summary."""
        records = load_history()
        total_count = len(records)

        # Date filter
        filter_choice = history_filter_var.get()
        today = date.today()
        if filter_choice == "today":
            records = [r for r in records if _parse_sent_at(r.get("sent_at")) == today]
        elif filter_choice == "week":
            week_start = today - timedelta(days=6)
            records = [r for r in records if week_start <= (_parse_sent_at(r.get("sent_at")) or date.min) <= today]
        elif filter_choice == "month":
            records = [r for r in records if _parse_sent_at(r.get("sent_at")) and _parse_sent_at(r.get("sent_at")).month == today.month and _parse_sent_at(r.get("sent_at")).year == today.year]

        history_search_keys = (
            "sent_at",
            "name",
            "email",
            "course",
            "month",
            "certificate_type",
            "gr_no",
            "year",
            "grade",
            "program_title",
            "total_marks",
            "percentage",
        )
        search_text = (history_search_var.get() or "").strip().lower()
        if search_text:
            def matches(r):
                for key in history_search_keys:
                    if search_text in (str(r.get(key, "") or "").lower()):
                        return True
                return False
            records = [r for r in records if matches(r)]

        for i in history_tree.get_children():
            history_tree.delete(i)
        for r in reversed(records):
            ctype = str(r.get("certificate_type", "1") or "1")
            if ctype == "1":
                month_year = r.get("month", "")
                gr_no = ""
                grade = ""
            elif ctype == "3":
                month_year = r.get("year", "")
                gr_no = r.get("gr_no", "")
                grade = r.get("grade", "")
            else:
                month_year = r.get("year", "")
                gr_no = r.get("gr_no", "")
                grade = r.get("grade", "")
            course_val = r.get("course", "") if ctype != "3" else "Marksheet"
            history_tree.insert("", tk.END, values=(
                r.get("sent_at", ""),
                ctype,
                r.get("name", ""),
                r.get("email", ""),
                course_val,
                month_year,
                gr_no,
                grade,
            ))
        if total_count == len(records) and not search_text:
            history_summary_var.set(f"Total certificates sent: {total_count}")
        else:
            history_summary_var.set(f"Showing {len(records)} of {total_count} certificates sent")

    def refresh_history_tab():
        _apply_history_filters()

    def _on_history_search_change(*args):
        _apply_history_filters()

    def clear_history():
        if not messagebox.askyesno("Clear history", "Remove all history entries? This cannot be undone."):
            return
        save_history([])
        refresh_history_tab()

    def export_history_csv():
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialdir=EXE_DIR,
            initialfile="certificate_history.csv",
        )
        if not path:
            return
        records = load_history()
        if not records:
            messagebox.showinfo("Export", "No history to export.")
            return
        # Apply same filters as view so export matches what's on screen
        filter_choice = history_filter_var.get()
        today = date.today()
        if filter_choice == "today":
            records = [r for r in records if _parse_sent_at(r.get("sent_at")) == today]
        elif filter_choice == "week":
            week_start = today - timedelta(days=6)
            records = [r for r in records if week_start <= (_parse_sent_at(r.get("sent_at")) or date.min) <= today]
        elif filter_choice == "month":
            records = [r for r in records if _parse_sent_at(r.get("sent_at")) and _parse_sent_at(r.get("sent_at")).month == today.month and _parse_sent_at(r.get("sent_at")).year == today.year]
        search_text = (history_search_var.get() or "").strip().lower()
        history_search_keys = (
            "sent_at",
            "name",
            "email",
            "course",
            "month",
            "certificate_type",
            "gr_no",
            "year",
            "grade",
            "program_title",
            "total_marks",
            "percentage",
        )
        if search_text:
            def matches(r):
                for key in history_search_keys:
                    if search_text in (str(r.get(key, "") or "").lower()):
                        return True
                return False
            records = [r for r in records if matches(r)]
        headers = [
            "Sent at",
            "Type",
            "Name",
            "Email",
            "Course",
            "Month/Year",
            "Gr.No.",
            "Grade",
            "Program title",
        ]

        def _format_sent_at_for_csv(s):
            """Format as readable date and prefix with tab so Excel shows it as text (avoids #####)."""
            if not s:
                return ""
            s = str(s).strip()
            try:
                dt = datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")
                # Tab prefix forces Excel to treat as text so the date is visible, not #####
                return "\t" + dt.strftime("%d-%b-%Y %H:%M")  # e.g. 12-Mar-2026 18:30
            except ValueError:
                return "\t" + s

        # Map record keys (sent_at, name, ...) to CSV column headers; format date for clear display in Excel
        rows_for_csv = []
        for r in records:
            ctype = str(r.get("certificate_type", "1") or "1")
            rows_for_csv.append(
                {
                    "Sent at": _format_sent_at_for_csv(r.get("sent_at", "")),
                    "Type": ctype,
                    "Name": r.get("name", ""),
                    "Email": r.get("email", ""),
                    "Course": r.get("course", "") if ctype != "3" else "Marksheet",
                    "Month/Year": r.get("month", "")
                    if ctype == "1"
                    else r.get("year", ""),
                    "Gr.No.": r.get("gr_no", "") if ctype in ("2", "3") else "",
                    "Grade": r.get("grade", "") if ctype in ("2", "3") else "",
                    "Program title": r.get("program_title", "")
                    if ctype == "2"
                    else "",
                }
            )
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
                w.writeheader()
                w.writerows(rows_for_csv)
            messagebox.showinfo("Export", f"Exported {len(records)} rows to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    ttk.Label(tab_history, textvariable=history_summary_var, font=("", 10, "bold")).pack(anchor=tk.W, padx=4, pady=4)
    history_filter_frame = ttk.Frame(tab_history)
    history_filter_frame.pack(fill=tk.X, padx=4, pady=2)
    ttk.Label(history_filter_frame, text="Filter:").pack(side=tk.LEFT, padx=(0, 6))
    ttk.Radiobutton(history_filter_frame, text="All", variable=history_filter_var, value="all", command=_apply_history_filters).pack(side=tk.LEFT, padx=(0, 4))
    ttk.Radiobutton(history_filter_frame, text="Today", variable=history_filter_var, value="today", command=_apply_history_filters).pack(side=tk.LEFT, padx=(0, 4))
    ttk.Radiobutton(history_filter_frame, text="This week", variable=history_filter_var, value="week", command=_apply_history_filters).pack(side=tk.LEFT, padx=(0, 4))
    ttk.Radiobutton(history_filter_frame, text="This month", variable=history_filter_var, value="month", command=_apply_history_filters).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Label(history_filter_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 4))
    history_search_entry = ttk.Entry(history_filter_frame, textvariable=history_search_var, width=28)
    history_search_entry.pack(side=tk.LEFT, padx=(0, 6))
    history_search_var.trace_add("write", _on_history_search_change)
    history_btn_frame = ttk.Frame(tab_history)
    history_btn_frame.pack(fill=tk.X, padx=4, pady=2)
    ttk.Button(history_btn_frame, text="Refresh", command=refresh_history_tab).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(history_btn_frame, text="Clear history", command=clear_history).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(history_btn_frame, text="Export to CSV", command=export_history_csv).pack(side=tk.LEFT, padx=(0, 6))

    history_cols = (
        "Sent at",
        "Type",
        "Name",
        "Email",
        "Course",
        "Month/Year",
        "Gr.No.",
        "Grade",
    )
    history_tree_frame = ttk.Frame(tab_history)
    history_tree_frame.pack(fill=tk.BOTH, expand=True, pady=4)
    history_tree = ttk.Treeview(history_tree_frame, columns=history_cols, show="headings", height=12, selectmode="extended")
    for c in history_cols:
        history_tree.heading(c, text=c)
        history_tree.column(c, width=90)
    history_tree.column("Sent at", width=150)
    history_tree.column("Name", width=110)
    history_tree.column("Email", width=160)
    history_tree.column("Course", width=120)
    history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    history_scroll = ttk.Scrollbar(history_tree_frame, orient=tk.VERTICAL, command=history_tree.yview)
    history_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    history_tree.configure(yscrollcommand=history_scroll.set)

    refresh_history_ui[0] = refresh_history_tab
    refresh_history_tab()  # Load history on startup

    # When running from Python (not exe), hide console at startup so only the GUI shows
    if sys.platform == "win32" and not getattr(sys, "frozen", False):
        root.after(100, lambda: toggle_console(False))

    root.mainloop()


if __name__ == "__main__":
    run_gui()
