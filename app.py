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
}


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


def load_entries_from_file(path):
    """
    Load list of entries from a file. Supports:
    - .json: array of objects with name, email, course, month, date_of_completion, scheduled_time
    - .csv: first row = headers, same column names (flexible: Name, Email, etc.)
    - .xlsx: first row = headers, same column names
    Returns list of dicts with keys name, course, month, date_of_completion, scheduled_time, email.
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
        return [
            {
                "name": str(e.get("name", "")).strip(),
                "course": str(e.get("course", "")).strip(),
                "month": str(e.get("month", "")).strip(),
                "date_of_completion": str(e.get("date_of_completion", "")).strip(),
                "scheduled_time": str(e.get("scheduled_time", "")).strip(),
                "email": str(e.get("email", "")).strip(),
            }
            for e in data
        ]
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
                entry = {k: "" for k in ("name", "course", "month", "date_of_completion", "scheduled_time", "email")}
                for file_col, key in header_to_key.items():
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
        keys = ("name", "course", "month", "date_of_completion", "scheduled_time", "email")
        rows = []
        for row in rows_iter:
            if not any(cell is not None and str(cell).strip() for cell in row):
                continue
            entry = {}
            for k in keys:
                idx = header_to_col.get(k)
                val = None if idx is None else row[idx]
                entry[k] = _format_cell_for_display(k, val)
            rows.append(entry)
        wb.close()
        return rows
    raise ValueError(f"Unsupported file type: {ext}. Use .json, .csv, or .xlsx (Excel).")


def parse_bulk_paste(text, first_row_is_header=False):
    """
    Parse pasted text from Google Sheets / Excel / Notepad.
    - If any line contains a tab, lines are split on tabs (TSV).
    - Otherwise each line is parsed as CSV (handles commas inside quoted fields).
    - Without header: columns must be Name, Course, Month, Email (in that order).
    - With header: first row is column titles; uses same names as CSV import (Name, Email, etc.).

    Returns (entries, warnings) where entries is a list of dicts with keys
    name, course, month, email; warnings is a list of human-readable strings.
    """
    text = (text or "").strip()
    if not text:
        return [], ["Nothing to paste."]

    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return [], ["No non-empty lines."]

    use_tab = any("\t" in ln for ln in lines)

    def split_line(ln):
        if use_tab:
            return [c.strip() for c in ln.split("\t")]
        try:
            return next(csv.reader(io.StringIO(ln)))
        except Exception:
            return [ln.strip()]

    col_to_key = {}
    start_idx = 0
    if first_row_is_header:
        header_cells = split_line(lines[0])
        for i, cell in enumerate(header_cells):
            norm = _normalize_header(cell)
            key = _HEADER_TO_KEY.get(norm)
            if key in (
                "name",
                "course",
                "month",
                "date_of_completion",
                "scheduled_time",
                "email",
            ):
                col_to_key[i] = key
        mapped = set(col_to_key.values())
        if "name" not in mapped or "email" not in mapped:
            first_row_is_header = False
            col_to_key = {}
            start_idx = 0
        else:
            start_idx = 1

    entries_out = []
    warnings = []

    for row_idx, ln in enumerate(lines[start_idx:], start=1):
        cells = split_line(ln)
        if col_to_key:
            d = {k: "" for k in ("name", "course", "month", "email")}
            for i, key in col_to_key.items():
                if i < len(cells) and cells[i] is not None:
                    val = str(cells[i]).strip()
                    if key in d:
                        d[key] = val
            name = d["name"]
            course = d["course"]
            month = d["month"]
            email = d["email"]
        else:
            padded = list(cells) + [""] * max(0, 4 - len(cells))
            if len(padded) > 4:
                padded = padded[:4]
            name, course, month, email = (
                padded[0].strip() if padded[0] else "",
                padded[1].strip() if len(padded) > 1 and padded[1] else "",
                padded[2].strip() if len(padded) > 2 and padded[2] else "",
                padded[3].strip() if len(padded) > 3 and padded[3] else "",
            )

        if not name and not email:
            continue
        if not name or not email:
            warnings.append(f"Row {row_idx}: skipped — Name and Email are required.")
            continue
        if not course or not month:
            warnings.append(
                f"Row {row_idx}: skipped — Course and Month are required ({name!r})."
            )
            continue
        entries_out.append(
            {"name": name, "course": course, "month": month, "email": email}
        )

    return entries_out, warnings


def _image_path(filename):
    return os.path.join(STATIC_DIR, "images", filename)


def _font_css_url():
    path = os.path.join(STATIC_DIR, "fonts", "poppins.css")
    if os.path.isfile(path):
        return "file:///" + os.path.normpath(path).replace("\\", "/")
    return ""


def send_email(receiver_email, certificate_path, name, course, month):
    unsubscribe_link = os.getenv("UNSUBSCRIBE_LINK", "https://leveluponline.shop/")
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
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = f"Certificate of Achievement: {course}"
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


def generate_and_send_certificate(name, course, month, email, log_callback=None):
    """
    Generate PDF certificate and send email. Returns (success: bool, message: str).
    Template uses name, course, month only (same as before).
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    if not name or not email or not course or not month:
        return False, "Name, Email, Course, Month are required."

    try:
        config = get_pdfkit_config()
    except Exception as e:
        return False, f"wkhtmltopdf config failed: {e}"

    try:
        logo_path = _image_path("logo.png")
        certify_path = _image_path("certify.png")
        sign_path = _image_path("sign.png")
        with open(logo_path, "rb") as f:
            logo_data_url = f"data:image/png;base64,{base64.b64encode(f.read()).decode('utf-8')}"
        with open(certify_path, "rb") as f:
            certify_data_url = f"data:image/png;base64,{base64.b64encode(f.read()).decode('utf-8')}"
        with open(sign_path, "rb") as f:
            signature_data_url = f"data:image/png;base64,{base64.b64encode(f.read()).decode('utf-8')}"
    except FileNotFoundError as e:
        return False, f"Missing image: {e}"

    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(TEMPLATE_NAME)
    html = template.render(
        name=name,
        course=course,
        month=month,
        logo_url=logo_data_url,
        certify_url=certify_data_url,
        signature_url=signature_data_url,
        font_css_url=_font_css_url(),
    )

    safe_name = name.replace(" ", "_")
    output_path = os.path.join(OUTPUT_DIR, f"{safe_name}_{course}_{month}.pdf")
    options = {
        "enable-local-file-access": None,
        "no-stop-slow-scripts": "",
        "quiet": "",
        "margin-top": "0mm", "margin-bottom": "0mm", "margin-left": "0mm", "margin-right": "0mm",
        "page-width": "215mm", "page-height": "158mm", "dpi": "300",
    }

    try:
        pdfkit.from_string(html, output_path, configuration=config, options=options)
    except Exception as e:
        return False, f"PDF failed: {e}"

    try:
        send_email(email, output_path, name, course, month)
        log(f"[SENT] {name} ({email})")
        return True, "Sent"
    except Exception as e:
        log(f"[ERROR] {name}: {e}")
        return False, str(e)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    ENTRY_KEYS = ("name", "course", "month", "email")

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
        row = (
            name_var.get().strip(),
            course_var.get().strip(),
            month_var.get().strip(),
            email_var.get().strip(),
        )
        if not row[0] or not row[3]:
            messagebox.showwarning("Add entry", "Name and Email are required.")
            return
        tree.insert("", tk.END, values=row)
        for v in entries_vars:
            v.set("")

    def remove_selected():
        sel = tree.selection()
        for i in sel:
            tree.delete(i)

    def clear_list():
        if messagebox.askyesno("Clear list", "Remove all entries from the list?"):
            for i in tree.get_children():
                tree.delete(i)

    def get_entries():
        out = []
        for row in tree.get_children():
            vals = tree.item(row, "values")
            if len(vals) >= 4:
                out.append({
                    "name": vals[0],
                    "course": vals[1],
                    "month": vals[2],
                    "email": vals[3],
                })
        return out

    def save_list():
        data = get_entries()
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
            for i in tree.get_children():
                tree.delete(i)
            for e in data:
                row = (
                    e.get("name", ""),
                    e.get("course", ""),
                    e.get("month", ""),
                    e.get("email", ""),
                )
                tree.insert("", tk.END, values=row)
            messagebox.showinfo("Load", f"Loaded {len(data)} entries from {os.path.basename(path)}.")
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

        for e in entries:
            ok, msg = generate_and_send_certificate(
                e["name"], e["course"], e["month"], e["email"], log_callback=log
            )
            if ok:
                sent += 1
                record = {
                    "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "name": e["name"],
                    "email": e["email"],
                    "course": e["course"],
                    "month": e["month"],
                }
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
    root.title("Certificate Auto")
    root.minsize(700, 550)
    root.geometry("800x600")

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

    # ---------- Tab 1: Entries ----------
    tab_entries = ttk.Frame(notebook, padding=4)
    notebook.add(tab_entries, text="Entries")

    # Form frame
    form = ttk.LabelFrame(tab_entries, text="Add entry", padding=8)
    form.pack(fill=tk.X, padx=4, pady=4)

    name_var = tk.StringVar()
    course_var = tk.StringVar()
    month_var = tk.StringVar()
    email_var = tk.StringVar()
    entries_vars = [name_var, course_var, month_var, email_var]

    ttk.Label(form, text="Name *").grid(row=0, column=0, sticky=tk.W, padx=(0, 4))
    ttk.Entry(form, textvariable=name_var, width=22).grid(row=0, column=1, padx=(0, 12))
    ttk.Label(form, text="Course *").grid(row=0, column=2, sticky=tk.W, padx=(0, 4))
    ttk.Entry(form, textvariable=course_var, width=18).grid(row=0, column=3, padx=(0, 12))
    ttk.Label(form, text="Month *").grid(row=0, column=4, sticky=tk.W, padx=(0, 4))
    ttk.Entry(form, textvariable=month_var, width=14).grid(row=0, column=5, padx=(0, 12))
    ttk.Label(form, text="Email *").grid(row=1, column=0, sticky=tk.W, padx=(0, 4))
    ttk.Entry(form, textvariable=email_var, width=40).grid(row=1, column=1, columnspan=3, padx=(0, 12), sticky=tk.EW)
    ttk.Button(form, text="Add to list", command=add_entry).grid(row=1, column=4, padx=(0, 6))
    # ❗ open_paste_rows_dialog is defined after `tree` exists (see below)
    paste_btn = ttk.Button(form, text="Paste rows…")
    paste_btn.grid(row=1, column=5, pady=(0, 0))

    # List frame
    list_frame = ttk.LabelFrame(tab_entries, text="Entries to send", padding=8)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    cols = ("Name", "Course", "Month", "Email")
    tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=8, selectmode="extended")
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=100)
    tree.column("Name", width=120)
    tree.column("Email", width=180)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=scroll.set)

    def open_paste_rows_dialog():
        # ✅ Bulk-add: paste TSV/CSV from Sheets, Excel, or Notepad
        win = tk.Toplevel(root)
        win.title("Paste rows")
        win.minsize(520, 420)
        win.geometry("640x480")

        header_var = tk.BooleanVar(value=False)

        help_txt = (
            "Paste from Google Sheets or Excel (select cells → Copy): columns are TAB-separated.\n"
            "Or type one person per line. Without a header row, use this order:\n"
            "Name — Course — Month — Email\n"
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
                raw, first_row_is_header=header_var.get()
            )
            for e in entries_parsed:
                tree.insert(
                    "",
                    tk.END,
                    values=(e["name"], e["course"], e["month"], e["email"]),
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

    btn_frame = ttk.Frame(tab_entries, padding=8)
    btn_frame.pack(fill=tk.X)

    ttk.Button(btn_frame, text="Remove selected", command=remove_selected).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="Clear list", command=clear_list).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="Save list...", command=save_list).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="Load list...", command=load_list).pack(side=tk.LEFT, padx=(0, 6))
    ttk.Button(btn_frame, text="Generate & Send", command=generate_and_send).pack(side=tk.LEFT, padx=(0, 6))

    # Only show console toggle when running from Python (not from built exe); exe is built windowed so no console
    if sys.platform == "win32" and not getattr(sys, "frozen", False):
        ttk.Button(btn_frame, text="Hide console", command=lambda: toggle_console(False)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text="Show console", command=lambda: toggle_console(True)).pack(side=tk.LEFT, padx=(0, 6))

    status_var = tk.StringVar(value="Ready")
    ttk.Label(tab_entries, textvariable=status_var).pack(anchor=tk.W, padx=8, pady=2)

    log_frame = ttk.LabelFrame(tab_entries, text="Log", padding=4)
    log_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
    log_text = tk.Text(log_frame, height=6, wrap=tk.WORD, state=tk.NORMAL)
    log_text.pack(fill=tk.BOTH, expand=True)

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

        # Search filter (case-insensitive substring in name, email, course, month, sent_at)
        search_text = (history_search_var.get() or "").strip().lower()
        if search_text:
            def matches(r):
                for key in ("sent_at", "name", "email", "course", "month"):
                    if search_text in (str(r.get(key, "") or "").lower()):
                        return True
                return False
            records = [r for r in records if matches(r)]

        for i in history_tree.get_children():
            history_tree.delete(i)
        for r in reversed(records):
            history_tree.insert("", tk.END, values=(
                r.get("sent_at", ""),
                r.get("name", ""),
                r.get("email", ""),
                r.get("course", ""),
                r.get("month", ""),
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
        if search_text:
            def matches(r):
                for key in ("sent_at", "name", "email", "course", "month"):
                    if search_text in (str(r.get(key, "") or "").lower()):
                        return True
                return False
            records = [r for r in records if matches(r)]
        headers = ["Sent at", "Name", "Email", "Course", "Month"]

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
        rows_for_csv = [
            {
                "Sent at": _format_sent_at_for_csv(r.get("sent_at", "")),
                "Name": r.get("name", ""),
                "Email": r.get("email", ""),
                "Course": r.get("course", ""),
                "Month": r.get("month", ""),
            }
            for r in records
        ]
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

    history_cols = ("Sent at", "Name", "Email", "Course", "Month")
    history_tree_frame = ttk.Frame(tab_history)
    history_tree_frame.pack(fill=tk.BOTH, expand=True, pady=4)
    history_tree = ttk.Treeview(history_tree_frame, columns=history_cols, show="headings", height=12, selectmode="extended")
    for c in history_cols:
        history_tree.heading(c, text=c)
        history_tree.column(c, width=100)
    history_tree.column("Sent at", width=160)
    history_tree.column("Name", width=120)
    history_tree.column("Email", width=180)
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
