# Certificate Auto – Step-by-step desktop setup

This guide is for **building** the Certificate Auto desktop app. The app is a **GUI** where you add entries (Name, Course, Month, Date of Completion, Scheduled Time, Email), then click **Generate & Send** to create and email certificates. No Google Sheet; no scheduler.

---

## Part A: One-time setup (prerequisites)

### Step A1: Check Python

**Where:** Any terminal (PowerShell or Command Prompt).

**What to do:**

1. Open PowerShell: press `Win + X`, choose **Windows PowerShell** or **Terminal**.
2. Run:
   ```powershell
   python --version
   ```
3. You should see something like `Python 3.10.x` or higher.  
   If you see "not recognized", install Python from https://www.python.org/downloads/ and check **"Add Python to PATH"**.

---

### Step A2: Install wkhtmltopdf (for PDF generation)

**Where:** Your PC (system-wide install).

**What to do:**

1. Go to: **https://wkhtmltopdf.org/downloads.html**
2. Download the **Windows 64-bit** installer.
3. Run the installer and use the **default path**:  
   `C:\Program Files\wkhtmltopdf\`  
   The app expects:  
   `C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe`

**Why:** The app uses wkhtmltopdf to turn the certificate HTML into PDF.

---

### Step A3: Have your `.env` file ready

**Where:** Same folder as the built app (or project root when running from source).

**What to do:**

1. Create or copy a **`.env`** file with **SMTP** settings only (no Google credentials needed):
   - `SMTP_SERVER` (e.g. `smtp.gmail.com`)
   - `SMTP_PORT` (e.g. `587`)
   - `SMTP_USER` (your email)
   - `SMTP_PASSWORD` (App Password if using Gmail)
   - `SENDER_EMAIL` (same as SMTP_USER or your from-address)
   - Optional: `OUTPUT_DIR`, `UNSUBSCRIBE_LINK`, `TEMPLATE_NAME`
2. When building, you’ll copy `.env` next to the **built** app in Part C.

**Why:** The app reads SMTP and optional config from `.env`. No `.env` = no email.

---

## Part B: Build the desktop executable

Do this **inside your project folder** (where `app.py` and `certificate_auto.spec` are).

**Project folder (example):**  
`C:\Gsn Soln\certificate_auto_desktop`

---

### Step B1: Open terminal in the project folder

**Where:** Project root (contains `app.py`, `certificate_auto.spec`, `requirements.txt`).

**What to do:**

1. In File Explorer go to the project folder.
2. In the address bar type `powershell` and press Enter, or right‑click → **Open in Terminal**.

**Check:** Run `dir` and you should see `app.py`, `certificate_auto.spec`, `requirements.txt`, `templates`, `static`.

---

### Step B2: Create and activate a virtual environment

**Where:** Same PowerShell, project folder.

**What to do:**

1. Create venv:
   ```powershell
   python -m venv venv
   ```
2. Activate:
   ```powershell
   .\venv\Scripts\Activate.ps1
   ```
   If you get an execution policy error, run once:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
   Then run the Activate command again.

3. Your prompt should start with `(venv)`.

---

### Step B3: Install dependencies

**Where:** Same window (venv active, project folder).

**What to do:**

1. Install app packages:
   ```powershell
   pip install -r requirements.txt
   ```
2. Install PyInstaller:
   ```powershell
   pip install -r requirements-build.txt
   ```
   Or: `pip install pyinstaller`

---

### Step B4: Run the build

**Where:** Same window (venv active, project folder).

**What to do:**

1. Run:
   ```powershell
   pyinstaller certificate_auto.spec --noconfirm
   ```
2. Wait 1–3 minutes.

3. When done, output is in:
   - **Folder:** `dist\CertificateAuto\`
   - **Contents:** `CertificateAuto.exe`, DLLs, `templates`, `static`

**If build fails:** Note the exact error; we can add missing modules to the spec’s `hiddenimports`.

---

## Part C: Run the desktop app

Run the app from the **built folder**, not from source.

---

### Step C1: Copy `.env` next to the exe

**Where:**  
**From:** your `.env` file.  
**To:** inside `dist\CertificateAuto\`.

**What to do:**

1. Open: `C:\Gsn Soln\certificate_auto_desktop\dist\CertificateAuto\`
2. Copy **`.env`** into this folder so it sits **next to `CertificateAuto.exe`**.

---

### Step C2: Start the app

**Where:** `dist\CertificateAuto\`.

**What to do:**

1. Double‑click **`CertificateAuto.exe`**  
   or run in PowerShell:
   ```powershell
   .\CertificateAuto.exe
   ```
2. A **console window** and a **GUI window** open. Both are normal.
3. Use the **Hide console** / **Show console** buttons in the GUI to show or hide the console.

**What the app does:**

- **Add entry:** Fill Name, Course, Month, Date of Completion (optional), Scheduled Time (optional), Email, then click **Add to list**.
- **List:** All entries appear in the table. Use **Remove selected** or **Clear list** as needed.
- **Save list** / **Load list:** Save or load the list as a JSON file (e.g. `certificate_entries.json`).
- **Generate & Send:** For each entry, generates the certificate PDF and sends it by email. Log appears in the app and in the console.

---

### Step C3: Where certificates are saved

**Where:**  
By default: `dist\CertificateAuto\certificates\`  
(or the path in `OUTPUT_DIR` in `.env` if set).  
The app creates the folder automatically.

---

## Part D: Run without building (development)

**Where:** Project folder.

**What to do:**

1. Open PowerShell in the project folder.
2. Activate venv: `.\venv\Scripts\Activate.ps1`
3. Put **`.env`** in the **project root** (same folder as `app.py`).
4. Run:
   ```powershell
   python app.py
   ```
5. The same GUI opens; certificates go to `certificates\` (or `OUTPUT_DIR`).

---

## Quick reference

| What                    | Where / Command |
|-------------------------|------------------|
| Project folder          | `C:\Gsn Soln\certificate_auto_desktop` |
| Build command           | `pyinstaller certificate_auto.spec --noconfirm` (venv active) |
| Built app folder        | `dist\CertificateAuto\` |
| Put `.env` for exe      | Inside `dist\CertificateAuto\` (next to exe) |
| Run the app             | Double‑click `CertificateAuto.exe` or `.\CertificateAuto.exe` |
| Certificate PDFs        | `dist\CertificateAuto\certificates\` (or `OUTPUT_DIR`) |
| wkhtmltopdf (default)   | `C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe` |

---

## Troubleshooting

- **"wkhtmltopdf not found"**  
  Install wkhtmltopdf (Step A2) to the default path.

- **".env not loading" / no email**  
  Put `.env` in the **same folder as `CertificateAuto.exe`**.

- **Antivirus blocks the exe**  
  Add an exception for the `dist\CertificateAuto` folder.

- **Build fails with "ModuleNotFoundError"**  
  Add the missing module to the PyInstaller spec’s `hiddenimports`.

- **Moving the app**  
  Copy the **entire** `dist\CertificateAuto` folder (exe + all files + `.env`). Keep everything together; put `.env` next to the exe.
