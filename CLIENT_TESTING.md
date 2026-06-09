# Certificate Auto — share with your client (testing)

## What you send

Zip the **entire folder**:

`dist\CertificateAuto\`

It must include:

- `CertificateAuto.exe`
- `_internal\` (folder — do not delete)
- `.env` (you create this for the client with their SMTP settings, or send `.env.example` and let them rename/fill in)

**Do not** send your project source code or Python venv unless you want them to develop.

---

## What the client must install once (Windows)

1. **wkhtmltopdf** (creates PDFs)  
   - Download: https://wkhtmltopdf.org/downloads.html  
   - Install to default path: `C:\Program Files\wkhtmltopdf\`

2. **`.env` file** next to `CertificateAuto.exe`  
   - Copy from `.env.example`, add Gmail (or other) SMTP + App Password.

---

## How the client runs the app

1. Unzip the folder anywhere (e.g. `Desktop\CertificateAuto\`).
2. Put `.env` in the same folder as `CertificateAuto.exe`.
3. Double-click **`CertificateAuto.exe`**.
4. If Windows SmartScreen appears: **More info → Run anyway** (unsigned app).

---

## Quick test checklist for client

- [ ] App window opens (“Certificate Auto”)
- [ ] Type 1: add one row → Generate & Send → PDF in `certificates\` folder
- [ ] Type 2: select certificate title (Post Grad / Diploma) → add row → Generate & Send
- [ ] Email received (if SMTP is correct)
- [ ] History tab shows sent entries

---

## Rebuild after you change code (for you, the developer)

In PowerShell, from the project folder:

```powershell
.\then\Scripts\Activate.ps1
pip install -r requirements.txt
pip install -r requirements-build.txt
pyinstaller certificate_auto.spec --noconfirm
```

Or run: `.\build.ps1`

New build output: `dist\CertificateAuto\`

Copy your `.env` into that folder again after each rebuild (PyInstaller does not copy `.env` automatically).
