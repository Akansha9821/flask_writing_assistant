# Smart Mail & Letter Writer

A Flask and Bootstrap application for creating formal emails, informal emails, letters, complaints, follow-ups, study-related requests, bank requests, leave applications, and customer-service messages.

## Important hosting note

GitHub Pages serves static files only and cannot run Flask. Keep the source code in GitHub and deploy the Flask application to Render, Railway, PythonAnywhere, or another Python server.

## Features

- User registration and login
- Session-based active-user identity
- CSV and Excel storage
- Formal email, informal email, and letter categories
- Automatic greeting, subject improvement, grammar cleanup, closing, and digital signature
- Study, banking, office, leave, customer-service, device complaint, payment, exam, project, placement, practical, viva, and class-related templates
- Attachment upload
- Print view
- PDF, DOCX, PNG, and TXT downloads
- Writing history

## Local setup

```bash
python -m venv venv
source venv/bin/activate
# Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`.

## Environment variables

Set a strong secret key in production:

```bash
export SECRET_KEY="replace-with-a-long-random-secret"
```

## Storage warning

CSV and Excel are acceptable for a college prototype or single-user demonstration. For production, replace them with SQLite, PostgreSQL, or MySQL because spreadsheet files are not safe for concurrent writes and should not store sensitive data in a public GitHub repository.

## Added capabilities

- Draw a digital signature on a signature canvas
- Upload a handwritten signature image
- Automatic browser-region language selection
- Hindi and multilingual translation
- Attachment information displayed after the writing body
- RPPS Software professional footer and copyright branding

Translation uses `deep-translator` and therefore requires internet access while generating a translated draft. Browser language and regional settings are used as a privacy-friendly default; the user can always change the selected language.


## Updated Version 3

### Two dashboards

- User dashboard: personal writing, download and print totals.
- Administrator dashboard: total users, writing count, download count and print count, with activity per user.
- Admin exports: CSV and Excel.

Default development administrator:

```text
Email: admin@rppssoftware.com
Password: Admin@123
```

Set `ADMIN_EMAIL` and `ADMIN_PASSWORD` environment variables before the first production start. Never keep the default password in production.

### Concurrent users

SQLite is now the primary live datastore. WAL mode, busy timeout, short transactions and threaded Flask development mode are enabled. CSV and Excel are exports only.

For a high-traffic production deployment or multiple application servers, use PostgreSQL rather than a shared SQLite file.

### Handwriting OCR

Python OCR requires both the Python package and the system Tesseract executable.

macOS:

```bash
brew install tesseract
brew install tesseract-lang
```

Ubuntu/Debian:

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-hin
```

### Voice input

Voice dictation uses the browser Web Speech API. It requires microphone permission and is not supported equally by every browser.

### Persistent login

The session is refreshed on each request and configured for up to ten years unless the user logs out. Production systems should normally use a shorter duration and secure HTTPS cookies.

### Run Command

 git push --set-upstream fwa main
