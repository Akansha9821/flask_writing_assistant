import base64
import io
import os
import re
import sqlite3
import uuid
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path

import pandas as pd
from docx import Document
from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image as PDFImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    from deep_translator import GoogleTranslator
except ImportError:
    GoogleTranslator = None

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"
DB_PATH = DATA_DIR / "rpps_writer.db"

ALLOWED_EXTENSIONS = {"pdf", "png", "jpg", "jpeg", "doc", "docx", "txt", "csv", "xlsx"}
IMAGE_EXTENSIONS = {"png", "jpg", "jpeg"}

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "replace-this-development-secret")
app.config.update(
    MAX_CONTENT_LENGTH=10 * 1024 * 1024,
    UPLOAD_FOLDER=str(UPLOAD_DIR),
    PERMANENT_SESSION_LIFETIME=timedelta(days=3650),
    SESSION_REFRESH_EACH_REQUEST=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true",
)

DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

WRITING_TYPES = {
    "formal_email": "Formal Email",
    "informal_email": "Informal Email",
    "formal_letter": "Formal Letter",
    "application": "Application",
    "mom": "Minutes of Meeting (MOM)",
}

CATEGORY_TREE = {
    "Office & Employment": {
        "leave_request": "Leave Request",
        "office_followup": "Office Follow-up",
        "peer_communication": "Communication with Peers",
        "manager_request": "Manager Request",
        "hr_request": "HR Request",
        "work_from_home": "Work From Home Request",
        "resignation": "Resignation Application",
        "experience_letter": "Experience Letter Request",
    },
    "Banking & Finance": {
        "bank_service": "Bank Customer Service",
        "payment_issue": "Payment Issue",
        "refund_request": "Refund Request",
        "transaction_dispute": "Transaction Dispute",
        "account_update": "Account Update Request",
        "loan_request": "Loan-related Request",
    },
    "Customer & Product Service": {
        "device_complaint": "Device Complaint",
        "customer_service": "Customer Service Request",
        "product_complaint": "Product Complaint",
        "replacement_request": "Replacement Request",
        "warranty_request": "Warranty Request",
        "service_followup": "Service Follow-up",
    },
    "Study & University": {
        "exam": "Exam Related",
        "project": "Project Related",
        "placement": "Placement Related",
        "practical": "Practical Related",
        "viva": "Viva Related",
        "fees_payment": "Fees or Payment Related",
        "class": "Class Related",
        "attendance": "Attendance Related",
        "academic_complaint": "Academic Complaint",
        "certificate": "Certificate Request",
        "scholarship": "Scholarship Application",
    },
    "Meeting & Administration": {
        "mom_internal": "Internal Meeting MOM",
        "mom_client": "Client Meeting MOM",
        "mom_project": "Project Meeting MOM",
        "mom_review": "Review Meeting MOM",
        "general_application": "General Application",
        "permission_application": "Permission Application",
        "complaint_application": "Complaint Application",
    },
    "General": {
        "help": "General Help",
        "complaint": "General Complaint",
        "request": "General Request",
        "invitation": "Invitation",
        "thank_you": "Thank-you Message",
    },
}

LANGUAGES = {
    "en": "English",
    "hi": "Hindi",
    "bn": "Bengali",
    "mr": "Marathi",
    "gu": "Gujarati",
    "pa": "Punjabi",
    "ta": "Tamil",
    "te": "Telugu",
    "kn": "Kannada",
    "ml": "Malayalam",
    "ur": "Urdu",
    "ne": "Nepali",
    "fr": "French",
    "de": "German",
    "es": "Spanish",
    "it": "Italian",
    "pt": "Portuguese",
    "ar": "Arabic",
    "ja": "Japanese",
    "ko": "Korean",
    "zh-CN": "Chinese (Simplified)",
    "ru": "Russian",
}

INFO_PAGES = {
    "about": {
        "eyebrow": "Who we are",
        "title": "About RPPS",
        "icon": "building",
        "intro": "RPPS Writing Assistant helps people turn everyday ideas into clear, polished professional communication.",
        "sections": [
            {
                "heading": "Our purpose",
                "body": "We make formal writing easier for office, academic, banking, customer-service and personal communication needs.",
            },
            {
                "heading": "How it works",
                "body": "Choose a writing type, provide the important details and review the structured draft before printing or exporting it.",
            },
            {
                "heading": "Built for practical work",
                "body": "The workspace keeps recent writings together and supports PDF, document and image exports for convenient sharing.",
            },
            {
                "heading": "Responsible assistance",
                "body": "Every generated draft should be reviewed by the user to confirm names, facts, tone and context before it is sent.",
            },
        ],
    },
    "services": {
        "eyebrow": "What we offer",
        "title": "Services",
        "icon": "grid-1x2",
        "intro": "A focused set of tools for drafting, refining, signing and exporting professional writing.",
        "sections": [
            {
                "heading": "Email and letter drafting",
                "body": "Create formal emails, informal emails, applications, letters and meeting minutes from guided inputs.",
            },
            {
                "heading": "Writing organization",
                "body": "Access recent work and history from your dashboard so drafts remain easy to find and reuse.",
            },
            {
                "heading": "Signatures and attachments",
                "body": "Add a drawn or uploaded signature and keep a supporting attachment associated with your writing.",
            },
            {
                "heading": "Flexible exports",
                "body": "Print a clean layout or export supported writings as PDF, DOCX, PNG or text files.",
            },
        ],
    },
    "contact": {
        "eyebrow": "Get in touch",
        "title": "Contact",
        "icon": "envelope",
        "intro": "Questions, feedback or technical concerns are welcome. Use the contact option below to reach the RPPS team.",
        "sections": [
            {
                "heading": "General enquiries",
                "body": "Contact us with questions about the writing assistant, its features or how to get started.",
                "link": {"href": "mailto:admin@rppssoftware.com", "label": "Email RPPS"},
            },
            {
                "heading": "Technical support",
                "body": "Include a short description of the issue, the page where it occurred and any error message you received.",
                "link": {"href": "/help", "label": "Visit Help and Support"},
            },
            {
                "heading": "Product feedback",
                "body": "Tell us which workflows are useful and what would make professional writing easier for you.",
                "link": {"href": "mailto:admin@rppssoftware.com?subject=Writing%20Assistant%20Feedback", "label": "Send feedback"},
            },
            {
                "heading": "Response information",
                "body": "Please avoid including passwords or other sensitive information in support messages.",
            },
        ],
    },
    "privacy": {
        "eyebrow": "Your information",
        "title": "Privacy Policy",
        "icon": "shield-check",
        "intro": "This page explains the categories of information used by the writing assistant and the choices available to users.",
        "sections": [
            {
                "heading": "Information you provide",
                "body": "The service stores account details and the writing, attachments and signatures you choose to submit so it can provide its features.",
            },
            {
                "heading": "How information is used",
                "body": "Information is used to authenticate your account, create and manage writings, produce exports and maintain service operation.",
            },
            {
                "heading": "Account security",
                "body": "Use a strong password, protect access to your device and sign out when using a shared computer.",
            },
            {
                "heading": "Privacy questions",
                "body": "Contact RPPS if you have a question about information associated with your account.",
                "link": {"href": "mailto:admin@rppssoftware.com?subject=Privacy%20Question", "label": "Ask a privacy question"},
            },
        ],
    },
    "terms": {
        "eyebrow": "Using the service",
        "title": "Terms and Conditions",
        "icon": "file-earmark-text",
        "intro": "These practical terms describe responsible use of RPPS Writing Assistant and its generated drafts.",
        "sections": [
            {
                "heading": "Review your drafts",
                "body": "Generated content is a drafting aid. You are responsible for reviewing accuracy, suitability and recipients before use.",
            },
            {
                "heading": "Acceptable use",
                "body": "Do not use the service to create unlawful, deceptive, abusive or rights-infringing material.",
            },
            {
                "heading": "Your account",
                "body": "Keep account credentials confidential and provide accurate registration information when using the service.",
            },
            {
                "heading": "Service availability",
                "body": "Features may be updated or temporarily unavailable for maintenance, security or operational reasons.",
            },
        ],
    },
    "help": {
        "eyebrow": "Guidance and support",
        "title": "Help and Support",
        "icon": "life-preserver",
        "intro": "Quick guidance for creating, reviewing, exporting and managing your writing.",
        "sections": [
            {
                "heading": "Create a writing",
                "body": "Open Create Writing, select the document type and category, add the essential details and submit the form to generate a draft.",
            },
            {
                "heading": "Review before exporting",
                "body": "Check the recipient, subject, dates, names and tone. Then use Print, PDF, DOCX or Image from the preview page.",
            },
            {
                "heading": "Find previous work",
                "body": "Use Dashboard for recent activity or History to browse the writings saved to your account.",
            },
            {
                "heading": "Still need help?",
                "body": "Send a concise description of the problem and include any displayed error message.",
                "link": {"href": "mailto:admin@rppssoftware.com?subject=Writing%20Assistant%20Support", "label": "Contact support"},
            },
        ],
    },
    "accessibility": {
        "eyebrow": "Inclusive experience",
        "title": "Accessibility",
        "icon": "universal-access",
        "intro": "RPPS Writing Assistant aims to provide a clear experience that works across common devices and input methods.",
        "sections": [
            {
                "heading": "Keyboard access",
                "body": "Interactive controls are designed with standard links, buttons and form fields that can be reached using a keyboard.",
            },
            {
                "heading": "Readable presentation",
                "body": "Layouts use clear headings, responsive spacing and strong text contrast to support comfortable reading.",
            },
            {
                "heading": "Responsive layouts",
                "body": "Pages adapt to desktop, tablet and mobile screens while keeping key actions available.",
            },
            {
                "heading": "Report a barrier",
                "body": "If something prevents you from using the service, tell us which page and assistive technology were involved.",
                "link": {"href": "mailto:admin@rppssoftware.com?subject=Accessibility%20Feedback", "label": "Send accessibility feedback"},
            },
        ],
    },
}


def db_connection():
    connection = sqlite3.connect(DB_PATH, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA busy_timeout = 30000")
    connection.execute("PRAGMA journal_mode = WAL")
    connection.execute("PRAGMA synchronous = NORMAL")
    return connection


def init_db():
    with db_connection() as db:
        db.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            designation TEXT DEFAULT '',
            organization TEXT DEFAULT '',
            role TEXT NOT NULL DEFAULT 'user',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS writings (
            id TEXT PRIMARY KEY,
            user_id TEXT NOT NULL,
            writing_type TEXT NOT NULL,
            category_group TEXT NOT NULL,
            category TEXT NOT NULL,
            recipient_name TEXT DEFAULT '',
            recipient_email TEXT DEFAULT '',
            subject TEXT NOT NULL,
            body TEXT NOT NULL,
            attachment TEXT DEFAULT '',
            attachment_original_name TEXT DEFAULT '',
            signature_type TEXT DEFAULT 'none',
            signature_file TEXT DEFAULT '',
            language TEXT DEFAULT 'en',
            download_count INTEGER NOT NULL DEFAULT 0,
            print_count INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_writings_user ON writings(user_id);
        CREATE INDEX IF NOT EXISTS idx_writings_created ON writings(created_at);
        """)
    create_default_admin()


def create_default_admin():
    email = os.getenv("ADMIN_EMAIL", "admin@rppssoftware.com").strip().lower()
    password = os.getenv("ADMIN_PASSWORD", "Admin@123")
    with db_connection() as db:
        exists = db.execute("SELECT id FROM users WHERE email = ?", (email,)).fetchone()
        if not exists:
            db.execute(
                """INSERT INTO users
                   (id, name, email, password_hash, designation, organization, role, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, 'admin', ?)""",
                (
                    str(uuid.uuid4()),
                    "RPPS Administrator",
                    email,
                    generate_password_hash(password),
                    "Administrator",
                    "RPPS Software",
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
            db.commit()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Administrator access is required.", "danger")
            return redirect(url_for("dashboard"))
        return view(*args, **kwargs)

    return wrapped


def allowed_file(filename, image_only=False):
    if "." not in filename:
        return False
    extension = filename.rsplit(".", 1)[1].lower()
    return extension in (IMAGE_EXTENSIONS if image_only else ALLOWED_EXTENSIONS)


def clean_sentence(text):
    text = re.sub(r"\s+", " ", (text or "").strip())
    replacements = {
        r"\bi\b": "I",
        r"\bpls\b": "please",
        r"\bplz\b": "please",
        r"\bu\b": "you",
        r"\bur\b": "your",
        r"\bcoz\b": "because",
        r"\bcant\b": "cannot",
        r"\bwont\b": "will not",
        r"\bdidnt\b": "did not",
        r"\bdoesnt\b": "does not",
        r"\bim\b": "I am",
    }
    for pattern, value in replacements.items():
        text = re.sub(pattern, value, text, flags=re.IGNORECASE)
    if text and text[-1] not in ".!?":
        text += "."
    return text[:1].upper() + text[1:] if text else ""


def improve_subject(subject, category_label):
    subject = clean_sentence(subject).rstrip(".")
    return (
        subject
        if len(subject.split()) >= 3
        else (f"{category_label}: {subject}" if subject else category_label)
    )


def build_standard_body(writing_type, category_label, recipient_name, details, user):
    details = clean_sentence(details)
    recipient_name = recipient_name.strip() or "Sir/Madam"
    informal = writing_type == "informal_email"
    greeting = f"Hi {recipient_name}," if informal else f"Dear {recipient_name},"
    closing = (
        "Best regards,"
        if informal
        else (
            "Yours sincerely,"
            if writing_type in {"formal_letter", "application"}
            else "Kind regards,"
        )
    )

    opener = f"I am writing regarding {category_label.lower()}."
    request_line = "I kindly request you to review the matter and provide the necessary assistance or resolution."

    signature = [user["name"]]
    if user["designation"]:
        signature.append(user["designation"])
    if user["organization"]:
        signature.append(user["organization"])
    signature.append(user["email"])

    return "\n\n".join(
        [
            greeting,
            opener,
            details,
            request_line,
            "Thank you for your time and consideration.",
            closing,
            "\n".join(signature),
        ]
    )


def build_mom_body(subject, details, user):
    today = datetime.now().strftime("%d %B %Y")
    return "\n\n".join(
        [
            "MINUTES OF MEETING",
            f"Meeting Title: {subject}",
            f"Date: {today}",
            f"Prepared by: {user['name']}",
            "Attendees:\n[Add attendee names and roles]",
            "Agenda:\n[Add the meeting agenda]",
            f"Discussion Summary:\n{clean_sentence(details)}",
            "Decisions Taken:\n1. [Add decision]\n2. [Add decision]",
            "Action Items:\n1. [Task] — [Owner] — [Due date]\n2. [Task] — [Owner] — [Due date]",
            "Next Meeting:\n[Add date, time and purpose]",
            f"Prepared and recorded by:\n{user['name']}\n{user['organization'] or 'RPPS Writing Assistant'}",
        ]
    )


def translate_text(text, target_language):
    if not text or target_language == "en" or GoogleTranslator is None:
        return text
    try:
        output = []
        for part in text.split("\n\n"):
            output.append(
                GoogleTranslator(source="auto", target=target_language).translate(part)
                if part.strip()
                else ""
            )
        return "\n\n".join(output)
    except Exception:
        return text


def save_signature(data_url, uploaded):
    if uploaded and uploaded.filename:
        if not allowed_file(uploaded.filename, image_only=True):
            raise ValueError("Signature must be a PNG, JPG or JPEG image.")
        original = secure_filename(uploaded.filename)
        filename = f"signature_{uuid.uuid4().hex}_{original}"
        uploaded.save(UPLOAD_DIR / filename)
        return "uploaded", filename
    if data_url and data_url.startswith("data:image"):
        _, encoded = data_url.split(",", 1)
        filename = f"signature_{uuid.uuid4().hex}.png"
        (UPLOAD_DIR / filename).write_bytes(base64.b64decode(encoded))
        return "drawn", filename
    return "none", ""


def current_user():
    with db_connection() as db:
        return db.execute(
            "SELECT * FROM users WHERE id = ?", (session["user_id"],)
        ).fetchone()


def writing_by_id(writing_id):
    with db_connection() as db:
        query = "SELECT * FROM writings WHERE id = ?"
        params = [writing_id]
        if session.get("role") != "admin":
            query += " AND user_id = ?"
            params.append(session["user_id"])
        return db.execute(query, params).fetchone()


init_db()


@app.route("/")
def index():
    return render_template("index.html")


def render_info_page(page_name):
    return render_template("info_page.html", page=INFO_PAGES[page_name])


@app.route("/about")
def about():
    return render_info_page("about")


@app.route("/services")
def services():
    return render_info_page("services")


@app.route("/contact")
def contact():
    return render_info_page("contact")


@app.route("/privacy")
def privacy():
    return render_info_page("privacy")


@app.route("/terms")
def terms():
    return render_info_page("terms")


@app.route("/help")
def help_support():
    return render_info_page("help")


@app.route("/accessibility")
def accessibility():
    return render_info_page("accessibility")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        if not name or not email or len(password) < 6:
            flash(
                "Name, email and a password of at least six characters are required.",
                "danger",
            )
            return render_template("register.html")
        try:
            with db_connection() as db:
                db.execute(
                    """INSERT INTO users
                       (id, name, email, password_hash, designation, organization, role, created_at)
                       VALUES (?, ?, ?, ?, ?, ?, 'user', ?)""",
                    (
                        str(uuid.uuid4()),
                        name,
                        email,
                        generate_password_hash(password),
                        request.form.get("designation", "").strip(),
                        request.form.get("organization", "").strip(),
                        datetime.now().isoformat(timespec="seconds"),
                    ),
                )
                db.commit()
            flash("Registration completed. Please log in.", "success")
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            flash("An account with this email already exists.", "warning")
    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        with db_connection() as db:
            user = db.execute(
                "SELECT * FROM users WHERE email = ?", (email,)
            ).fetchone()
        if user and check_password_hash(
            user["password_hash"], request.form.get("password", "")
        ):
            session.clear()
            session.permanent = True
            session["user_id"] = user["id"]
            session["user_name"] = user["name"]
            session["role"] = user["role"]
            return redirect(
                url_for("admin_dashboard" if user["role"] == "admin" else "dashboard")
            )
        flash("Invalid email or password.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("index"))


@app.route("/dashboard")
@login_required
def dashboard():
    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))
    with db_connection() as db:
        stats = db.execute(
            """SELECT COUNT(*) writing_count,
                      COALESCE(SUM(download_count), 0) download_count,
                      COALESCE(SUM(print_count), 0) print_count
               FROM writings WHERE user_id = ?""",
            (session["user_id"],),
        ).fetchone()
        records = db.execute(
            "SELECT * FROM writings WHERE user_id = ? ORDER BY created_at DESC LIMIT 8",
            (session["user_id"],),
        ).fetchall()
    return render_template("dashboard.html", records=records, stats=stats)


@app.route("/admin")
@login_required
@admin_required
def admin_dashboard():
    with db_connection() as db:
        totals = db.execute(
            """SELECT
                (SELECT COUNT(*) FROM users WHERE role='user') total_users,
                (SELECT COUNT(*) FROM writings) total_writings,
                (SELECT COALESCE(SUM(download_count),0) FROM writings) total_downloads,
                (SELECT COALESCE(SUM(print_count),0) FROM writings) total_prints"""
        ).fetchone()
        users = db.execute(
            """SELECT u.id, u.name, u.email, u.designation, u.organization, u.created_at,
                      COUNT(w.id) writing_count,
                      COALESCE(SUM(w.download_count),0) download_count,
                      COALESCE(SUM(w.print_count),0) print_count
               FROM users u
               LEFT JOIN writings w ON w.user_id=u.id
               WHERE u.role='user'
               GROUP BY u.id
               ORDER BY u.created_at DESC"""
        ).fetchall()
    return render_template("admin_dashboard.html", totals=totals, users=users)


@app.route("/admin/export/<file_type>")
@login_required
@admin_required
def admin_export(file_type):
    with db_connection() as db:
        users = pd.read_sql_query(
            "SELECT id,name,email,designation,organization,role,created_at FROM users",
            db,
        )
        writings = pd.read_sql_query("SELECT * FROM writings", db)

    output = io.BytesIO()
    if file_type == "xlsx":
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            users.to_excel(writer, sheet_name="Users", index=False)
            writings.to_excel(writer, sheet_name="Writings", index=False)
        output.seek(0)
        return send_file(
            output, as_attachment=True, download_name="rpps_admin_report.xlsx"
        )
    if file_type == "csv":
        output.write(users.to_csv(index=False).encode("utf-8"))
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name="rpps_users.csv",
            mimetype="text/csv",
        )
    flash("Unsupported export type.", "warning")
    return redirect(url_for("admin_dashboard"))


@app.route("/compose", methods=["GET", "POST"])
@login_required
def compose():
    if request.method == "POST":
        writing_type = request.form.get("writing_type", "formal_email")
        group = request.form.get("category_group", "General")
        category = request.form.get("category", "help")
        category_label = next(
            (items[category] for items in CATEGORY_TREE.values() if category in items),
            "General Request",
        )
        recipient_name = request.form.get("recipient_name", "").strip()
        recipient_email = request.form.get("recipient_email", "").strip()
        language = request.form.get("language", "en")
        subject = improve_subject(request.form.get("subject", ""), category_label)
        details = request.form.get("details", "")

        attachment = attachment_original = ""
        uploaded = request.files.get("attachment")
        if uploaded and uploaded.filename:
            if not allowed_file(uploaded.filename):
                flash("Unsupported attachment type.", "danger")
                return redirect(url_for("compose"))
            attachment_original = secure_filename(uploaded.filename)
            attachment = f"{uuid.uuid4().hex}_{attachment_original}"
            uploaded.save(UPLOAD_DIR / attachment)

        try:
            signature_type, signature_file = save_signature(
                request.form.get("signature_data", ""),
                request.files.get("signature_upload"),
            )
        except ValueError as exc:
            flash(str(exc), "danger")
            return redirect(url_for("compose"))

        user = current_user()
        body = (
            build_mom_body(subject, details, user)
            if writing_type == "mom"
            else build_standard_body(
                writing_type, category_label, recipient_name, details, user
            )
        )
        subject = translate_text(subject, language)
        body = translate_text(body, language)
        writing_id = str(uuid.uuid4())

        with db_connection() as db:
            db.execute(
                """INSERT INTO writings
                   (id,user_id,writing_type,category_group,category,recipient_name,
                    recipient_email,subject,body,attachment,attachment_original_name,
                    signature_type,signature_file,language,created_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    writing_id,
                    session["user_id"],
                    writing_type,
                    group,
                    category,
                    recipient_name,
                    recipient_email,
                    subject,
                    body,
                    attachment,
                    attachment_original,
                    signature_type,
                    signature_file,
                    language,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
            db.commit()
        return redirect(url_for("preview", writing_id=writing_id))

    return render_template(
        "compose.html",
        writing_types=WRITING_TYPES,
        category_tree=CATEGORY_TREE,
        languages=LANGUAGES,
    )


@app.route("/ocr", methods=["POST"])
@login_required
def ocr_image():
    image_file = request.files.get("handwriting_image")
    if not image_file or not image_file.filename:
        return jsonify({"error": "Select a handwritten image first."}), 400
    if not allowed_file(image_file.filename, image_only=True):
        return jsonify({"error": "Only PNG, JPG and JPEG images are supported."}), 400
    if pytesseract is None:
        return jsonify({"error": "pytesseract is not installed on the server."}), 503
    try:
        image = Image.open(image_file.stream).convert("RGB")
        language = request.form.get("ocr_language", "eng")
        text = pytesseract.image_to_string(image, lang=language).strip()
        return jsonify({"text": text})
    except Exception as exc:
        return (
            jsonify(
                {
                    "error": "OCR failed. Confirm that the Tesseract OCR system package and selected language pack are installed."
                }
            ),
            500,
        )


@app.route("/preview/<writing_id>")
@login_required
def preview(writing_id):
    writing = writing_by_id(writing_id)
    if not writing:
        flash("Writing not found.", "danger")
        return redirect(url_for("dashboard"))
    return render_template("preview.html", writing=writing)


@app.route("/history")
@login_required
def history():
    with db_connection() as db:
        records = db.execute(
            "SELECT * FROM writings WHERE user_id=? ORDER BY created_at DESC",
            (session["user_id"],),
        ).fetchall()
    return render_template("history.html", records=records)


@app.route("/count-print/<writing_id>", methods=["POST"])
@login_required
def count_print(writing_id):
    writing = writing_by_id(writing_id)
    if not writing:
        return jsonify({"error": "Not found"}), 404
    with db_connection() as db:
        db.execute(
            """UPDATE writings
               SET print_count=print_count+1,
                   download_count=download_count+1
               WHERE id=?""",
            (writing_id,),
        )
        db.commit()
    return jsonify({"success": True})


@app.route("/download/<writing_id>/<file_type>")
@login_required
def download(writing_id, file_type):
    writing = writing_by_id(writing_id)
    if not writing:
        flash("Writing not found.", "danger")
        return redirect(url_for("dashboard"))

    with db_connection() as db:
        db.execute(
            "UPDATE writings SET download_count=download_count+1 WHERE id=?",
            (writing_id,),
        )
        db.commit()

    subject, body = writing["subject"], writing["body"]
    safe_base = re.sub(r"[^A-Za-z0-9_-]+", "_", subject)[:60] or "writing"

    if file_type == "txt":
        return send_file(
            io.BytesIO(f"Subject: {subject}\n\n{body}".encode()),
            as_attachment=True,
            download_name=f"{safe_base}.txt",
            mimetype="text/plain",
        )

    if file_type == "docx":
        document = Document()
        document.add_heading(subject, level=1)
        for paragraph in body.split("\n\n"):
            document.add_paragraph(paragraph)
        if writing["signature_file"]:
            document.add_picture(str(UPLOAD_DIR / writing["signature_file"]))
        if writing["attachment_original_name"]:
            document.add_paragraph(f"Attachment: {writing['attachment_original_name']}")
        output = io.BytesIO()
        document.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=f"{safe_base}.docx")

    if file_type == "pdf":
        output = io.BytesIO()
        document = SimpleDocTemplate(
            output,
            pagesize=A4,
            rightMargin=20 * mm,
            leftMargin=20 * mm,
            topMargin=20 * mm,
            bottomMargin=20 * mm,
        )
        styles = getSampleStyleSheet()
        story = [Paragraph(subject, styles["Title"]), Spacer(1, 10)]
        for paragraph in body.split("\n\n"):
            escaped = (
                paragraph.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\n", "<br/>")
            )
            story.extend([Paragraph(escaped, styles["BodyText"]), Spacer(1, 8)])
        if writing["signature_file"]:
            story.extend(
                [
                    Spacer(1, 10),
                    PDFImage(
                        str(UPLOAD_DIR / writing["signature_file"]),
                        width=50 * mm,
                        height=20 * mm,
                    ),
                ]
            )
        if writing["attachment_original_name"]:
            story.extend(
                [
                    Spacer(1, 15),
                    Paragraph(
                        f"Attachment: {writing['attachment_original_name']}",
                        styles["BodyText"],
                    ),
                ]
            )
        document.build(story)
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name=f"{safe_base}.pdf",
            mimetype="application/pdf",
        )

    if file_type == "png":
        width, margin = 1400, 80
        font = ImageFont.load_default(size=25)
        title_font = ImageFont.load_default(size=34)
        lines = []
        for paragraph in body.splitlines():
            words, line = paragraph.split(), ""
            for word in words:
                test = f"{line} {word}".strip()
                if len(test) > 85:
                    lines.append(line)
                    line = word
                else:
                    line = test
            lines.append(line)
        height = max(1000, 240 + len(lines) * 38)
        image = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(image)
        draw.text((margin, margin), subject, fill="black", font=title_font)
        y = margin + 70
        for line in lines:
            draw.text((margin, y), line, fill="black", font=font)
            y += 38
        output = io.BytesIO()
        image.save(output, "PNG")
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name=f"{safe_base}.png",
            mimetype="image/png",
        )

    flash("Unsupported download type.", "warning")
    return redirect(url_for("preview", writing_id=writing_id))


@app.route("/uploads/<filename>")
@login_required
def uploaded_file(filename):
    safe_name = secure_filename(filename)
    file_path = UPLOAD_DIR / safe_name
    if not file_path.exists():
        flash("File not found.", "danger")
        return redirect(url_for("dashboard"))
    return send_file(file_path)


@app.errorhandler(413)
def too_large(_):
    flash("File is too large. Maximum size is 10 MB.", "danger")
    return redirect(url_for("compose"))


if __name__ == "__main__":
    app.run(debug=True, threaded=True)
