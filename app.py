import io
import os
import secrets
import sqlite3
import uuid
from datetime import date, datetime
from functools import wraps
from pathlib import Path

from flask import (
    Flask,
    abort,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_from_directory,
    session,
    url_for,
)
from PIL import Image, ImageOps, UnidentifiedImageError
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass


BASE_DIR = Path(__file__).resolve().parent
Image.MAX_IMAGE_PIXELS = 30_000_000

app = Flask(__name__, instance_relative_config=True)
app.config.from_mapping(
    SECRET_KEY=os.environ.get("SECRET_KEY", "dev-change-me-before-deploy"),
    DATABASE=os.environ.get("DATABASE_PATH", os.path.join(app.instance_path, "app.db")),
    UPLOAD_FOLDER=os.environ.get(
        "UPLOAD_FOLDER", os.path.join(app.instance_path, "uploads")
    ),
    MAX_CONTENT_LENGTH=16 * 1024 * 1024,
    IMAGE_MAX_BYTES=500 * 1024,
    INVOICE_MAX_BYTES=10 * 1024 * 1024,
)

STATUS_LABELS = {
    "pending": "待审批",
    "approved": "已同意",
    "rejected": "未同意",
}

STATUS_CLASSES = {
    "pending": "status-pending",
    "approved": "status-approved",
    "rejected": "status-rejected",
}

DEFAULT_USERS = [
    {
        "username": "approver",
        "display_name": "审批人",
        "role": "approver",
        "password_env": "APPROVER_PASSWORD",
        "default_password": "approver123",
    },
    {
        "username": "user1",
        "display_name": "申请用户一",
        "role": "user",
        "password_env": "USER1_PASSWORD",
        "default_password": "user1123",
    },
    {
        "username": "user2",
        "display_name": "申请用户二",
        "role": "user",
        "password_env": "USER2_PASSWORD",
        "default_password": "user2123",
    },
]

SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    display_name TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('approver', 'user')),
    password_hash TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS purchase_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    requester_id INTEGER NOT NULL REFERENCES users(id),
    request_date TEXT NOT NULL,
    content TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending'
        CHECK (status IN ('pending', 'approved', 'rejected')),
    decision_comment TEXT,
    decided_by INTEGER REFERENCES users(id),
    decided_at TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS attachments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    request_id INTEGER NOT NULL REFERENCES purchase_requests(id)
        ON DELETE CASCADE,
    kind TEXT NOT NULL CHECK (kind IN ('image', 'invoice')),
    original_name TEXT NOT NULL,
    stored_name TEXT NOT NULL UNIQUE,
    mime_type TEXT NOT NULL,
    size_bytes INTEGER NOT NULL,
    uploaded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_purchase_requests_status
    ON purchase_requests(status);
CREATE INDEX IF NOT EXISTS idx_purchase_requests_created_at
    ON purchase_requests(created_at);
CREATE INDEX IF NOT EXISTS idx_attachments_request_id
    ON attachments(request_id);
"""


class ValidationError(Exception):
    pass


def utc_now() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(app.config["DATABASE"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(error=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    os.makedirs(app.instance_path, exist_ok=True)
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

    db = sqlite3.connect(app.config["DATABASE"])
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    seed_users(db)
    db.commit()
    db.close()


def seed_users(db):
    for user in DEFAULT_USERS:
        password = os.environ.get(user["password_env"], user["default_password"])
        password_hash = generate_password_hash(password)
        existing = db.execute(
            "SELECT id FROM users WHERE username = ?", (user["username"],)
        ).fetchone()

        if existing:
            db.execute(
                """
                UPDATE users
                SET display_name = ?, role = ?, password_hash = ?
                WHERE username = ?
                """,
                (
                    user["display_name"],
                    user["role"],
                    password_hash,
                    user["username"],
                ),
            )
        else:
            db.execute(
                """
                INSERT INTO users (username, display_name, role, password_hash)
                VALUES (?, ?, ?, ?)
                """,
                (
                    user["username"],
                    user["display_name"],
                    user["role"],
                    password_hash,
                ),
            )


@app.before_request
def load_current_user():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)

    user_id = session.get("user_id")
    g.user = None
    if user_id is not None:
        g.user = get_db().execute(
            "SELECT id, username, display_name, role FROM users WHERE id = ?",
            (user_id,),
        ).fetchone()


@app.context_processor
def inject_globals():
    return {
        "current_user": g.get("user"),
        "csrf_token": session.get("csrf_token"),
        "status_labels": STATUS_LABELS,
        "status_classes": STATUS_CLASSES,
    }


def login_required(view):
    @wraps(view)
    def wrapped_view(**kwargs):
        if g.user is None:
            return redirect(url_for("login", next=request.path))
        return view(**kwargs)

    return wrapped_view


def approver_required(view):
    @wraps(view)
    def wrapped_view(**kwargs):
        if g.user is None:
            return redirect(url_for("login", next=request.path))
        if g.user["role"] != "approver":
            abort(403)
        return view(**kwargs)

    return wrapped_view


def validate_csrf():
    token = request.form.get("csrf_token")
    if not token or token != session.get("csrf_token"):
        abort(400)


def request_summary(content: str) -> str:
    cleaned = " ".join(content.split())
    if len(cleaned) <= 42:
        return cleaned
    return cleaned[:42] + "..."


def parse_request_date(value: str) -> str:
    try:
        return date.fromisoformat(value).isoformat()
    except (TypeError, ValueError):
        raise ValidationError("请选择有效的申请日期。")


def normalize_image(image: Image.Image) -> Image.Image:
    image = ImageOps.exif_transpose(image)
    image.thumbnail((2400, 2400), Image.Resampling.LANCZOS)

    if image.mode in ("RGBA", "LA") or (
        image.mode == "P" and "transparency" in image.info
    ):
        background = Image.new("RGB", image.size, (255, 255, 255))
        alpha = image.convert("RGBA").getchannel("A")
        background.paste(image.convert("RGBA"), mask=alpha)
        return background

    return image.convert("RGB")


def compress_image(file_storage) -> tuple[str, str, int]:
    try:
        file_storage.stream.seek(0)
        with Image.open(file_storage.stream) as opened:
            original = normalize_image(opened)
    except (UnidentifiedImageError, OSError):
        raise ValidationError("图片文件无法识别，请上传 jpg、png 或 webp。")

    working = original.copy()
    quality = 86
    scale = 1.0
    payload = None

    for _ in range(36):
        buffer = io.BytesIO()
        working.save(buffer, "JPEG", quality=quality, optimize=True, progressive=True)
        size = buffer.tell()

        if size <= app.config["IMAGE_MAX_BYTES"]:
            payload = buffer.getvalue()
            break

        if quality > 42:
            quality -= 8
            continue

        scale *= 0.82
        width = max(240, int(original.width * scale))
        height = max(240, int(original.height * scale))
        if width == working.width and height == working.height:
            quality = max(30, quality - 4)
            continue
        working = original.resize((width, height), Image.Resampling.LANCZOS)
        quality = 78

    if payload is None:
        raise ValidationError("图片压缩后仍超过 500KB，请换一张更小的图片。")

    stored_name = f"{uuid.uuid4().hex}.jpg"
    path = os.path.join(app.config["UPLOAD_FOLDER"], stored_name)
    with open(path, "wb") as output:
        output.write(payload)

    return stored_name, "image/jpeg", len(payload)


def save_invoice_pdf(file_storage) -> tuple[str, str, int]:
    file_storage.stream.seek(0)
    payload = file_storage.read()
    if len(payload) > app.config["INVOICE_MAX_BYTES"]:
        raise ValidationError("PDF 发票不能超过 10MB。")

    stored_name = f"{uuid.uuid4().hex}.pdf"
    path = os.path.join(app.config["UPLOAD_FOLDER"], stored_name)
    with open(path, "wb") as output:
        output.write(payload)

    return stored_name, "application/pdf", len(payload)


def save_attachment(file_storage, kind: str):
    original_name = secure_filename(file_storage.filename or "")
    if not original_name:
        return None

    suffix = Path(original_name).suffix.lower()
    if kind == "invoice" and suffix == ".pdf":
        stored_name, mime_type, size_bytes = save_invoice_pdf(file_storage)
    else:
        stored_name, mime_type, size_bytes = compress_image(file_storage)

    return {
        "kind": kind,
        "original_name": original_name,
        "stored_name": stored_name,
        "mime_type": mime_type,
        "size_bytes": size_bytes,
    }


def insert_attachment(db, request_id: int, attachment):
    if attachment is None:
        return

    db.execute(
        """
        INSERT INTO attachments
            (request_id, kind, original_name, stored_name, mime_type, size_bytes)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            request_id,
            attachment["kind"],
            attachment["original_name"],
            attachment["stored_name"],
            attachment["mime_type"],
            attachment["size_bytes"],
        ),
    )


def load_request(request_id: int):
    purchase_request = get_db().execute(
        """
        SELECT
            pr.*,
            requester.display_name AS requester_name,
            approver.display_name AS approver_name
        FROM purchase_requests pr
        JOIN users requester ON requester.id = pr.requester_id
        LEFT JOIN users approver ON approver.id = pr.decided_by
        WHERE pr.id = ?
        """,
        (request_id,),
    ).fetchone()

    if purchase_request is None:
        abort(404)
    return purchase_request


def status_counts():
    rows = get_db().execute(
        "SELECT status, COUNT(*) AS count FROM purchase_requests GROUP BY status"
    ).fetchall()
    counts = {status: 0 for status in STATUS_LABELS}
    for row in rows:
        counts[row["status"]] = row["count"]
    counts["all"] = sum(counts.values())
    return counts


@app.template_filter("bytes")
def format_bytes(value):
    value = int(value or 0)
    if value < 1024:
        return f"{value} B"
    if value < 1024 * 1024:
        return f"{value / 1024:.1f} KB"
    return f"{value / 1024 / 1024:.1f} MB"


@app.template_filter("datetime")
def format_datetime(value):
    if not value:
        return ""
    return str(value).replace("T", " ").replace("Z", "")


@app.route("/login", methods=("GET", "POST"))
def login():
    if g.user is not None:
        return redirect(url_for("index"))

    if request.method == "POST":
        validate_csrf()
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = get_db().execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()

        if user is None or not check_password_hash(user["password_hash"], password):
            flash("用户名或密码不正确。", "error")
        else:
            session.clear()
            session["user_id"] = user["id"]
            session["csrf_token"] = secrets.token_urlsafe(32)
            return redirect(request.args.get("next") or url_for("index"))

    return render_template("login.html")


@app.route("/logout", methods=("POST",))
@login_required
def logout():
    validate_csrf()
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    selected_status = request.args.get("status", "all")
    if selected_status not in {"all", *STATUS_LABELS.keys()}:
        selected_status = "all"

    params = []
    status_clause = ""
    if selected_status != "all":
        status_clause = "WHERE pr.status = ?"
        params.append(selected_status)

    requests = get_db().execute(
        f"""
        SELECT
            pr.*,
            requester.display_name AS requester_name,
            COUNT(a.id) AS attachment_count
        FROM purchase_requests pr
        JOIN users requester ON requester.id = pr.requester_id
        LEFT JOIN attachments a ON a.request_id = pr.id
        {status_clause}
        GROUP BY pr.id
        ORDER BY pr.created_at DESC, pr.id DESC
        """,
        params,
    ).fetchall()

    return render_template(
        "index.html",
        requests=requests,
        counts=status_counts(),
        selected_status=selected_status,
        summary=request_summary,
    )


@app.route("/requests/new", methods=("GET", "POST"))
@login_required
def new_request():
    if request.method == "POST":
        validate_csrf()
        content = request.form.get("content", "").strip()
        request_date = request.form.get("request_date", "")

        try:
            parsed_date = parse_request_date(request_date)
            if not content:
                raise ValidationError("请填写采购内容。")

            saved_files = []
            db = get_db()
            cursor = db.execute(
                """
                INSERT INTO purchase_requests
                    (requester_id, request_date, content, updated_at)
                VALUES (?, ?, ?, ?)
                """,
                (g.user["id"], parsed_date, content, utc_now()),
            )
            request_id = cursor.lastrowid

            for image in request.files.getlist("images"):
                attachment = save_attachment(image, "image")
                if attachment:
                    saved_files.append(attachment["stored_name"])
                    insert_attachment(db, request_id, attachment)

            for invoice in request.files.getlist("invoice"):
                attachment = save_attachment(invoice, "invoice")
                if attachment:
                    saved_files.append(attachment["stored_name"])
                    insert_attachment(db, request_id, attachment)

            db.commit()
            flash("采购申请已提交。", "success")
            return redirect(url_for("request_detail", request_id=request_id))
        except ValidationError as error:
            get_db().rollback()
            for stored_name in locals().get("saved_files", []):
                path = os.path.join(app.config["UPLOAD_FOLDER"], stored_name)
                if os.path.exists(path):
                    os.remove(path)
            flash(str(error), "error")

    return render_template(
        "request_form.html",
        today=date.today().isoformat(),
        content=request.form.get("content", ""),
        request_date=request.form.get("request_date", date.today().isoformat()),
    )


@app.route("/requests/<int:request_id>")
@login_required
def request_detail(request_id):
    purchase_request = load_request(request_id)
    attachments = get_db().execute(
        "SELECT * FROM attachments WHERE request_id = ? ORDER BY id",
        (request_id,),
    ).fetchall()

    return render_template(
        "request_detail.html",
        purchase_request=purchase_request,
        attachments=attachments,
    )


@app.route("/requests/<int:request_id>/approve", methods=("POST",))
@approver_required
def approve_request(request_id):
    validate_csrf()
    decision_comment = request.form.get("decision_comment", "").strip() or None
    load_request(request_id)
    get_db().execute(
        """
        UPDATE purchase_requests
        SET status = 'approved',
            decision_comment = ?,
            decided_by = ?,
            decided_at = ?,
            updated_at = ?
        WHERE id = ?
        """,
        (decision_comment, g.user["id"], utc_now(), utc_now(), request_id),
    )
    get_db().commit()
    flash("已同意该采购申请。", "success")
    return redirect(url_for("request_detail", request_id=request_id))


@app.route("/requests/<int:request_id>/reject", methods=("POST",))
@approver_required
def reject_request(request_id):
    validate_csrf()
    decision_comment = request.form.get("decision_comment", "").strip() or None
    load_request(request_id)
    get_db().execute(
        """
        UPDATE purchase_requests
        SET status = 'rejected',
            decision_comment = ?,
            decided_by = ?,
            decided_at = ?,
            updated_at = ?
        WHERE id = ?
        """,
        (decision_comment, g.user["id"], utc_now(), utc_now(), request_id),
    )
    get_db().commit()
    flash("已标记为未同意。", "success")
    return redirect(url_for("request_detail", request_id=request_id))


@app.route("/uploads/<path:filename>")
@login_required
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


@app.errorhandler(413)
def payload_too_large(error):
    flash("上传内容太大，请减少文件数量或压缩后再上传。", "error")
    return redirect(request.referrer or url_for("new_request"))


init_db()


if __name__ == "__main__":
    app.run(debug=True)
