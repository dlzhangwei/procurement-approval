import os
import sqlite3
import sys
import tempfile
from io import BytesIO
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def csrf(client):
    with client.session_transaction() as sess:
        return sess["csrf_token"]


def login(client, username, password):
    client.get("/login")
    response = client.post(
        "/login",
        data={
            "csrf_token": csrf(client),
            "username": username,
            "password": password,
        },
        follow_redirects=False,
    )
    assert response.status_code == 302, response.data.decode()


def main():
    with tempfile.TemporaryDirectory(prefix="procurement-smoke-") as tmp:
        tmp_path = Path(tmp)
        os.environ["DATABASE_PATH"] = str(tmp_path / "app.db")
        os.environ["UPLOAD_FOLDER"] = str(tmp_path / "uploads")
        os.environ["SECRET_KEY"] = "smoke-test-secret"

        from app import app

        app.config.update(TESTING=True)
        client = app.test_client()

        login(client, "user1", "user1123")

        image = BytesIO()
        Image.new("RGB", (3000, 3000), (170, 42, 54)).save(image, "PNG")
        image.seek(0)
        invoice = BytesIO(b"%PDF-1.4\n% procurement smoke test\n")

        response = client.post(
            "/requests/new",
            data={
                "csrf_token": csrf(client),
                "request_date": "2026-05-28",
                "content": "Smoke test purchase request",
                "images": (image, "purchase.png"),
                "invoice": (invoice, "invoice.pdf"),
            },
            content_type="multipart/form-data",
            follow_redirects=False,
        )
        assert response.status_code == 302, response.data.decode()
        request_id = int(response.headers["Location"].rstrip("/").split("/")[-1])

        db = sqlite3.connect(os.environ["DATABASE_PATH"])
        db.row_factory = sqlite3.Row
        image_row = db.execute(
            """
            SELECT size_bytes
            FROM attachments
            WHERE request_id = ? AND kind = 'image'
            """,
            (request_id,),
        ).fetchone()
        assert image_row is not None
        assert image_row["size_bytes"] <= 500 * 1024

        response = client.post(
            "/logout",
            data={"csrf_token": csrf(client)},
            follow_redirects=False,
        )
        assert response.status_code == 302

        login(client, "approver", "approver123")
        response = client.post(
            f"/requests/{request_id}/approve",
            data={"csrf_token": csrf(client), "decision_comment": "OK"},
            follow_redirects=False,
        )
        assert response.status_code == 302, response.data.decode()

        status = db.execute(
            "SELECT status FROM purchase_requests WHERE id = ?", (request_id,)
        ).fetchone()["status"]
        assert status == "approved"
        db.close()

    print("smoke-ok")


if __name__ == "__main__":
    main()
