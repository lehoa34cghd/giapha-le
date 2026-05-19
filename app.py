from flask import Flask, render_template, request, redirect, url_for, session, abort, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from datetime import datetime
import sqlite3
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = "DOI_CHUOI_BI_MAT_NAY_KHI_DEPLOY"

# Luôn lưu database trong chính thư mục project.
# Như vậy dù bạn chạy app.py từ VS Code, CMD hay PowerShell ở thư mục khác,
# dữ liệu vẫn nằm cố định tại: <project>/giapha_le.db và không bị tạo nhầm DB mới.
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB = os.path.join(BASE_DIR, "giapha_le.db")

# Tài khoản mẫu. Lưu plain_password chỉ để kiểm tra dự phòng, tránh lỗi môi trường
# khi hash sinh ra khác phiên hoặc máy chưa đồng bộ thư viện. Khi dùng thật, bạn nên đổi mật khẩu.
USERS = {
    "admin": {"password_hash": generate_password_hash("admin123"), "plain_password": "admin123", "role": "admin"},
    "viewer": {"password_hash": generate_password_hash("viewer"), "plain_password": "xem123", "role": "viewer"},
}


@app.context_processor
def inject_comment_badge():
    """Hiển thị số góp ý chưa xử lý trên thanh menu."""
    if "user" not in session:
        return {"pending_comments_count": 0}
    try:
        conn = get_db()
        count = conn.execute("SELECT COUNT(*) AS c FROM comments WHERE status != 'done'").fetchone()["c"]
        conn.close()
        return {"pending_comments_count": count}
    except Exception:
        return {"pending_comments_count": 0}


def is_edit_mode():
    """Admin có thể chuyển giữa chế độ Sửa đổi và Công khai.
    Viewer luôn là Công khai.
    """
    return session.get("role") == "admin" and session.get("view_mode", "edit") == "edit"


@app.context_processor
def inject_view_mode():
    mode = session.get("view_mode", "edit" if session.get("role") == "admin" else "public")
    if session.get("role") != "admin":
        mode = "public"
    return {"view_mode": mode, "can_edit": is_edit_mode()}


def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS people (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            title TEXT,
            birth TEXT,
            death TEXT,
            note TEXT,
            parent_id INTEGER,
            spouse TEXT,
            gender TEXT DEFAULT 'unknown',
            FOREIGN KEY(parent_id) REFERENCES people(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            person_id INTEGER,
            username TEXT,
            content TEXT NOT NULL,
            status TEXT DEFAULT 'new',
            created_at TEXT,
            FOREIGN KEY(person_id) REFERENCES people(id)
        )
    """)
    conn.commit()

    # Nâng cấp database cũ: thêm cột đánh dấu vợ/chồng nếu chưa có.
    existing_cols = [row["name"] for row in cur.execute("PRAGMA table_info(people)").fetchall()]
    if "spouse_of" not in existing_cols:
        cur.execute("ALTER TABLE people ADD COLUMN spouse_of INTEGER")
        conn.commit()

    count = cur.execute("SELECT COUNT(*) AS c FROM people").fetchone()["c"]
    if count == 0:
        cur.execute("INSERT INTO people(name, title, note) VALUES (?, ?, ?)", ("Ông Tổ - Bà Tổ", "Thủy tổ họ Lê", "Gốc phả hệ"))
        root = cur.lastrowid
        for ten in ["Lê Văn Đức", "Lê Văn Phúc", "Lê Văn Nhân"]:
            cur.execute("INSERT INTO people(name, title, parent_id) VALUES (?, ?, ?)", (ten, "Đời 1", root))
            parent = cur.lastrowid
            for i in range(1, 4):
                cur.execute("INSERT INTO people(name, title, parent_id) VALUES (?, ?, ?)", (f"Lê Văn {ten.split()[-1]} {i}", "Đời 2", parent))
        conn.commit()
    conn.close()


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get("role") != "admin":
            abort(403)
        return fn(*args, **kwargs)
    return wrapper


def build_tree(rows, parent_id=None):
    """Tạo cây gia phả.
    - Người có spouse_of sẽ không đứng thành nhánh riêng.
    - Họ được gắn ngang bên cạnh người chính.
    """
    nodes = []
    for r in rows:
        if r["parent_id"] == parent_id and not r["spouse_of"]:
            item = dict(r)
            item["spouses"] = [dict(x) for x in rows if x["spouse_of"] == r["id"]]
            item["children"] = build_tree(rows, r["id"])
            nodes.append(item)
    return nodes


def export_people_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Gia pha ho Le"
    headers = ["ID", "Họ tên", "Vai vế", "Năm sinh", "Năm mất", "Vợ/Chồng", "Giới tính", "ID cấp trên", "Tên cấp trên", "ID vợ/chồng của", "Ghi chú"]
    ws.append(headers)

    fill = PatternFill("solid", fgColor="B70F0F")
    font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for r in rows:
        ws.append([
            r["id"], r["name"], r["title"], r["birth"], r["death"], r["spouse"],
            r["gender"], r["parent_id"], r["parent_name"], r["spouse_of"], r["note"]
        ])

    widths = [8, 26, 16, 12, 12, 22, 12, 12, 26, 14, 40]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = width

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "").strip()
        user = USERS.get(username)
        valid_password = False
        if user:
            try:
                valid_password = check_password_hash(user["password_hash"], password)
            except Exception:
                valid_password = False
            # Dự phòng để tài khoản khach / xem123 luôn đăng nhập được.
            if not valid_password and password == user.get("plain_password"):
                valid_password = True
        if user and valid_password:
            session["user"] = username
            session["role"] = user["role"]
            session["view_mode"] = "edit" if user["role"] == "admin" else "public"
            return redirect(url_for("index"))
        error = "Sai tài khoản hoặc mật khẩu"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/view-mode/<mode>")
@login_required
def set_view_mode(mode):
    if session.get("role") != "admin":
        session["view_mode"] = "public"
        return redirect(url_for("index"))
    if mode not in ["edit", "public"]:
        abort(404)
    session["view_mode"] = mode
    return redirect(request.referrer or url_for("index"))


@app.route("/")
@login_required
def index():
    conn = get_db()
    rows = conn.execute("SELECT * FROM people ORDER BY id").fetchall()
    conn.close()
    tree = build_tree(rows)
    return render_template("index.html", tree=tree)


@app.route("/search")
@login_required
def search():
    q = request.args.get("q", "").strip()
    rows = []
    if q:
        like = f"%{q}%"
        conn = get_db()
        rows = conn.execute("""
            SELECT p.*, parent.name AS parent_name
            FROM people p
            LEFT JOIN people parent ON p.parent_id = parent.id
            WHERE p.name LIKE ? OR p.title LIKE ? OR p.note LIKE ? OR p.spouse LIKE ?
            ORDER BY p.name
        """, (like, like, like, like)).fetchall()
        conn.close()
    return render_template("search.html", q=q, rows=rows)


@app.route("/export/excel")
@login_required
@admin_required
def export_excel():
    conn = get_db()
    rows = conn.execute("""
        SELECT p.*, parent.name AS parent_name
        FROM people p
        LEFT JOIN people parent ON p.parent_id = parent.id
        ORDER BY p.id
    """).fetchall()
    conn.close()
    output = export_people_excel(rows)
    filename = f"gia_pha_ho_le_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/export/image")
@login_required
@admin_required
def export_image():
    conn = get_db()
    rows = conn.execute("SELECT * FROM people ORDER BY id").fetchall()
    conn.close()
    tree = build_tree(rows)
    return render_template("export_image.html", tree=tree)


@app.route("/person/<int:person_id>", methods=["GET", "POST"])
@login_required
def person(person_id):
    conn = get_db()
    p = conn.execute("SELECT * FROM people WHERE id=?", (person_id,)).fetchone()
    if not p:
        conn.close()
        abort(404)
    if request.method == "POST":
        content = request.form.get("content", "").strip()
        if content:
            conn.execute(
                "INSERT INTO comments(person_id, username, content, created_at) VALUES (?, ?, ?, ?)",
                (person_id, session.get("user"), content, datetime.now().strftime("%Y-%m-%d %H:%M")),
            )
            conn.commit()
        conn.close()
        return redirect(url_for("person", person_id=person_id))
    comments = conn.execute("SELECT * FROM comments WHERE person_id=? ORDER BY id DESC", (person_id,)).fetchall()
    conn.close()
    return render_template("person.html", p=p, comments=comments)


@app.route("/comments", methods=["GET", "POST"])
@login_required
def comments():
    conn = get_db()
    if request.method == "POST":
        person_id = request.form.get("person_id") or None
        content = request.form.get("content", "").strip()
        if content:
            conn.execute(
                "INSERT INTO comments(person_id, username, content, created_at) VALUES (?, ?, ?, ?)",
                (person_id, session.get("user"), content, datetime.now().strftime("%Y-%m-%d %H:%M")),
            )
            conn.commit()
        conn.close()
        return redirect(url_for("comments"))

    rows = conn.execute("""
        SELECT c.*, p.name AS person_name
        FROM comments c LEFT JOIN people p ON c.person_id=p.id
        ORDER BY c.id DESC
    """).fetchall()
    people = conn.execute("SELECT id, name FROM people ORDER BY name").fetchall()
    conn.close()
    return render_template("comments.html", rows=rows, people=people)


@app.route("/admin/people")
@login_required
@admin_required
def admin_people():
    if not is_edit_mode():
        return redirect(url_for("index"))
    conn = get_db()
    rows = conn.execute("""
        SELECT p.*, parent.name AS parent_name
        FROM people p LEFT JOIN people parent ON p.parent_id=parent.id
        ORDER BY p.id
    """).fetchall()
    conn.close()
    return render_template("admin_people.html", rows=rows)


@app.route("/admin/person/<int:person_id>", methods=["GET", "POST"])
@login_required
@admin_required
def edit_person(person_id):
    if not is_edit_mode():
        return redirect(url_for("index"))
    conn = get_db()
    people = conn.execute("SELECT id, name FROM people ORDER BY id").fetchall()
    p = None if person_id == 0 else conn.execute("SELECT * FROM people WHERE id=?", (person_id,)).fetchone()

    relation = request.args.get("relation", "")
    ref_id = request.args.get("ref_id", type=int)
    ref = conn.execute("SELECT * FROM people WHERE id=?", (ref_id,)).fetchone() if ref_id else None

    defaults = {
        "name": "",
        "title": "",
        "birth": "",
        "death": "",
        "note": "",
        "parent_id": "",
        "spouse": "",
        "gender": "unknown",
        "spouse_of": "",
        "relation_mode": "branch",
    }

    if person_id == 0 and ref:
        if relation == "child":
            # Dấu + bên dưới: tạo con/đời sau, nối xuống dưới người được chọn.
            defaults["relation_mode"] = "branch"
            defaults["parent_id"] = ref["id"]
            defaults["title"] = "Đời sau"
            defaults["note"] = f"Con/đời sau của {ref['name']}"
        elif relation == "same":
            # Dấu + bên phải: mặc định tạo người cùng đời, vẫn cho đổi sang vợ/chồng trong form.
            defaults["relation_mode"] = "branch"
            defaults["parent_id"] = ref["parent_id"] or ""
            defaults["title"] = ref["title"] or ""
            defaults["note"] = f"Cùng đời với {ref['name']}"
        elif relation == "spouse":
            # Dấu + bên trái: mặc định tạo vợ/chồng nằm ngang bên cạnh, vẫn cho đổi sang thuộc nhánh.
            defaults["relation_mode"] = "spouse"
            defaults["parent_id"] = ""
            defaults["spouse_of"] = ref["id"]
            defaults["title"] = "Vợ/chồng"
            defaults["spouse"] = ref["name"]
            defaults["note"] = f"Vợ/chồng của {ref['name']}"

    if request.method == "POST":
        relation_mode = request.form.get("relation_mode", "branch")
        data = {k: request.form.get(k) or None for k in ["name", "title", "birth", "death", "note", "parent_id", "spouse", "gender", "spouse_of"]}

        # Chỉ dùng một kiểu liên kết tại một thời điểm:
        # - branch: thuộc nhánh của ai -> hiện dưới dạng con/đời sau hoặc cùng đời theo parent_id
        # - spouse: vợ/chồng của ai -> hiện ngang bên cạnh người được chọn
        if relation_mode == "spouse":
            data["parent_id"] = None
            if data["spouse_of"]:
                data["spouse_of"] = int(data["spouse_of"])
            elif relation == "spouse" and ref:
                data["spouse_of"] = ref["id"]
        else:
            data["spouse_of"] = None
            if data["parent_id"]:
                data["parent_id"] = int(data["parent_id"])

        if person_id == 0:
            conn.execute("""
                INSERT INTO people(name,title,birth,death,note,parent_id,spouse,gender,spouse_of)
                VALUES (:name,:title,:birth,:death,:note,:parent_id,:spouse,:gender,:spouse_of)
            """, data)

            # Nếu tạo dạng vợ/chồng, tự ghi tên vào ô Vợ/chồng của người gốc để dễ nhận biết khi export/sửa.
            if relation_mode == "spouse" and data["spouse_of"]:
                base = conn.execute("SELECT spouse FROM people WHERE id=?", (data["spouse_of"],)).fetchone()
                old_spouse = base["spouse"] if base else ""
                spouse_names = [x.strip() for x in (old_spouse or "").split(",") if x.strip()]
                if data["name"] not in spouse_names:
                    spouse_names.append(data["name"])
                conn.execute("UPDATE people SET spouse=? WHERE id=?", (", ".join(spouse_names), data["spouse_of"]))
        else:
            data["id"] = person_id
            conn.execute("""
                UPDATE people SET name=:name,title=:title,birth=:birth,death=:death,
                note=:note,parent_id=:parent_id,spouse=:spouse,gender=:gender,spouse_of=:spouse_of WHERE id=:id
            """, data)
        conn.commit()
        conn.close()
        return redirect(url_for("index"))

    conn.close()
    return render_template(
        "edit_person.html",
        p=p,
        people=people,
        person_id=person_id,
        defaults=defaults,
        relation=relation,
        ref=ref,
    )


@app.route("/admin/person/<int:person_id>/delete", methods=["POST"])
@login_required
@admin_required
def delete_person(person_id):
    if not is_edit_mode():
        return redirect(url_for("index"))
    conn = get_db()

    def collect_subtree(pid):
        ids = [pid]
        children = conn.execute("SELECT id FROM people WHERE parent_id=?", (pid,)).fetchall()
        spouses = conn.execute("SELECT id FROM people WHERE spouse_of=?", (pid,)).fetchall()
        for row in list(children) + list(spouses):
            ids.extend(collect_subtree(row["id"]))
        return ids

    ids = sorted(set(collect_subtree(person_id)))
    placeholders = ",".join(["?"] * len(ids))
    conn.execute(f"DELETE FROM comments WHERE person_id IN ({placeholders})", ids)
    conn.execute(f"DELETE FROM people WHERE id IN ({placeholders})", ids)
    conn.execute("UPDATE people SET spouse_of=NULL WHERE spouse_of=?", (person_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("index"))


@app.route("/admin/comment/<int:comment_id>/done")
@login_required
@admin_required
def comment_done(comment_id):
    conn = get_db()
    conn.execute("UPDATE comments SET status='done' WHERE id=?", (comment_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("comments"))


@app.errorhandler(403)
def forbidden(e):
    return render_template("403.html"), 403

init_db()
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
