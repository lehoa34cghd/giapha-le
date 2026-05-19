from flask import Flask, render_template, request, redirect, url_for, session, send_file
import sqlite3, os
from werkzeug.security import generate_password_hash, check_password_hash
from io import BytesIO
import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "doi_chuoi_bi_mat_nay")
DB = "giapha_le.db"

USERS = {
    "admin": {"password_hash": generate_password_hash("admin123"), "role": "admin"},
    "khach": {"password_hash": generate_password_hash("xem123"), "role": "viewer"},
}

def db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = db()
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS people(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        title TEXT,
        birth TEXT,
        death TEXT,
        note TEXT,
        parent_id INTEGER,
        spouse_id INTEGER,
        gender TEXT
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS comments(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person_id INTEGER,
        username TEXT,
        content TEXT NOT NULL,
        status TEXT DEFAULT 'new',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    count = cur.execute("SELECT COUNT(*) c FROM people").fetchone()["c"]
    if count == 0:
        cur.execute("INSERT INTO people(name,title,note) VALUES(?,?,?)", ("Ông Tổ Họ Lê", "Thủy tổ", "Dữ liệu mẫu ban đầu"))
        root = cur.lastrowid
        cur.execute("INSERT INTO people(name,title,parent_id) VALUES(?,?,?)", ("Lê Văn A", "Đời 1", root))
        cur.execute("INSERT INTO people(name,title,parent_id) VALUES(?,?,?)", ("Lê Văn B", "Đời 1", root))
    conn.commit(); conn.close()

def rows():
    conn = db(); data = conn.execute("SELECT * FROM people ORDER BY id").fetchall(); conn.close(); return data

def children(all_rows, parent_id):
    return [r for r in all_rows if r["parent_id"] == parent_id]

def render_node(p, all_rows, mode):
    admin = session.get("role") == "admin" and mode == "edit"
    spouses = [r for r in all_rows if r["spouse_id"] == p["id"]]
    action = ""
    if admin:
        action = f"<div class='actions'><a href='/person/{p['id']}/edit'>Sửa</a><a class='del' onclick=\"return confirm('Xóa người này?')\" href='/person/{p['id']}/delete'>Xóa</a></div>"
    card = f"<div class='card'><b>{p['name']}</b><span>{p['title'] or ''}</span>{action}</div>"
    spouse_cards = "".join([f"<div class='card spouse'><b>{s['name']}</b><span>{s['title'] or 'Vợ/chồng'}</span>{('<div class=actions><a href=/person/'+str(s['id'])+'/edit>Sửa</a><a class=del onclick=\"return confirm(\'Xóa người này?\')\" href=/person/'+str(s['id'])+'/delete>Xóa</a></div>') if admin else ''}</div>" for s in spouses])
    plus = ""
    if admin:
        plus = f"<div class='plus-row'><a class='plus' href='/person/add?parent_id={p['id']}'>+ con</a><a class='plus' href='/person/add?spouse_id={p['id']}'>+ vợ/chồng</a></div>"
    kids = children(all_rows, p["id"])
    kids_html = ""
    if kids:
        kids_html = "<ul>" + "".join(["<li>" + render_node(k, all_rows, mode) + "</li>" for k in kids]) + "</ul>"
    return f"<div class='couple'>{card}{spouse_cards}</div>{plus}{kids_html}"

@app.route('/login', methods=['GET','POST'])
def login():
    error = ''
    if request.method == 'POST':
        u = request.form.get('username','')
        p = request.form.get('password','')
        user = USERS.get(u)
        if user and check_password_hash(user['password_hash'], p):
            session['user'] = u; session['role'] = user['role']; session['mode'] = 'public'
            return redirect(url_for('index'))
        error = 'Sai tài khoản hoặc mật khẩu'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

def require_login():
    if 'user' not in session: return False
    return True

@app.route('/')
def index():
    if not require_login(): return redirect(url_for('login'))
    mode = request.args.get('mode') or session.get('mode','public')
    if mode not in ['public','edit']: mode='public'
    if session.get('role') != 'admin': mode='public'
    session['mode'] = mode
    all_rows = rows()
    roots = [r for r in all_rows if r['parent_id'] is None and r['spouse_id'] is None]
    tree = "<ul>" + "".join(["<li>" + render_node(r, all_rows, mode) + "</li>" for r in roots]) + "</ul>"
    conn=db(); pending=conn.execute("SELECT COUNT(*) c FROM comments WHERE status='new'").fetchone()['c']; conn.close()
    return render_template('index.html', tree=tree, mode=mode, pending=pending)

@app.route('/person/add', methods=['GET','POST'])
def add_person():
    if not require_login(): return redirect(url_for('login'))
    if session.get('role')!='admin': return redirect(url_for('index'))
    if request.method=='POST':
        conn=db(); conn.execute("INSERT INTO people(name,title,birth,death,note,parent_id,spouse_id,gender) VALUES(?,?,?,?,?,?,?,?)", (
            request.form['name'], request.form.get('title'), request.form.get('birth'), request.form.get('death'), request.form.get('note'),
            request.form.get('parent_id') or None, request.form.get('spouse_id') or None, request.form.get('gender')
        )); conn.commit(); conn.close(); return redirect(url_for('index', mode='edit'))
    return render_template('edit_person.html', people=rows(), p=None, parent_id=request.args.get('parent_id',''), spouse_id=request.args.get('spouse_id',''))

@app.route('/person/<int:id>/edit', methods=['GET','POST'])
def edit_person(id):
    if not require_login(): return redirect(url_for('login'))
    if session.get('role')!='admin': return redirect(url_for('index'))
    conn=db(); p=conn.execute('SELECT * FROM people WHERE id=?',(id,)).fetchone()
    if request.method=='POST':
        conn.execute("UPDATE people SET name=?,title=?,birth=?,death=?,note=?,parent_id=?,spouse_id=?,gender=? WHERE id=?", (
            request.form['name'], request.form.get('title'), request.form.get('birth'), request.form.get('death'), request.form.get('note'),
            request.form.get('parent_id') or None, request.form.get('spouse_id') or None, request.form.get('gender'), id
        )); conn.commit(); conn.close(); return redirect(url_for('index', mode='edit'))
    conn.close(); return render_template('edit_person.html', people=rows(), p=p, parent_id=p['parent_id'] or '', spouse_id=p['spouse_id'] or '')

@app.route('/person/<int:id>/delete')
def delete_person(id):
    if not require_login(): return redirect(url_for('login'))
    if session.get('role')=='admin':
        conn=db(); conn.execute('DELETE FROM people WHERE id=?',(id,)); conn.commit(); conn.close()
    return redirect(url_for('index', mode='edit'))

@app.route('/comments', methods=['GET','POST'])
def comments():
    if not require_login(): return redirect(url_for('login'))
    conn=db()
    if request.method=='POST':
        conn.execute('INSERT INTO comments(username,content) VALUES(?,?)',(session.get('user'), request.form['content']))
        conn.commit()
    data=conn.execute('SELECT * FROM comments ORDER BY id DESC').fetchall(); conn.close()
    return render_template('comments.html', comments=data)

@app.route('/comment/<int:id>/done')
def comment_done(id):
    if session.get('role')=='admin':
        conn=db(); conn.execute("UPDATE comments SET status='done' WHERE id=?",(id,)); conn.commit(); conn.close()
    return redirect(url_for('comments'))

@app.route('/export/excel')
def export_excel():
    if not require_login(): return redirect(url_for('login'))
    df = pd.DataFrame([dict(r) for r in rows()])
    bio = BytesIO(); df.to_excel(bio, index=False); bio.seek(0)
    return send_file(bio, download_name='gia_pha_le.xlsx', as_attachment=True)

init_db()
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
