# -*- coding: utf-8 -*-
"""
Nisira GCH Mobile - Render Ready
Login + Home + Asistencia GPS + Firma Digital + Documentos + Perfil.
Usuario demo: 11223344 / 123456
"""
import os, re, sqlite3, base64
from datetime import datetime, date
from functools import wraps
from flask import Flask, request, redirect, url_for, session, flash, jsonify, send_file, render_template, Response
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import psycopg2
    import psycopg2.extras
except Exception:
    psycopg2 = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR = os.getenv("PERSIST_DIR", "/data" if os.path.isdir("/data") else os.path.join(BASE_DIR, "data"))
DOC_DIR = os.path.join(PERSIST_DIR, "documentos")
SIGN_DIR = os.path.join(PERSIST_DIR, "firmas")
os.makedirs(DOC_DIR, exist_ok=True)
os.makedirs(SIGN_DIR, exist_ok=True)
DB_PATH = os.path.join(PERSIST_DIR, "nisira_gch.db")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
USE_POSTGRES = bool(DATABASE_URL and psycopg2)

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "cambiar-clave-en-render")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


def is_pg():
    return USE_POSTGRES


def q(sql):
    return sql.replace("?", "%s") if is_pg() else sql


def get_conn():
    if is_pg():
        return psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def execute(sql, params=(), one=False, all=False, commit=False):
    con = get_conn(); cur = con.cursor()
    cur.execute(q(sql), params)
    data = None
    if one: data = cur.fetchone()
    if all: data = cur.fetchall()
    if commit: con.commit()
    cur.close(); con.close()
    return data


def rowdict(r):
    return dict(r) if r else None


def rowsdict(rows):
    return [dict(x) for x in rows or []]


def now_str(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def today_str(): return date.today().isoformat()
def clean_dni(v): return re.sub(r"\D", "", str(v or ""))[-8:]


def init_db():
    con = get_conn(); cur = con.cursor()
    idtype = "SERIAL PRIMARY KEY" if is_pg() else "INTEGER PRIMARY KEY AUTOINCREMENT"
    cur.execute(q(f"""CREATE TABLE IF NOT EXISTS empresas(
        id {idtype}, nombre TEXT UNIQUE, ruc TEXT, estado TEXT DEFAULT 'ACTIVO')"""))
    cur.execute(q(f"""CREATE TABLE IF NOT EXISTS usuarios(
        id {idtype}, dni TEXT UNIQUE NOT NULL, nombres TEXT, empresa TEXT,
        password_hash TEXT, rol TEXT DEFAULT 'trabajador', cargo TEXT, area TEXT,
        foto TEXT, firma_path TEXT, creado_en TEXT)"""))
    cur.execute(q(f"""CREATE TABLE IF NOT EXISTS asistencias(
        id {idtype}, dni TEXT, nombres TEXT, empresa TEXT, tipo TEXT, fecha TEXT,
        hora TEXT, fecha_hora TEXT, latitud TEXT, longitud TEXT, direccion TEXT,
        precision_gps TEXT, observacion TEXT)"""))
    cur.execute(q(f"""CREATE TABLE IF NOT EXISTS documentos(
        id {idtype}, dni TEXT, categoria TEXT, titulo TEXT, periodo TEXT, filename TEXT,
        content_type TEXT, size_bytes INTEGER, creado_en TEXT)"""))
    cur.execute(q("INSERT INTO empresas(nombre,ruc,estado) VALUES(?,?,?) ON CONFLICT DO NOTHING" if is_pg() else "INSERT OR IGNORE INTO empresas(nombre,ruc,estado) VALUES(?,?,?)"),
                ("NISIRA SYSTEMS S.A.C", "20600000000", "ACTIVO"))
    cur.execute(q("SELECT id FROM usuarios WHERE dni=?"), ("11223344",))
    if not cur.fetchone():
        cur.execute(q("""INSERT INTO usuarios(dni,nombres,empresa,password_hash,rol,cargo,area,creado_en)
            VALUES(?,?,?,?,?,?,?,?)"""),
            ("11223344", "OMAR AZABACHE", "NISIRA SYSTEMS S.A.C", generate_password_hash("123456"),
             "admin", "Analista RRHH", "Gestión del Capital Humano", now_str()))
    con.commit(); cur.close(); con.close()
    seed_documents()


def seed_documents():
    demo = [
        ("Boletas Normales", "Boleta Febrero 2026", "FEBRERO 2026", "boleta_febrero_2026.pdf"),
        ("Empresa", "Contrato Anexo", "2026", "contrato_anexo.pdf"),
        ("Empresa", "Política de SST", "2026", "politica_sst.pdf"),
        ("Boletas Normales", "Boleta Enero 2026", "ENERO 2026", "boleta_enero_2026.pdf"),
        ("Constancias de Utilidades", "Constancia de Utilidades", "2026", "constancia_utilidades.pdf"),
        ("Boletas de Vacaciones", "Boleta de Vacaciones", "2026", "vacaciones_2026.pdf"),
        ("Constancias de CTS", "Constancia de CTS", "2026", "constancia_cts.pdf"),
        ("Constancias de Liquidaciones", "Constancia de Liquidación", "2026", "liquidacion_2026.pdf"),
        ("Boletas de Gratificaciones", "Boleta de Gratificación", "2026", "gratificacion_2026.pdf"),
        ("Constancias de Gratificaciones", "Constancia de Gratificación", "2026", "constancia_gratificacion_2026.pdf"),
    ]
    for categoria, titulo, periodo, filename in demo:
        path = os.path.join(DOC_DIR, filename)
        if not os.path.exists(path):
            pdf = f"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Count 0>>endobj\n% {titulo}\ntrailer<</Root 1 0 R>>\n%%EOF".encode()
            open(path, "wb").write(pdf)
        exists = execute("SELECT id FROM documentos WHERE dni=? AND titulo=?", ("11223344", titulo), one=True)
        if not exists:
            execute("""INSERT INTO documentos(dni,categoria,titulo,periodo,filename,content_type,size_bytes,creado_en)
                VALUES(?,?,?,?,?,?,?,?)""", ("11223344", categoria, titulo, periodo, filename, "application/pdf", os.path.getsize(path), now_str()), commit=True)


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("dni"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def current_user():
    return rowdict(execute("SELECT * FROM usuarios WHERE dni=?", (session.get("dni"),), one=True))


@app.context_processor
def inject_user():
    return {"current_user": current_user() if session.get("dni") else None}


@app.route("/")
def index():
    return redirect(url_for("home") if session.get("dni") else url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    empresas = rowsdict(execute("SELECT nombre FROM empresas WHERE estado='ACTIVO' ORDER BY nombre", all=True))
    if request.method == "POST":
        dni = clean_dni(request.form.get("dni"))
        password = request.form.get("password", "")
        empresa = request.form.get("empresa", "")
        user = rowdict(execute("SELECT * FROM usuarios WHERE dni=?", (dni,), one=True))
        if user and check_password_hash(user["password_hash"], password) and user["empresa"] == empresa:
            session.clear(); session["dni"] = dni; session["empresa"] = empresa
            return redirect(url_for("splash"))
        flash("Usuario, empresa o contraseña incorrectos.", "danger")
    return render_template("login.html", empresas=empresas)


@app.route("/splash")
@login_required
def splash():
    return render_template("splash.html")


@app.route("/home")
@login_required
def home():
    u = current_user()
    docs = rowsdict(execute("SELECT * FROM documentos WHERE dni=? ORDER BY creado_en DESC LIMIT 4", (u["dni"],), all=True))
    return render_template("home.html", docs=docs, active="home")


@app.route("/asistencia")
@login_required
def asistencia():
    u = current_user()
    ultimas = rowsdict(execute("SELECT * FROM asistencias WHERE dni=? ORDER BY id DESC LIMIT 5", (u["dni"],), all=True))
    return render_template("asistencia.html", ultimas=ultimas, active="asistencia")


@app.post("/api/marcar")
@login_required
def api_marcar():
    u = current_user()
    tipo = request.form.get("tipo", "INGRESO")
    hora = datetime.now().strftime("%H:%M:%S")
    execute("""INSERT INTO asistencias(dni,nombres,empresa,tipo,fecha,hora,fecha_hora,latitud,longitud,direccion,precision_gps,observacion)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
        (u["dni"], u["nombres"], u["empresa"], tipo, today_str(), hora, now_str(),
         request.form.get("latitud"), request.form.get("longitud"), request.form.get("direccion"),
         request.form.get("precision_gps"), request.form.get("observacion")), commit=True)
    return jsonify(ok=True, msg=f"Marcación {tipo.replace('_',' ').title()} registrada correctamente a las {hora}.")


@app.route("/firma", methods=["GET", "POST"])
@login_required
def firma():
    u = current_user()
    if request.method == "POST":
        data = request.form.get("firma_data", "")
        if data.startswith("data:image"):
            raw = base64.b64decode(data.split(",", 1)[1])
            filename = f"firma_{u['dni']}.png"
            open(os.path.join(SIGN_DIR, filename), "wb").write(raw)
            execute("UPDATE usuarios SET firma_path=? WHERE dni=?", (filename, u["dni"]), commit=True)
            flash("Firma guardada correctamente.", "success")
            return redirect(url_for("firma"))
        flash("Dibuja o sube una firma antes de guardar.", "danger")
    firma_b64 = None
    u = current_user()
    if u.get("firma_path"):
        path = os.path.join(SIGN_DIR, u["firma_path"])
        if os.path.exists(path):
            firma_b64 = base64.b64encode(open(path, "rb").read()).decode()
    return render_template("firma.html", firma_b64=firma_b64, active="perfil")


@app.get("/firma/descargar")
@login_required
def descargar_firma():
    u = current_user()
    if not u.get("firma_path"):
        flash("Aún no tienes firma guardada.", "danger")
        return redirect(url_for("firma"))
    path = os.path.join(SIGN_DIR, u["firma_path"])
    return send_file(path, as_attachment=True, download_name=u["firma_path"])


@app.route("/documentos")
@login_required
def documentos():
    cats = [
        ("Boletas Normales", "bi-file-earmark-text"),
        ("Constancias de Utilidades", "bi-card-checklist"),
        ("Boletas de Vacaciones", "bi-balloon-heart"),
        ("Constancias de CTS", "bi-bank"),
        ("Constancias de Liquidaciones", "bi-file-earmark-arrow-down"),
        ("Boletas de Gratificaciones", "bi-gift"),
        ("Constancias de Gratificaciones", "bi-patch-check"),
    ]
    return render_template("documentos.html", cats=cats, active="documentos")


@app.route("/documentos/<categoria>")
@login_required
def documentos_categoria(categoria):
    docs = rowsdict(execute("SELECT * FROM documentos WHERE dni=? AND categoria=? ORDER BY creado_en DESC", (session["dni"], categoria), all=True))
    return render_template("documentos_categoria.html", categoria=categoria, docs=docs, active="documentos")


@app.route("/documento/<int:doc_id>")
@login_required
def descargar_documento(doc_id):
    doc = rowdict(execute("SELECT * FROM documentos WHERE id=? AND dni=?", (doc_id, session["dni"]), one=True))
    if not doc:
        flash("Documento no encontrado.", "danger")
        return redirect(url_for("documentos"))
    return send_file(os.path.join(DOC_DIR, doc["filename"]), as_attachment=True, download_name=doc["filename"], mimetype=doc.get("content_type") or "application/pdf")


@app.route("/perfil")
@login_required
def perfil():
    total = rowdict(execute("SELECT COUNT(*) AS c FROM asistencias WHERE dni=?", (session["dni"],), one=True))["c"]
    return render_template("perfil.html", total=total, active="perfil")


@app.route("/promo/<tipo>")
def promo(tipo):
    data = {
        "asistencia": ("bi-calendar3", "Registra tu asistencia fácil", "Registra tu ingreso, salida y refrigerio de forma rápida y segura desde tu celular."),
        "documentos": ("bi-file-earmark-text-fill", "Gestión de documentos", "Accede a tus boletas, constancias y documentos laborales de forma rápida, organizada y segura."),
        "todo": ("bi-grid-fill", "Todo en un solo lugar", "Accede rápidamente a tus documentos y funciones más importantes desde el inicio."),
    }.get(tipo, ("bi-grid-fill", "Todo en un solo lugar", "Accede rápidamente a tus funciones."))
    return render_template("promo.html", icon=data[0], title=data[1], text=data[2])


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/manifest.json")
def manifest():
    return send_file(os.path.join(BASE_DIR, "static", "manifest.json"), mimetype="application/manifest+json")


@app.route("/sw.js")
def sw():
    return Response(open(os.path.join(BASE_DIR, "static", "sw.js"), encoding="utf-8").read(), mimetype="text/javascript")


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
else:
    init_db()
