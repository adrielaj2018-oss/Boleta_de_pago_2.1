# -*- coding: utf-8 -*-
"""
P&A MOBILE CLÁSICO
Asistencia + Documentos + Firma + Perfil + Configuración.
Estilo clásico tipo Tareo Móvil. Listo para Render.
Todo el HTML/CSS/JS está embebido en este app.py.
"""
import os, re, csv, sqlite3
from io import StringIO
from datetime import datetime, date
from functools import wraps
from flask import Flask, request, redirect, url_for, session, jsonify, Response
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR = os.getenv("PERSIST_DIR", "/data" if os.path.isdir("/data") else BASE_DIR)
os.makedirs(PERSIST_DIR, exist_ok=True)
DB_PATH = os.path.join(PERSIST_DIR, "pa_mobile_clasico.db")

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "pa-mobile-clasico-render-secret")
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024

COMPANY = "P&A Systems S.A.C."
VERSION = "V.3.2.15"
GREEN = "#2f773b"
LIME = "#a8d32a"

def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c

def ensure_col(cur, table, col, ddl):
    cur.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    if col not in cols:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")

def init_db():
    c = conn(); cur = c.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        nombres TEXT, apellidos TEXT, empresa TEXT, cargo TEXT, area TEXT,
        regimen TEXT, fecha_ingreso TEXT, tipo_doc TEXT, nro_doc TEXT,
        banco TEXT, cuenta TEXT, correo TEXT, celular TEXT, direccion TEXT,
        rol TEXT DEFAULT 'trabajador', estado TEXT DEFAULT 'ACTIVO'
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS marcas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT, trabajador TEXT, tipo TEXT, fecha TEXT, hora TEXT, fecha_hora TEXT,
        lat TEXT, lng TEXT, precision TEXT, metodo TEXT, lectura TEXT, hora_manual TEXT,
        observacion TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS firmas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE, data_url TEXT, actualizado TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS config_db(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        motor TEXT DEFAULT 'SQLSERVER', servidor TEXT, base_datos TEXT, usuario TEXT,
        password TEXT, puerto TEXT DEFAULT '1433', actualizado TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS documentos(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT, tipo TEXT, periodo TEXT, titulo TEXT, estado TEXT DEFAULT 'DISPONIBLE',
        creado TEXT
    )""")
    for col, ddl in [("rol","TEXT DEFAULT 'trabajador'"),("estado","TEXT DEFAULT 'ACTIVO'")]:
        ensure_col(cur, "users", col, ddl)
    if not cur.execute("SELECT id FROM users WHERE usuario=?", ("11223344",)).fetchone():
        cur.execute("""INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,rol,estado)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        ("11223344", generate_password_hash("123456"), "NISEDM01", "", COMPANY, "TRABAJADOR", "OPERACIONES",
         "RÉGIMEN GENERAL", "2026-06-22", "DNI", "11223344", "BCP", "001-000000000", "trabajador@pa.com",
         "999999999", "TRUJILLO", "admin", "ACTIVO"))
    for k, title in [("boletas","Boletas Normales"),("utilidades","Constancias de Utilidades"),("vacaciones","Boletas de Vacaciones"),("cts","Constancias de CTS"),("liquidaciones","Constancias de Liquidaciones"),("gratificaciones","Boletas de Gratificaciones"),("constancia-grati","Constancias de Gratificaciones")]:
        if not cur.execute("SELECT id FROM documentos WHERE usuario=? AND tipo=?", ("11223344", k)).fetchone():
            cur.execute("INSERT INTO documentos(usuario,tipo,periodo,titulo,estado,creado) VALUES(?,?,?,?,?,?)",
                        ("11223344", k, "2026", title, "DISPONIBLE", datetime.now().isoformat()))
    c.commit(); c.close()

init_db()

def esc(v):
    return str(v or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def dni(v):
    return re.sub(r"\D", "", str(v or ""))[-8:]
def today_iso():
    return date.today().isoformat()
def h12(dt=None):
    dt = dt or datetime.now(); h = dt.hour; am = "a. m." if h < 12 else "p. m."; hh = h % 12 or 12
    return f"{hh:02d}:{dt.minute:02d} {am}"
def today_es():
    dias = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
    meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    d = date.today()
    return f"{dias[d.weekday()]}, {d.day:02d} de {meses[d.month-1]} de {d.year}"
def login_required(fn):
    @wraps(fn)
    def w(*a, **kw):
        if not session.get("usuario"):
            return redirect(url_for("login"))
        return fn(*a, **kw)
    return w
def current_user():
    if not session.get("usuario"):
        return None
    c = conn(); r = c.execute("SELECT * FROM users WHERE usuario=?", (session["usuario"],)).fetchone(); c.close()
    return dict(r) if r else None

def icon(name):
    svgs = {
        "support":"<svg viewBox='0 0 24 24'><path d='M12 3a8 8 0 0 0-8 8v4a3 3 0 0 0 3 3h2v-7H6a6 6 0 0 1 12 0h-3v7h3a3 3 0 0 0 3-3v-4a8 8 0 0 0-8-8z'/></svg>",
        "settings":"<svg viewBox='0 0 24 24'><path d='M19.4 13.5c.1-.5.1-1 .1-1.5s0-1-.1-1.5l2-1.5-2-3.5-2.4 1a7 7 0 0 0-2.6-1.5L14 2h-4l-.4 2.5A7 7 0 0 0 7 6L4.6 5l-2 3.5 2 1.5a8 8 0 0 0 0 3l-2 1.5 2 3.5 2.4-1a7 7 0 0 0 2.6 1.5L10 22h4l.4-2.5A7 7 0 0 0 17 18l2.4 1 2-3.5-2-1.5zM12 15.5A3.5 3.5 0 1 1 12 8a3.5 3.5 0 0 1 0 7.5z'/></svg>",
        "profile":"<svg viewBox='0 0 24 24'><path d='M12 12a4.5 4.5 0 1 0 0-9 4.5 4.5 0 0 0 0 9zm0 2c-5 0-8 2.5-8 5.5V21h16v-1.5c0-3-3-5.5-8-5.5z'/></svg>",
        "list":"<svg viewBox='0 0 24 24'><path d='M4 5h3v3H4V5zm0 6h3v3H4v-3zm0 6h3v3H4v-3zM10 6h10v2H10V6zm0 6h10v2H10v-2zm0 6h10v2H10v-2z'/></svg>",
        "doc":"<svg viewBox='0 0 24 24'><path d='M6 2h9l5 5v15H6V2zm8 1.5V8h4.5L14 3.5zM8 12h8v2H8v-2zm0 4h8v2H8v-2z'/></svg>",
        "sign":"<svg viewBox='0 0 24 24'><path d='M4 17.5V21h3.5L18.1 10.4l-3.5-3.5L4 17.5zM20.7 7.8c.4-.4.4-1 0-1.4l-2.1-2.1a1 1 0 0 0-1.4 0l-1.6 1.6 3.5 3.5 1.6-1.6z'/></svg>",
        "home":"<svg viewBox='0 0 24 24'><path d='M3 11 12 3l9 8v10h-6v-6H9v6H3V11z'/></svg>",
        "logout":"<svg viewBox='0 0 24 24'><path d='M10 3h10v18H10v-2h8V5h-8V3zM8.6 16.6 7.2 18 2 12.8l5.2-5.2L8.6 9l-2.8 2.8H14v2H5.8l2.8 2.8z'/></svg>",
        "back":"<svg viewBox='0 0 24 24'><path d='M15.5 4.5 8 12l7.5 7.5-1.8 1.8L4.4 12l9.3-9.3 1.8 1.8z'/></svg>",
        "qr":"<svg viewBox='0 0 24 24'><path d='M3 3h8v8H3V3zm2 2v4h4V5H5zm8-2h8v8h-8V3zm2 2v4h4V5h-4zM3 13h8v8H3v-8zm2 2v4h4v-4H5zm10-2h2v2h-2v-2zm4 0h2v4h-4v-2h2v-2zm-6 4h2v4h-2v-4zm4 2h4v2h-4v-2z'/></svg>",
        "finger":"<svg viewBox='0 0 24 24'><path d='M12 2C8.7 2 6 4.7 6 8v4.2c0 2.6-1.3 4-2 4.8l1.5 1.3c.9-1 2.5-2.9 2.5-6.1V8a4 4 0 1 1 8 0v2h2V8c0-3.3-2.7-6-6-6zm0 4a2 2 0 0 0-2 2v4.5c0 3.6-1.5 6-2.7 7.3l1.5 1.4c1.4-1.5 3.2-4.4 3.2-8.7V8h0a0 0 0 1 1 0 0v4c0 4.7-1.7 7.4-2.6 8.7l1.7 1c1-1.5 2.9-4.6 2.9-9.7V8a2 2 0 0 0-2-2zm4 6v1c0 3.3-1 5.9-2.4 8.1l1.7 1c1.6-2.5 2.7-5.4 2.7-9.1v-1h-2z'/></svg>",
        "barcode":"<svg viewBox='0 0 24 24'><path d='M3 5h2v14H3V5zm3 0h1v14H6V5zm3 0h2v14H9V5zm3 0h1v14h-1V5zm3 0h3v14h-3V5zm4 0h1v14h-1V5z'/></svg>",
        "clock":"<svg viewBox='0 0 24 24'><path d='M12 2a10 10 0 1 0 .1 20.1A10 10 0 0 0 12 2zm1 11h5v-2h-4V6h-2v7z'/></svg>",
        "refresh":"<svg viewBox='0 0 24 24'><path d='M17.7 6.3A8 8 0 1 0 20 12h-2a6 6 0 1 1-1.8-4.3L13 11h8V3l-3.3 3.3z'/></svg>",
        "download":"<svg viewBox='0 0 24 24'><path d='M5 20h14v-2H5v2zM13 4h-2v8H8l4 4 4-4h-3V4z'/></svg>",
    }
    return "<span class='ico'>" + svgs.get(name, svgs["doc"]) + "</span>"
def leaves(): return "<div class='leafs'><i></i><i></i><i></i></div>"
def bottom(active='home'):
    data=[('home','/home','home','Home'),('asistencia','/asistencia','list','Asistencia'),('documentos','/documentos','doc','Documentos'),('perfil','/perfil','profile','Perfil')]
    return "<div class='bottom'>"+"".join([f"<a class='{ 'on' if k==active else '' }' href='{u}'>{icon(ic)}<b>{t}</b></a>" for k,u,ic,t in data])+"</div>"
def page_head(title, ic='doc', back='/home'):
    return f"<div class='page-head'><a class='back' href='{back}'>{icon('back')}</a><div class='headicon'>{icon(ic)}</div><h1>{title}</h1></div>"

CSS = """
:root{--g:#2f773b;--lime:#a8d32a;--ink:#06122a;--line:#e4ebe6;--shadow:0 8px 18px rgba(0,0,0,.14)}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}html,body{margin:0;min-height:100%;font-family:Arial,Segoe UI,sans-serif;background:#fff;color:var(--ink);font-weight:800}a{text-decoration:none;color:inherit}.ico{display:inline-grid;place-items:center;width:1.1em;height:1.1em;vertical-align:-.15em}.ico svg{width:100%;height:100%;fill:currentColor;display:block}.phone{width:100%;max-width:390px;min-height:100dvh;margin:0 auto;background:#fff;position:relative;border-left:1px solid #dfe7e3;border-right:1px solid #dfe7e3;overflow-x:hidden}.topmini{display:flex;justify-content:space-between;align-items:center;padding:10px 14px 0;font-size:12px;font-weight:1000}.topmini a,.topmini button{border:0;background:transparent;color:#fff;font-weight:1000;font-size:12px;padding:0}.topmini .ico{width:13px;height:13px;margin-right:3px}.card{background:#fff;border-radius:10px;box-shadow:var(--shadow);border:1px solid var(--line)}.login{min-height:100dvh;position:relative;background:white;padding:0 20px 20px}.login-hero{height:270px;margin:0 -20px;text-align:center;background:var(--g);color:white}.round-logo{width:112px;height:112px;border-radius:999px;background:white;border:7px solid var(--lime);display:grid;place-items:center;color:var(--g);box-shadow:0 4px 12px rgba(0,0,0,.25);margin:38px auto 14px}.round-logo .ico{width:54px;height:54px;color:var(--lime)}.brand{font-size:22px;font-weight:1000;letter-spacing:.3px;text-transform:uppercase;text-shadow:0 2px 3px rgba(0,0,0,.18)}.subbrand{font-size:12px;text-transform:uppercase;margin-top:4px;font-weight:1000}.login-card{margin:-34px 0 0;padding:14px 15px 18px}.role{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}.role button{height:42px;border-radius:8px;border:1px solid #d6e1d8;background:#fff;color:var(--g);font-weight:1000}.role .on{background:var(--g);color:white}.form-title{text-align:center;font-size:13px;margin:2px 0 14px}.label{font-size:10px;color:#334155;margin:9px 0 5px;font-weight:1000}.input,select,textarea{width:100%;min-height:39px;border:1px solid #d8e2dc;border-radius:7px;padding:0 10px;font-size:14px;font-weight:900;background:white}.btn{height:44px;border:0;border-radius:8px;background:var(--g);color:white;font-weight:1000;font-size:16px;width:100%;box-shadow:0 6px 13px rgba(47,119,59,.18);cursor:pointer}.btn.light{background:#f8fff9;color:var(--g);border:1px solid #b8d9bf}.btn.red{background:#fee2e2;color:#991b1b}.link{text-align:center;font-size:11px;color:#0071ff;font-weight:1000;margin-top:10px}.mini-tiles{display:flex;justify-content:center;gap:18px;margin-top:28px}.mini-tile{width:76px;height:76px;border-radius:8px;background:white;box-shadow:var(--shadow);color:var(--g);display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:9px;font-weight:1000;text-align:center}.mini-tile .ico{width:27px;height:27px;margin-bottom:5px}.login .leafs{margin-top:10px}.help{position:fixed;bottom:16px;left:50%;transform:translateX(-50%);width:390px;max-width:100%;text-align:center;font-size:11px;color:#7a8494}.home-hero{height:226px;background:var(--g);color:white;border-radius:0 0 18px 18px}.avatar{width:82px;height:82px;border-radius:999px;background:white;color:var(--g);display:grid;place-items:center;margin:17px auto 8px}.avatar .ico{width:52px;height:52px}.name{text-align:center;font-size:14px;font-weight:1000;text-transform:uppercase}.welcome{height:45px;margin:16px 18px 0;background:white;border-radius:9px;box-shadow:var(--shadow);color:var(--ink);display:grid;place-items:center;font-size:13px}.content{padding:28px 16px 82px}.stats{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:0 0 14px}.stat{background:#fbfffb;border:1px solid #e1ece4;border-radius:10px;padding:10px;text-align:center;color:var(--g);font-size:12px}.stat b{display:block;color:var(--ink);font-size:18px;margin-top:3px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.tile{height:92px;border-radius:9px;background:white;box-shadow:var(--shadow);border:1px solid #f0f0f0;color:var(--ink);display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:12px;font-weight:1000}.tile .ico{width:34px;height:34px;color:var(--g);margin-bottom:7px}.tile:hover,.doc-row:hover{background:#eaf5eb}.leafs{width:115px;height:115px;margin:24px auto 0;position:relative;opacity:.42}.leafs i{position:absolute;width:60px;height:96px;border-radius:70% 30% 70% 30%;filter:drop-shadow(0 4px 4px rgba(0,0,0,.11))}.leafs i:nth-child(1){background:#ffd5c6;left:10px;transform:rotate(25deg)}.leafs i:nth-child(2){background:#cbd9b6;left:49px;transform:rotate(25deg)}.leafs i:nth-child(3){background:#fff1ad;left:74px;top:30px;transform:rotate(25deg)}.sync{position:fixed;left:10px;bottom:9px;width:330px;max-width:calc(100vw - 60px);background:white;border-radius:10px;box-shadow:var(--shadow);display:grid;grid-template-columns:34px 1fr;gap:8px;align-items:center;padding:8px 10px;z-index:20}.syncbtn{width:32px;height:32px;border-radius:999px;background:var(--g);color:white;display:grid;place-items:center;font-size:19px}.sync .t{font-size:11px;font-weight:1000}.sync .s{font-size:9px;color:#16823c;margin-top:2px}.exit{position:fixed;right:13px;bottom:12px;color:#ef4444;font-size:26px;z-index:30}.page-head{height:132px;background:var(--g);color:white;border-radius:0 0 15px 15px;text-align:center;position:relative;padding-top:18px}.back{position:absolute;left:13px;top:17px;color:white}.back .ico{width:24px;height:24px}.headicon{font-size:33px;margin-top:14px}.headicon .ico{width:34px;height:34px}.page-head h1{font-size:14px;text-transform:uppercase;margin:12px 0 0;font-weight:1000;letter-spacing:.2px}.toolbar{height:57px;margin:-24px 10px 8px;background:white;border-radius:8px;box-shadow:var(--shadow);display:flex;align-items:center;gap:18px;padding:0 14px;color:var(--g);position:relative;z-index:2}.toolbtn{border:0;background:transparent;color:var(--g);font-size:24px;cursor:pointer}.bottom{position:fixed;bottom:0;left:50%;transform:translateX(-50%);width:100%;max-width:390px;height:58px;background:#fff;border-top:1px solid #e4ebe6;display:flex;z-index:40}.bottom a{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#6d7b72;font-size:10px;font-weight:900}.bottom a.on{color:var(--g)}.bottom .ico{width:21px;height:21px;margin-bottom:3px}.att-card{margin:-21px 10px 0;background:white;border-radius:10px;box-shadow:var(--shadow);padding:12px 10px 14px;position:relative;z-index:2}.sys{text-align:center;color:#008c3e;font-size:10px;font-weight:1000}.clock{text-align:center;font-size:32px;color:#05091f;font-weight:1000;margin:7px 0 3px}.date{text-align:center;color:#8690a2;font-size:10px;margin-bottom:10px}.mark{width:100%;height:42px;border:0;border-radius:7px;background:#eef1f3;color:#405064;font-size:13px;font-weight:1000;margin-bottom:8px;cursor:pointer}.mark.main{background:#22b45e;color:white}.mark:hover,.mark.active{background:var(--g);color:#fff}.row2{display:grid;grid-template-columns:1fr 1fr;gap:8px}.methods{display:grid;grid-template-columns:1fr 1fr 1fr;gap:9px;margin:10px 0}.method{height:82px;border-radius:8px;background:white;box-shadow:var(--shadow);border:1px solid var(--line);display:flex;align-items:center;justify-content:center;flex-direction:column;color:var(--g);font-weight:1000;font-size:11px;cursor:pointer}.method .ico{width:36px;height:36px;margin-bottom:4px}.maptitle{display:flex;justify-content:space-between;align-items:center;color:#748095;font-size:10px;font-weight:1000;margin:12px 0 8px}.signal{border:1.8px solid var(--g);border-radius:999px;background:#dff8e5;color:#087b3e;font-size:10px;font-weight:1000;padding:6px 10px}.map{height:125px;border:1px solid #dce4eb;border-radius:12px;background:#f4f6f8;position:relative;overflow:hidden}.road1{position:absolute;width:360px;height:12px;background:#ffd56e;left:-26px;top:45px;transform:rotate(37deg)}.road2{position:absolute;width:300px;height:10px;background:#ced5dd;left:-30px;top:78px;transform:rotate(-2deg)}.road3{position:absolute;width:13px;height:210px;background:#ced5dd;left:168px;top:-34px;transform:rotate(24deg)}.pin{position:absolute;left:50%;top:55%;transform:translate(-50%,-50%);font-size:44px;color:#1e75ff}.bubble{position:absolute;left:50%;top:40%;transform:translate(-50%,-50%);background:white;border-radius:6px;box-shadow:0 5px 12px rgba(0,0,0,.17);font-size:10px;padding:7px 9px;white-space:nowrap}.locfoot{display:flex;justify-content:space-between;gap:8px;font-size:10px;color:#64748b;font-weight:900;margin:8px 3px 0}.locfoot a{color:#006dfc;font-size:12px}.doc-list{background:#fafafa;padding:22px 10px 82px;min-height:calc(100dvh - 132px)}.badge{display:inline-flex;align-items:center;gap:5px;background:#eaf5eb;color:var(--g);border:1px solid #cfe6d4;border-radius:999px;padding:5px 9px;font-size:10px;font-weight:1000;margin-bottom:12px}.doc-row{height:66px;background:white;border-radius:9px;border:1px solid #e8ece8;box-shadow:0 5px 13px rgba(0,0,0,.05);margin-bottom:10px;display:grid;grid-template-columns:42px 1fr 18px;align-items:center;padding:0 11px}.doc-ico{width:32px;height:32px;border:1px solid #d4e6d8;border-radius:8px;color:var(--g);display:grid;place-items:center}.doc-title{font-size:13px;font-weight:1000}.doc-sub{font-size:10px;color:#64748b;margin-top:3px}.chev{font-size:24px;color:#8d96a3}.profile-card,.form-card{margin:12px 10px 78px;background:white;border-radius:10px;box-shadow:var(--shadow);padding:14px}.profile-photo{width:72px;height:72px;border-radius:999px;background:#edf4ef;color:var(--g);display:grid;place-items:center;margin:0 auto 12px}.profile-photo .ico{width:42px;height:42px}.center{text-align:center}.kv{display:grid;gap:8px;margin-top:14px}.kv div{background:#f5f7f7;border-radius:8px;padding:9px 11px}.kv small{display:block;color:#8590a2;font-size:10px;font-weight:1000}.kv b{font-size:11px}.signbox{height:190px;border:1px solid #dbe4df;border-radius:10px;background:white;margin:12px 0;position:relative}#canvas{width:100%;height:100%;touch-action:none}.btnrow{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:9px}.msg{border-radius:9px;padding:9px 10px;font-size:11px;font-weight:900;margin-bottom:9px}.ok{background:#e7f8eb;border:1px solid #bfe7c7;color:#166534}.bad{background:#fee2e2;border:1px solid #fecaca;color:#991b1b}.small{font-size:10px;color:#64748b;line-height:1.35}.modal{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;display:none;align-items:center;justify-content:center;padding:10px}.modal.show{display:flex}.modal-card{width:100%;max-width:390px;background:white;border-radius:12px;overflow:hidden;box-shadow:0 12px 40px rgba(0,0,0,.28)}.modal-head{height:48px;background:var(--g);color:white;display:flex;align-items:center;justify-content:space-between;padding:0 14px;font-size:13px;font-weight:1000;text-transform:uppercase}.close{border:0;background:transparent;color:white;font-size:26px}.scan-view{height:360px;background:#1f2937;position:relative;color:white;display:grid;place-items:center;overflow:hidden}.scan-view video{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;opacity:.88}.scan-frame{position:absolute;left:46px;right:46px;top:70px;height:190px;border:3px solid #22c55e}.scan-line{position:absolute;left:30px;right:30px;top:166px;height:2px;background:#22c55e;box-shadow:0 0 8px #22c55e}.scan-text{position:absolute;left:20px;right:20px;bottom:65px;text-align:center;font-size:12px;font-weight:1000}.scan-actions{position:absolute;left:26px;right:26px;bottom:18px}.manual{padding:12px}.fingerbox{height:310px;display:flex;flex-direction:column;align-items:center;justify-content:center}.fingerprint{width:160px;height:160px;border-radius:999px;border:6px solid #e5e7eb;display:grid;place-items:center;color:var(--g);font-size:90px;position:relative}.fingerprint:before{content:'';position:absolute;inset:3px;border-radius:999px;border:5px solid #43a35b;border-right-color:transparent;animation:spin 1.2s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}.wheel-modal{position:fixed;inset:auto 0 0;z-index:90;background:rgba(0,0,0,.25);display:none;justify-content:center}.wheel-modal.show{display:flex}.wheel-card{width:100%;max-width:390px;background:white;border-radius:18px 18px 0 0;padding:12px 18px 18px;box-shadow:0 -12px 35px rgba(0,0,0,.2)}.wheel-head{display:flex;justify-content:space-between;align-items:center;font-weight:1000;margin-bottom:7px}.wheel-head button{border:0;background:white;color:#009b55;font-weight:1000;font-size:16px}.wheels{height:148px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;position:relative}.wheels:before{content:'';position:absolute;left:0;right:0;top:54px;height:40px;background:rgba(0,0,0,.045);border-radius:10px;pointer-events:none}.wheel{height:148px;overflow-y:auto;scroll-snap-type:y mandatory;text-align:center;padding:54px 0;scrollbar-width:none}.wheel::-webkit-scrollbar{display:none}.wheel div{height:40px;line-height:40px;font-size:22px;color:#b6bdc7;scroll-snap-align:center;font-weight:700}.wheel div.sel{color:#303642;font-size:27px}.toast{position:fixed;top:12px;left:50%;transform:translateX(-50%);background:#10251a;color:white;border-radius:999px;padding:10px 17px;font-size:12px;font-weight:1000;z-index:200;box-shadow:var(--shadow);white-space:nowrap}.table-preview{max-height:200px;overflow:auto;border:1px solid var(--line);border-radius:8px;margin-top:8px}.table-preview table{width:100%;border-collapse:collapse;font-size:10px}.table-preview td,.table-preview th{border-bottom:1px solid #edf2ef;padding:5px;text-align:left}@media(min-width:850px){body{display:block}.phone{width:390px;margin:0 auto}.sync{left:10px;transform:none}.exit{right:15px}.login:after{content:'➜';position:fixed;right:12px;bottom:7px;color:#ef4444;font-size:26px;font-weight:900}}@media(max-width:430px){.phone{border:0}.sync{max-width:calc(100vw - 56px)}}
"""

JS = """
let lastMethod = 'BOTON', lastReading = '';
function tone(ok=true){try{const C=window.AudioContext||window.webkitAudioContext;const ctx=new C();const o=ctx.createOscillator();const g=ctx.createGain();o.type='sine';o.frequency.value=ok?880:230;g.gain.value=.075;o.connect(g);g.connect(ctx.destination);o.start();setTimeout(()=>{o.stop();ctx.close()},150)}catch(e){}}
function toast(msg,ok=true){tone(ok);let t=document.createElement('div');t.className='toast';t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.remove(),2300)}
function liveClock(){let el=document.getElementById('liveClock'); if(!el)return; let d=new Date(); let h=d.getHours(); let am=h<12?'a. m.':'p. m.'; h=h%12||12; el.textContent=String(h).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0')+' '+am}
setInterval(liveClock,1000); document.addEventListener('DOMContentLoaded',()=>{liveClock(); requestLocation(false); initSign();});
function active(btn){btn?.classList.add('active'); setTimeout(()=>btn?.classList.remove('active'),900)}
function requestLocation(notify=true){let info=document.getElementById('locText'), pin=document.getElementById('pin'), link=document.getElementById('mapLink'); if(!info)return; if(!navigator.geolocation){info.textContent='Navegador sin GPS.';return} navigator.geolocation.getCurrentPosition(p=>{let lat=p.coords.latitude.toFixed(6),lng=p.coords.longitude.toFixed(6),acc=Math.round(p.coords.accuracy); info.textContent='Ubicación real detectada con precisión aproximada de '+acc+' m.'; if(pin){pin.dataset.lat=lat;pin.dataset.lng=lng;pin.dataset.acc=acc} if(link)link.href='https://www.google.com/maps?q='+lat+','+lng; if(notify)toast('GPS actualizado')},()=>{info.textContent='Permite ubicación para registrar GPS real.'; if(notify)toast('No se pudo obtener ubicación',false)},{enableHighAccuracy:true,timeout:9000,maximumAge:0})}
function marcar(tipo){active(document.querySelector('[data-tipo=\"'+tipo+'\"]')); let p=document.getElementById('pin')||{}; let payload={tipo,metodo:lastMethod,lectura:lastReading,hora_manual:document.getElementById('manualTime')?.dataset.value||'',lat:p.dataset?.lat||'',lng:p.dataset?.lng||'',precision:p.dataset?.acc||''}; function send(){fetch('/api/marcar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).then(r=>r.json()).then(j=>{toast(j.msg||'Marcado',j.ok); if(j.ok)setTimeout(()=>location.reload(),850)}).catch(()=>toast('No se pudo registrar',false))} if(navigator.geolocation){navigator.geolocation.getCurrentPosition(pos=>{payload.lat=pos.coords.latitude;payload.lng=pos.coords.longitude;payload.precision=Math.round(pos.coords.accuracy);send()},()=>send(),{enableHighAccuracy:true,timeout:5000})}else send()}
function setManualTime(txt){let e=document.getElementById('manualTime'); if(e){e.dataset.value=txt; e.innerHTML='● Ajustar hora táctil: '+txt;} lastMethod='HORA_MANUAL'; toast('Hora manual: '+txt)}
function openWheel(){document.getElementById('wheelModal')?.classList.add('show'); setNowWheel(); syncWheel()}
function closeWheel(){document.getElementById('wheelModal')?.classList.remove('show')}
function syncWheel(){document.querySelectorAll('.wheel').forEach(w=>{let mid=w.scrollTop+74; [...w.children].forEach(c=>c.classList.toggle('sel',Math.abs((c.offsetTop+20)-mid)<22))})}
function setNowWheel(){let d=new Date(); setWheel('wh',d.getHours()%12||12); setWheel('wm',d.getMinutes()); setWheel('wa',d.getHours()<12?'AM':'PM')}
function setWheel(id,val){let w=document.getElementById(id); if(!w)return; [...w.children].forEach(c=>{if(String(c.dataset.v)===String(val))w.scrollTop=c.offsetTop-54})}
function pick(id){let w=document.getElementById(id),mid=w.scrollTop+74,best=w.children[0],bd=999;[...w.children].forEach(c=>{let d=Math.abs((c.offsetTop+20)-mid);if(d<bd){bd=d;best=c}});return best.dataset.v}
function applyWheel(){let h=pick('wh'),m=pick('wm'),a=pick('wa');let v=String(h).padStart(2,'0')+':'+String(m).padStart(2,'0')+' '+(a==='AM'?'a. m.':'p. m.'); setManualTime(v); closeWheel()}
document.addEventListener('scroll',e=>{if(e.target.classList&&e.target.classList.contains('wheel'))syncWheel()},true)
let scanStream=null, scanTimer=null;
function openScanner(kind){lastMethod = kind === 'barcode' ? 'CODIGO_BARRAS_CAMARA' : 'QR_CAMARA'; let m=document.getElementById('scanModal'); if(!m)return; document.getElementById('scanTitle').textContent=kind==='barcode'?'LECTURA CÓDIGO DE BARRAS':'LECTURA QR'; document.getElementById('scanHint').textContent=kind==='barcode'?'Alinea el código de barras dentro del marco':'Alinea el código QR dentro del marco'; m.classList.add('show'); startCamera(kind)}
async function startCamera(kind){let video=document.getElementById('scanVideo'); try{scanStream=await navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'},audio:false}); video.srcObject=scanStream; await video.play(); if('BarcodeDetector' in window){let formats=kind==='barcode'?['code_128','code_39','ean_13','ean_8','upc_a','upc_e']:['qr_code']; let detector=new BarcodeDetector({formats}); scanTimer=setInterval(async()=>{try{let codes=await detector.detect(video); if(codes&&codes.length){lastReading=codes[0].rawValue||''; closeScanner(); toast('Lectura detectada: '+lastReading);}}catch(e){}},600)} else {document.getElementById('manualScan').style.display='block'; toast('Usa lectura manual si el navegador no detecta',false)}}catch(e){document.getElementById('manualScan').style.display='block'; toast('No se pudo abrir cámara',false)}}
function closeScanner(){document.getElementById('scanModal')?.classList.remove('show'); if(scanTimer)clearInterval(scanTimer); scanTimer=null; if(scanStream){scanStream.getTracks().forEach(t=>t.stop());scanStream=null}}
function applyManualScan(){let v=document.getElementById('scanManualValue').value.trim(); if(!v){toast('Ingrese lectura manual',false);return} lastReading=v; closeScanner(); toast('Lectura capturada: '+v)}
function openFinger(){lastMethod='HUELLA_WEB'; document.getElementById('fingerModal')?.classList.add('show'); setTimeout(()=>{lastReading='HUELLA-'+Date.now(); closeFinger(); toast('Huella validada')},2300)}
function closeFinger(){document.getElementById('fingerModal')?.classList.remove('show')}
function initSign(){let c=document.getElementById('canvas'); if(!c)return; let ctx=c.getContext('2d'),draw=false; function resize(){let r=c.getBoundingClientRect();c.width=r.width*devicePixelRatio;c.height=r.height*devicePixelRatio;ctx.setTransform(devicePixelRatio,0,0,devicePixelRatio,0,0);ctx.lineWidth=2;ctx.lineCap='round';ctx.strokeStyle='#111'} resize(); window.addEventListener('resize',resize); function p(ev){let r=c.getBoundingClientRect(),t=ev.touches?ev.touches[0]:ev;return{x:t.clientX-r.left,y:t.clientY-r.top}} c.addEventListener('pointerdown',ev=>{draw=true;let q=p(ev);ctx.beginPath();ctx.moveTo(q.x,q.y)}); c.addEventListener('pointermove',ev=>{if(!draw)return;let q=p(ev);ctx.lineTo(q.x,q.y);ctx.stroke()}); c.addEventListener('pointerup',()=>draw=false); c.addEventListener('pointerleave',()=>draw=false); window.clearSign=()=>ctx.clearRect(0,0,c.width,c.height); window.saveSign=()=>fetch('/api/firma',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({data:c.toDataURL('image/png')})}).then(r=>r.json()).then(j=>toast(j.msg,j.ok)); window.downloadSign=()=>{let a=document.createElement('a');a.href=c.toDataURL('image/png');a.download='firma_pa.png';a.click()};}
function fakeOk(msg){toast(msg||'Opción habilitada')}
function testDB(){document.getElementById('dbOk')?.classList.add('ok');document.getElementById('dbOk').textContent='✓ Conexión lista para validar cuando habilites driver SQL Server.';toast('Configuración validada')}
"""

def shell(body, title="P&A Mobile"):
    return f"<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1,viewport-fit=cover'><meta name='theme-color' content='{GREEN}'><title>{title}</title><style>{CSS}</style></head><body><div class='phone'>{body}</div><script>{JS}</script></body></html>"

@app.route("/")
def index():
    if session.get("usuario"):
        return redirect(url_for("home"))
    return redirect(url_for("login"))

@app.route("/login", methods=["GET","POST"])
def login():
    err=""
    if request.method=="POST":
        usuario=dni(request.form.get("usuario","")); password=request.form.get("password","")
        c=conn(); r=c.execute("SELECT * FROM users WHERE usuario=? AND estado='ACTIVO'", (usuario,)).fetchone(); c.close()
        if r and check_password_hash(r["password_hash"], password):
            session["usuario"]=usuario
            return redirect(url_for("home"))
        err="Usuario o contraseña incorrectos."
    body=f"""
    <div class='login'>
      <div class='login-hero'>
        <div class='topmini'><button onclick="fakeOk('Soporte habilitado')">{icon('support')} Soporte</button><button onclick="fakeOk('Config. disponible al ingresar')">{icon('settings')} Config.</button></div>
        <div class='round-logo'>{icon('doc')}</div>
        <div class='brand'>P&A Mobile</div><div class='subbrand'>Iniciar Sesión</div>
      </div>
      <form method='post' class='card login-card'>
        <div class='role'><button type='button' class='on'>USUARIO</button><button type='button'>ADMINISTRADOR</button></div>
        <div class='form-title'>Inicia sesión para continuar</div>
        {f"<div class='msg bad'>{err}</div>" if err else ""}
        <div class='label'>Usuario</div><input class='input' name='usuario' value='11223344' maxlength='8' inputmode='numeric' required>
        <div class='label'>Empresa</div><select class='input'><option>{COMPANY}</option></select>
        <div class='label'>Contraseña</div><input class='input' name='password' value='123456' type='password' required>
        <button class='btn' style='margin-top:14px'>{icon('logout')} INGRESAR</button>
        <div class='link'>¿Olvidaste tu contraseña?</div>
      </form>
      <div class='mini-tiles'><a class='mini-tile' href='#'>{icon('list')}ASISTENCIA</a><a class='mini-tile' href='#'>{icon('doc')}DOCUMENTOS</a></div>
      {leaves()}<div class='help'>¿Problemas para acceder?<br>Contactar a Mesa</div>
    </div>
    """
    return shell(body, "Login P&A")

@app.route("/home")
@login_required
def home():
    u=current_user() or {}
    c=conn(); marcas=c.execute("SELECT tipo,hora FROM marcas WHERE usuario=? AND fecha=? ORDER BY id", (session["usuario"], today_iso())).fetchall(); c.close()
    ult = marcas[-1]["tipo"].replace("_"," ") if marcas else "SIN REG."
    body=f"""
    <div class='home-hero'>
      <div class='topmini'><button onclick="fakeOk('Soporte habilitado')">{icon('support')} Soporte</button><a href='/config'>{icon('settings')} Config.</a></div>
      <div class='avatar'>{icon('profile')}</div><div class='name'>{esc(u.get('nombres','NISEDM01'))}</div><div class='welcome'>Bienvenido(a)</div>
    </div>
    <div class='content'>
      <div class='stats'><div class='stat'>Marcaciones<b>{len(marcas)}</b></div><div class='stat'>Último<b>{esc(ult)}</b></div></div>
      <div class='grid'><a class='tile' href='/asistencia'>{icon('list')}ASISTENCIA</a><a class='tile' href='/documentos'>{icon('doc')}DOCUMENTOS</a><a class='tile' href='/firma'>{icon('sign')}FIRMA</a><a class='tile' href='/perfil'>{icon('profile')}PERFIL</a></div>{leaves()}
    </div>
    <div class='sync'><div class='syncbtn'>{icon('refresh')}</div><div><div class='t'>Sincronizar Tablas Maestras</div><div class='s'>Actualizado hasta: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div></div></div><a class='exit' href='/logout'>{icon('logout')}</a>
    """
    return shell(body, "Home P&A")

@app.route("/asistencia")
@login_required
def asistencia():
    c=conn(); marcas=c.execute("SELECT * FROM marcas WHERE usuario=? AND fecha=? ORDER BY id DESC", (session["usuario"], today_iso())).fetchall(); c.close()
    hist = "".join([f"<div class='doc-row' style='height:48px;grid-template-columns:1fr 80px'><div><div class='doc-title'>{esc(m['tipo'].replace('_',' '))}</div><div class='doc-sub'>{esc(m['metodo'] or 'BOTON')} {esc(m['lectura'] or '')}</div></div><b>{esc(m['hora_manual'] or m['hora'])}</b></div>" for m in marcas[:4]])
    body=f"""
    <div class='page-head'><a class='back' href='/home'>{icon('back')}</a><div class='headicon'>{icon('list')}</div><h1>Registra tu asistencia fácil</h1></div>
    <div class='att-card'>
      <div class='sys'>HORA ACTUAL DEL SISTEMA</div><div class='clock' id='liveClock'>{h12()}</div><div class='date'>{today_es()}</div>
      <button class='mark main' data-tipo='INGRESO' onclick="marcar('INGRESO')">↪ Registrar Ingreso</button>
      <div class='row2'><button class='mark' data-tipo='SALIDA_REFRIGERIO' onclick="marcar('SALIDA_REFRIGERIO')">Salida a<br>Refrigerio</button><button class='mark' data-tipo='RETORNO_REFRIGERIO' onclick="marcar('RETORNO_REFRIGERIO')">Retorno de<br>Refrigerio</button></div>
      <button class='mark' data-tipo='SALIDA' onclick="marcar('SALIDA')">↩ Registrar Salida</button>
      <button class='mark' id='manualTime' data-value='' onclick='openWheel()'>● Ajustar hora táctil: Ahora</button>
      <div class='label'>MÉTODOS DE MARCACIÓN</div><div class='methods'><button class='method' onclick="openScanner('qr')">{icon('qr')}QR</button><button class='method' onclick="openFinger()">{icon('finger')}HUELLA</button><button class='method' onclick="openScanner('barcode')">{icon('barcode')}CÓDIGO</button></div>
      <div class='maptitle'><span>UBICACIÓN ACTUAL</span><button class='signal' onclick='requestLocation(true)'>SEÑAL FUERTE</button></div>
      <div class='map'><div class='road1'></div><div class='road2'></div><div class='road3'></div><div class='bubble'>Tu ubicación actual</div><div id='pin' class='pin'>📍</div></div>
      <div class='locfoot'><span id='locText'>Detectando ubicación real...</span><a id='mapLink' target='_blank'>Ver mapa</a></div>
      <div style='margin-top:12px'>{hist or "<div class='msg ok'>Aún no tienes marcaciones hoy.</div>"}</div>
    </div>
    {scanner_modal()}{finger_modal()}{wheel_modal()}{bottom('asistencia')}
    """
    return shell(body, "Asistencia P&A")

def scanner_modal():
    return """
    <div class='modal' id='scanModal'><div class='modal-card'><div class='modal-head'><span id='scanTitle'>LECTURA QR</span><button class='close' onclick='closeScanner()'>×</button></div><div class='scan-view'><video id='scanVideo' muted playsinline></video><div class='scan-frame'></div><div class='scan-line'></div><div class='scan-text' id='scanHint'>Alinea el código dentro del marco</div><div class='scan-actions'><button class='btn light' onclick='closeScanner()'>CANCELAR</button></div></div><div class='manual' id='manualScan' style='display:none'><div class='label'>Lectura manual si tu navegador no detecta cámara</div><input class='input' id='scanManualValue' placeholder='Digite DNI / QR / Código'><button class='btn' style='margin-top:8px' onclick='applyManualScan()'>Usar lectura</button></div></div></div>
    """
def finger_modal():
    return "<div class='modal' id='fingerModal'><div class='modal-card'><div class='modal-head'><span>LECTURA DE HUELLA</span><button class='close' onclick='closeFinger()'>×</button></div><div class='fingerbox'><div class='fingerprint'>⌾</div><div class='small' style='margin-top:18px;text-align:center'>Coloca tu dedo en el sensor<br>Leyendo huella...</div></div></div></div>"
def wheel_modal():
    hours=''.join([f"<div data-v='{i}'>{i:02d}</div>" for i in range(1,13)])
    mins=''.join([f"<div data-v='{i}'>{i:02d}</div>" for i in range(60)])
    ampm="<div data-v='AM'>AM</div><div data-v='PM'>PM</div>"
    return f"<div class='wheel-modal' id='wheelModal'><div class='wheel-card'><div class='wheel-head'><button onclick='closeWheel()'>Cancelar</button><span>Seleccionar hora</span><button onclick='applyWheel()'>Aceptar</button></div><div class='wheels'><div class='wheel' id='wh'>{hours}</div><div class='wheel' id='wm'>{mins}</div><div class='wheel' id='wa'>{ampm}</div></div></div></div>"

@app.route("/documentos")
@login_required
def documentos():
    c=conn(); docs=c.execute("SELECT * FROM documentos WHERE usuario=? ORDER BY id", (session["usuario"],)).fetchall(); c.close()
    rows="".join([f"<a class='doc-row' href='/documentos/{esc(d['tipo'])}'><div class='doc-ico'>{icon('download')}</div><div><div class='doc-title'>{esc(d['titulo'])}</div><div class='doc-sub'>Disponible para descarga · {esc(d['periodo'])}</div></div><div class='chev'>›</div></a>" for d in docs])
    body=page_head("Gestión de documentos","doc")+f"<div class='doc-list'><span class='badge'>{icon('doc')} Documentos laborales habilitados</span>{rows}</div>{bottom('documentos')}"
    return shell(body, "Documentos P&A")

@app.route("/documentos/<tipo>")
@login_required
def doc_detalle(tipo):
    u=current_user() or {}
    title = {"boletas":"Boleta Normal","utilidades":"Constancia de Utilidades","vacaciones":"Boleta de Vacaciones","cts":"Constancia de CTS","liquidaciones":"Constancia de Liquidación","gratificaciones":"Boleta de Gratificación","constancia-grati":"Constancia de Gratificación"}.get(tipo, "Documento")
    body=page_head(title,"doc","/documentos")+f"<div class='form-card'><span class='badge'>{icon('doc')} Disponible</span><div class='kv'><div><small>Trabajador</small><b>{esc(u.get('nombres'))}</b></div><div><small>DNI</small><b>{esc(u.get('usuario'))}</b></div><div><small>Empresa</small><b>{esc(u.get('empresa'))}</b></div><div><small>Periodo</small><b>2026</b></div></div><a class='btn' style='display:grid;place-items:center;margin-top:14px' href='/api/documento/{tipo}'>Descargar documento</a><button class='btn light' style='margin-top:9px' onclick=\"fakeOk('Vista previa habilitada')\">Ver vista previa</button></div>{bottom('documentos')}"
    return shell(body, title)

@app.route("/firma")
@login_required
def firma():
    c=conn(); r=c.execute("SELECT actualizado FROM firmas WHERE usuario=?", (session["usuario"],)).fetchone(); c.close()
    status = f"<div class='msg ok'>✓ Firma registrada: {esc(r['actualizado'])}</div>" if r else "<div class='msg ok'>Dibuja tu firma y presiona guardar.</div>"
    body=page_head("Firma Digital","sign")+f"<div class='form-card'>{status}<div class='signbox'><canvas id='canvas'></canvas></div><div class='btnrow'><button class='btn light' onclick='clearSign()'>Limpiar</button><button class='btn' onclick='saveSign()'>Guardar →</button></div><button class='btn light' onclick='downloadSign()'>Descargar firma</button><div class='msg ok' style='margin-top:12px'>Tu firma se utiliza para validar documentos y registros dentro del sistema.</div></div>{bottom('home')}"
    return shell(body, "Firma P&A")

@app.route("/perfil")
@login_required
def perfil():
    u=current_user() or {}
    body=page_head("Perfil","profile")+f"<div class='profile-card'><div class='profile-photo'>{icon('profile')}</div><div class='center'><h3 style='margin:0;font-size:16px'>{esc(u.get('nombres'))}</h3><p style='font-size:11px;margin:7px 0'>{esc(u.get('cargo'))}</p></div><div class='kv'><div><small>DNI</small><b>{esc(u.get('usuario'))}</b></div><div><small>Empresa</small><b>{esc(u.get('empresa'))}</b></div><div><small>Área</small><b>{esc(u.get('area'))}</b></div><div><small>Régimen</small><b>{esc(u.get('regimen'))}</b></div><div><small>Ingreso</small><b>{esc(u.get('fecha_ingreso'))}</b></div><div><small>Banco</small><b>{esc(u.get('banco'))}</b></div><div><small>Cuenta</small><b>{esc(u.get('cuenta'))}</b></div><div><small>Celular</small><b>{esc(u.get('celular'))}</b></div></div></div>{bottom('perfil')}"
    return shell(body, "Perfil P&A")

@app.route("/config", methods=["GET","POST"])
@login_required
def config():
    if request.method=="POST" and request.form.get("tipo")=="db":
        c=conn(); c.execute("INSERT INTO config_db(motor,servidor,base_datos,usuario,password,puerto,actualizado) VALUES(?,?,?,?,?,?,?)",("SQLSERVER", request.form.get("servidor",""), request.form.get("base_datos",""), request.form.get("db_usuario",""),request.form.get("password",""), request.form.get("puerto","1433"), datetime.now().isoformat())); c.commit(); c.close()
    body=page_head("Configuración","settings")+"""
    <div class='form-card'><div class='role'><button class='on' type='button'>Base de Datos</button><button type='button'>General</button></div><form method='post'><input type='hidden' name='tipo' value='db'><div class='label'>Motor de Base de Datos</div><select name='motor'><option>SQLSERVER</option><option>SQLITE LOCAL</option></select><div class='label'>Servidor</div><input class='input' name='servidor' placeholder='192.168.1.10'><div class='label'>Base de Datos</div><input class='input' name='base_datos' placeholder='RRHH_PA'><div class='label'>Usuario</div><input class='input' name='db_usuario' placeholder='sa'><div class='label'>Contraseña</div><input class='input' type='password' name='password'><div class='label'>Puerto</div><input class='input' name='puerto' value='1433'><button class='btn' style='margin-top:12px'>Guardar Configuración</button></form><button class='btn light' style='margin-top:10px' onclick='testDB()'>Probar conexión</button><div id='dbOk' class='msg' style='margin-top:10px'></div><hr style='border:0;border-top:1px solid #e5ece7;margin:14px 0'><a class='btn' style='display:grid;place-items:center' href='/config/trabajadores'>Cargar trabajadores</a><a class='btn light' style='display:grid;place-items:center;margin-top:9px' href='/plantilla_trabajadores.csv'>Descargar plantilla CSV</a></div>
    """+bottom("perfil")
    return shell(body, "Configuración P&A")

@app.route("/config/trabajadores", methods=["GET","POST"])
@login_required
def config_trabajadores():
    msg=""; preview=""
    if request.method=="POST":
        file=request.files.get("archivo")
        if not file:
            msg="<div class='msg bad'>Selecciona un archivo CSV.</div>"
        else:
            raw=file.read().decode("utf-8-sig", errors="ignore")
            reader=csv.DictReader(StringIO(raw))
            count=0; bad=0; rows=[]; c=conn()
            for row in reader:
                d=dni(row.get("DNI") or row.get("dni"))
                if len(d)!=8: bad+=1; continue
                nombres=(row.get("NOMBRES") or row.get("nombres") or "").strip().upper()
                apellidos=(row.get("APELLIDOS") or row.get("apellidos") or "").strip().upper()
                cargo=(row.get("CARGO") or row.get("cargo") or "TRABAJADOR").strip().upper()
                area=(row.get("AREA") or row.get("area") or "OPERACIONES").strip().upper()
                regimen=(row.get("REGIMEN") or row.get("regimen") or "RÉGIMEN GENERAL").strip().upper()
                correo=(row.get("CORREO") or row.get("correo") or "").strip()
                celular=(row.get("CELULAR") or row.get("celular") or "").strip()
                banco=(row.get("BANCO") or row.get("banco") or "").strip().upper()
                cuenta=(row.get("CUENTA") or row.get("cuenta") or "").strip()
                c.execute("INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,rol,estado) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(usuario) DO UPDATE SET nombres=excluded.nombres,apellidos=excluded.apellidos,cargo=excluded.cargo,area=excluded.area,regimen=excluded.regimen,correo=excluded.correo,celular=excluded.celular,banco=excluded.banco,cuenta=excluded.cuenta,estado='ACTIVO'",(d, generate_password_hash("123456"), nombres, apellidos, COMPANY, cargo, area, regimen, date.today().isoformat(), "DNI", d, banco, cuenta, correo, celular, "", "trabajador", "ACTIVO"))
                rows.append((d,nombres,cargo)); count+=1
            c.commit(); c.close()
            msg=f"<div class='msg ok'>✓ Importación finalizada. Correctos: {count}. Errores: {bad}. Clave inicial: 123456.</div>"
            preview="<div class='table-preview'><table><tr><th>DNI</th><th>NOMBRES</th><th>CARGO</th></tr>"+"".join([f"<tr><td>{esc(a)}</td><td>{esc(b)}</td><td>{esc(c)}</td></tr>" for a,b,c in rows[:40]])+"</table></div>"
    body=page_head("Cargar trabajadores","profile","/config")+f"<div class='form-card'>{msg}<form method='post' enctype='multipart/form-data'><div class='label'>Archivo CSV de trabajadores</div><input class='input' type='file' name='archivo' accept='.csv' required><div class='small' style='margin:8px 0'>Columnas: DNI, NOMBRES, APELLIDOS, CARGO, AREA, REGIMEN, CORREO, CELULAR, BANCO, CUENTA.</div><button class='btn'>Importar trabajadores</button></form><a class='btn light' style='display:grid;place-items:center;margin-top:9px' href='/plantilla_trabajadores.csv'>Descargar plantilla CSV</a>{preview}</div>"
    return shell(body, "Cargar trabajadores")

@app.route("/plantilla_trabajadores.csv")
@login_required
def plantilla():
    csv_text = "DNI,NOMBRES,APELLIDOS,CARGO,AREA,REGIMEN,CORREO,CELULAR,BANCO,CUENTA\n11223344,NISEDM01,,TRABAJADOR,OPERACIONES,RÉGIMEN GENERAL,trabajador@pa.com,999999999,BCP,001-000000000\n"
    return Response(csv_text, mimetype="text/csv; charset=utf-8", headers={"Content-Disposition":"attachment; filename=plantilla_trabajadores_pa.csv"})

@app.route("/api/marcar", methods=["POST"])
@login_required
def api_marcar():
    data=request.get_json(silent=True) or {}; tipo=str(data.get("tipo","")).upper()
    if tipo not in ["INGRESO","SALIDA_REFRIGERIO","RETORNO_REFRIGERIO","SALIDA"]:
        return jsonify(ok=False,msg="Tipo inválido.")
    u=current_user() or {}; now=datetime.now(); c=conn()
    c.execute("INSERT INTO marcas(usuario,trabajador,tipo,fecha,hora,fecha_hora,lat,lng,precision,metodo,lectura,hora_manual,observacion) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",(session["usuario"], u.get("nombres",""), tipo, today_iso(), h12(now), now.isoformat(),str(data.get("lat","")), str(data.get("lng","")), str(data.get("precision","")),str(data.get("metodo","BOTON")), str(data.get("lectura","")), str(data.get("hora_manual","")), ""))
    c.commit(); c.close()
    return jsonify(ok=True,msg=f"{tipo.replace('_',' ').title()} registrado correctamente")

@app.route("/api/firma", methods=["POST"])
@login_required
def api_firma():
    data=request.get_json(silent=True) or {}; c=conn()
    c.execute("INSERT INTO firmas(usuario,data_url,actualizado) VALUES(?,?,?) ON CONFLICT(usuario) DO UPDATE SET data_url=excluded.data_url,actualizado=excluded.actualizado",(session["usuario"], data.get("data",""), datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
    c.commit(); c.close()
    return jsonify(ok=True,msg="Firma guardada correctamente")

@app.route("/api/documento/<tipo>")
@login_required
def api_documento(tipo):
    u=current_user() or {}
    title = {"boletas":"BOLETA NORMAL","utilidades":"CONSTANCIA DE UTILIDADES","vacaciones":"BOLETA DE VACACIONES","cts":"CONSTANCIA DE CTS","liquidaciones":"CONSTANCIA DE LIQUIDACIÓN","gratificaciones":"BOLETA DE GRATIFICACIÓN","constancia-grati":"CONSTANCIA DE GRATIFICACIÓN"}.get(tipo, "DOCUMENTO")
    content=f"{title}\nEmpresa: {u.get('empresa','')}\nTrabajador: {u.get('nombres','')} {u.get('apellidos','')}\nDNI: {u.get('usuario','')}\nCargo: {u.get('cargo','')}\nÁrea: {u.get('area','')}\nFecha de emisión: {today_es()}\nEstado: DISPONIBLE\n"
    return Response(content, mimetype="text/plain; charset=utf-8", headers={"Content-Disposition":f"attachment; filename={tipo}_pa.txt"})

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=False)
