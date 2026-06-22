# -*- coding: utf-8 -*-
"""
P&A Mobile - estilo Tareo Movil clasico.
Render ready: HTML/CSS/JS embebido en este app.py.
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
DB_PATH = os.path.join(PERSIST_DIR, "pa_mobile.db")

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "pa-mobile-render-secret")

COMPANY = "P&A Systems S.A.C."
VERSION = "V.3.2.15"
GREEN = "#2f773b"
LIME = "#a8d32a"

# ---------------- DB ----------------
def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    c = conn(); cur = c.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        nombres TEXT, apellidos TEXT, empresa TEXT, cargo TEXT, area TEXT,
        regimen TEXT, fecha_ingreso TEXT, tipo_doc TEXT, nro_doc TEXT,
        banco TEXT, cuenta TEXT, correo TEXT, celular TEXT, direccion TEXT, rol TEXT DEFAULT 'trabajador'
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS marcas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT, tipo TEXT, fecha TEXT, hora TEXT, fecha_hora TEXT,
        lat TEXT, lng TEXT, precision TEXT, metodo TEXT, lectura TEXT, hora_manual TEXT
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
    if not cur.execute("SELECT id FROM users WHERE usuario=?", ("11223344",)).fetchone():
        cur.execute("""INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,rol)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        ("11223344", generate_password_hash("123456"), "NISEDM01", "", COMPANY, "TRABAJADOR", "OPERACIONES", "RÉGIMEN GENERAL", "2026-06-22", "DNI", "11223344", "BCP", "001-000000000", "trabajador@pa.com", "999999999", "TRUJILLO", "admin"))
    c.commit(); c.close()
init_db()

# ---------------- Helpers ----------------
def dni(v): return re.sub(r"\D", "", str(v or ""))[-8:]
def esc(v): return str(v or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def h12(dt=None):
    dt = dt or datetime.now(); h=dt.hour; am="a. m." if h<12 else "p. m."; hh=h%12 or 12
    return f"{hh:02d}:{dt.minute:02d} {am}"
def today_es():
    dias=["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
    meses=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    d=date.today(); return f"{dias[d.weekday()]}, {d.day:02d} de {meses[d.month-1]} de {d.year}"
def login_required(fn):
    @wraps(fn)
    def w(*a, **kw):
        if not session.get("usuario"): return redirect(url_for("login"))
        return fn(*a, **kw)
    return w
def current_user():
    if not session.get("usuario"): return None
    c=conn(); r=c.execute("SELECT * FROM users WHERE usuario=?", (session["usuario"],)).fetchone(); c.close()
    return dict(r) if r else None

def icon(name):
    svgs = {
        "support":"<svg viewBox='0 0 24 24'><path d='M12 3a8 8 0 0 0-8 8v4a3 3 0 0 0 3 3h2v-7H6a6 6 0 0 1 12 0h-3v7h3a3 3 0 0 0 3-3v-4a8 8 0 0 0-8-8z'/></svg>",
        "settings":"<svg viewBox='0 0 24 24'><path d='M19.4 13.5c.1-.5.1-1 .1-1.5s0-1-.1-1.5l2-1.5-2-3.5-2.4 1a7 7 0 0 0-2.6-1.5L14 2h-4l-.4 2.5A7 7 0 0 0 7 6L4.6 5l-2 3.5 2 1.5a8 8 0 0 0 0 3l-2 1.5 2 3.5 2.4-1a7 7 0 0 0 2.6 1.5L10 22h4l.4-2.5A7 7 0 0 0 17 18l2.4 1 2-3.5-2-1.5zM12 15.5A3.5 3.5 0 1 1 12 8a3.5 3.5 0 0 1 0 7.5z'/></svg>",
        "profile":"<svg viewBox='0 0 24 24'><path d='M12 12a4.5 4.5 0 1 0 0-9 4.5 4.5 0 0 0 0 9zm0 2c-5 0-8 2.5-8 5.5V21h16v-1.5c0-3-3-5.5-8-5.5z'/></svg>",
        "list":"<svg viewBox='0 0 24 24'><path d='M4 5h3v3H4V5zm0 6h3v3H4v-3zm0 6h3v3H4v-3zM10 6h10v2H10V6zm0 6h10v2H10v-2zm0 6h10v2H10v-2z'/></svg>",
        "doc":"<svg viewBox='0 0 24 24'><path d='M6 2h9l5 5v15H6V2zm8 1.5V8h4.5L14 3.5z'/></svg>",
        "sign":"<svg viewBox='0 0 24 24'><path d='M4 17.5V21h3.5L18.1 10.4l-3.5-3.5L4 17.5zM20.7 7.8c.4-.4.4-1 0-1.4l-2.1-2.1a1 1 0 0 0-1.4 0l-1.6 1.6 3.5 3.5 1.6-1.6z'/></svg>",
        "home":"<svg viewBox='0 0 24 24'><path d='M3 11 12 3l9 8v10h-6v-6H9v6H3V11z'/></svg>",
        "logout":"<svg viewBox='0 0 24 24'><path d='M10 3h10v18H10v-2h8V5h-8V3zM8.6 16.6 7.2 18 2 12.8l5.2-5.2L8.6 9l-2.8 2.8H14v2H5.8l2.8 2.8z'/></svg>",
        "back":"<svg viewBox='0 0 24 24'><path d='M15.5 4.5 8 12l7.5 7.5-1.8 1.8L4.4 12l9.3-9.3 1.8 1.8z'/></svg>",
        "qr":"<svg viewBox='0 0 24 24'><path d='M3 3h8v8H3V3zm2 2v4h4V5H5zm8-2h8v8h-8V3zm2 2v4h4V5h-4zM3 13h8v8H3v-8zm2 2v4h4v-4H5zm10-2h2v2h-2v-2zm4 0h2v4h-4v-2h2v-2zm-6 4h2v4h-2v-4zm4 2h4v2h-4v-2z'/></svg>",
        "finger":"<svg viewBox='0 0 24 24'><path d='M12 2C8.7 2 6 4.7 6 8v4.2c0 2.6-1.3 4-2 4.8l1.5 1.3c.9-1 2.5-2.9 2.5-6.1V8a4 4 0 1 1 8 0v2h2V8c0-3.3-2.7-6-6-6zm0 4a2 2 0 0 0-2 2v4.5c0 3.6-1.5 6-2.7 7.3l1.5 1.4c1.4-1.5 3.2-4.4 3.2-8.7V8h0a0 0 0 1 1 0 0v4c0 4.7-1.7 7.4-2.6 8.7l1.7 1c1-1.5 2.9-4.6 2.9-9.7V8a2 2 0 0 0-2-2zm4 6v1c0 3.3-1 5.9-2.4 8.1l1.7 1c1.6-2.5 2.7-5.4 2.7-9.1v-1h-2z'/></svg>",
        "barcode":"<svg viewBox='0 0 24 24'><path d='M3 5h2v14H3V5zm3 0h1v14H6V5zm3 0h2v14H9V5zm3 0h1v14h-1V5zm3 0h3v14h-3V5zm4 0h1v14h-1V5z'/></svg>",
        "clock":"<svg viewBox='0 0 24 24'><path d='M12 2a10 10 0 1 0 .1 20.1A10 10 0 0 0 12 2zm1 11h5v-2h-4V6h-2v7z'/></svg>",
        "plus":"<svg viewBox='0 0 24 24'><path d='M11 5h2v6h6v2h-6v6h-2v-6H5v-2h6V5z'/></svg>",
        "refresh":"<svg viewBox='0 0 24 24'><path d='M17.7 6.3A8 8 0 1 0 20 12h-2a6 6 0 1 1-1.8-4.3L13 11h8V3l-3.3 3.3z'/></svg>",
    }
    return "<span class='ico'>" + svgs.get(name, svgs["doc"]) + "</span>"

def leaves(): return "<div class='leafs'><i></i><i></i><i></i></div>"
def bottom(active='home'):
    data=[('home','/home','home','Home'),('asistencia','/asistencia','list','Asistencia'),('documentos','/documentos','doc','Documentos'),('perfil','/perfil','profile','Perfil')]
    return "<div class='bottom'>"+"".join([f"<a class='{ 'on' if k==active else '' }' href='{u}'>{icon(ic)}<b>{t}</b></a>" for k,u,ic,t in data])+"</div>"
def page_head(title, ic='doc', back='/home'):
    return f"<div class='page-head'><a class='back' href='{back}'>{icon('back')}</a><div class='headicon'>{icon(ic)}</div><h1>{title}</h1></div>"

CSS = r'''
:root{--g:#2f773b;--gd:#276a33;--lime:#a8d32a;--ink:#06122a;--muted:#768096;--soft:#f4f6f7;--line:#e4ebe6;--shadow:0 9px 22px rgba(0,0,0,.15)}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}html,body{margin:0;min-height:100%;font-family:Arial,Segoe UI,sans-serif;background:#fff;color:var(--ink);font-weight:800}a{text-decoration:none;color:inherit}.ico{display:inline-grid;place-items:center;width:1.1em;height:1.1em;vertical-align:-.15em}.ico svg{width:100%;height:100%;fill:currentColor;display:block}.phone{width:100%;max-width:390px;min-height:100dvh;margin:0 auto;background:#fff;position:relative;border-left:1px solid #dfe7e3;border-right:1px solid #dfe7e3;overflow-x:hidden}.green{background:var(--g);color:#fff}.topmini{display:flex;justify-content:space-between;align-items:center;padding:10px 14px 0;font-size:12px;font-weight:1000}.topmini a,.topmini button{border:0;background:transparent;color:#fff;font-weight:1000;font-size:12px;padding:0}.topmini .ico{width:13px;height:13px;margin-right:3px}.login{min-height:100dvh;position:relative;background:white;padding:0 20px 20px}.login-hero{height:270px;margin:0 -20px;text-align:center;background:var(--g);color:white}.round-logo{width:112px;height:112px;border-radius:999px;background:white;border:7px solid var(--lime);display:grid;place-items:center;color:var(--g);box-shadow:0 4px 12px rgba(0,0,0,.25);margin:38px auto 14px}.round-logo .ico{width:54px;height:54px;color:var(--lime)}.brand{font-size:22px;font-weight:1000;letter-spacing:.3px;text-transform:uppercase;text-shadow:0 2px 3px rgba(0,0,0,.18)}.subbrand{font-size:12px;text-transform:uppercase;margin-top:4px;font-weight:1000}.card{background:#fff;border-radius:10px;box-shadow:var(--shadow);border:1px solid var(--line)}.login-card{margin:-34px 0 0;padding:14px 15px 18px}.role{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px}.role button{height:42px;border-radius:8px;border:1px solid #d6e1d8;background:#fff;color:var(--g);font-weight:1000}.role .on{background:var(--g);color:white}.form-title{text-align:center;font-size:13px;margin:2px 0 14px}.label{font-size:10px;color:#334155;margin:9px 0 5px;font-weight:1000}.input{width:100%;height:39px;border:1px solid #d8e2dc;border-radius:7px;padding:0 10px;font-size:14px;font-weight:900}.btn{height:44px;border:0;border-radius:8px;background:var(--g);color:white;font-weight:1000;font-size:16px;width:100%;box-shadow:0 6px 13px rgba(47,119,59,.18)}.link{text-align:center;font-size:11px;color:#0071ff;font-weight:1000;margin-top:10px}.mini-tiles{display:flex;justify-content:center;gap:18px;margin-top:28px}.mini-tile{width:76px;height:76px;border-radius:8px;background:white;box-shadow:var(--shadow);color:var(--g);display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:9px;font-weight:1000;text-align:center}.mini-tile .ico{width:27px;height:27px;margin-bottom:5px}.login .leafs{margin-top:10px}.help{position:fixed;bottom:16px;left:50%;transform:translateX(-50%);width:390px;max-width:100%;text-align:center;font-size:11px;color:#7a8494}.home-hero{height:226px;background:var(--g);color:white;border-radius:0 0 18px 18px}.avatar{width:82px;height:82px;border-radius:999px;background:white;color:var(--g);display:grid;place-items:center;margin:17px auto 8px}.avatar .ico{width:52px;height:52px}.name{text-align:center;font-size:14px;font-weight:1000;text-transform:uppercase}.welcome{height:45px;margin:16px 18px 0;background:white;border-radius:9px;box-shadow:var(--shadow);color:var(--ink);display:grid;place-items:center;font-size:13px}.content{padding:28px 16px 82px}.stats{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:0 0 14px}.stat{background:#fbfffb;border:1px solid #e1ece4;border-radius:10px;padding:10px;text-align:center;color:var(--g);font-size:12px}.stat b{display:block;color:var(--ink);font-size:18px;margin-top:3px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.tile{height:92px;border-radius:9px;background:white;box-shadow:var(--shadow);border:1px solid #f0f0f0;color:var(--ink);display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:12px;font-weight:1000}.tile .ico{width:34px;height:34px;color:var(--g);margin-bottom:7px}.tile:hover,.doc-row:hover{background:#eaf5eb}.leafs{width:115px;height:115px;margin:24px auto 0;position:relative;opacity:.42}.leafs i{position:absolute;width:60px;height:96px;border-radius:70% 30% 70% 30%;filter:drop-shadow(0 4px 4px rgba(0,0,0,.11))}.leafs i:nth-child(1){background:#ffd5c6;left:10px;transform:rotate(25deg)}.leafs i:nth-child(2){background:#cbd9b6;left:49px;transform:rotate(25deg)}.leafs i:nth-child(3){background:#fff1ad;left:74px;top:30px;transform:rotate(25deg)}.sync{position:fixed;left:10px;bottom:9px;width:330px;max-width:calc(100vw - 60px);background:white;border-radius:10px;box-shadow:var(--shadow);display:grid;grid-template-columns:34px 1fr;gap:8px;align-items:center;padding:8px 10px;z-index:20}.syncbtn{width:32px;height:32px;border-radius:999px;background:var(--g);color:white;display:grid;place-items:center;font-size:19px}.sync .t{font-size:11px;font-weight:1000}.sync .s{font-size:9px;color:#16823c;margin-top:2px}.exit{position:fixed;right:13px;bottom:12px;color:#ef4444;font-size:26px;z-index:30}.page-head{height:132px;background:var(--g);color:white;border-radius:0 0 15px 15px;text-align:center;position:relative;padding-top:18px}.back{position:absolute;left:13px;top:17px;color:white}.back .ico{width:24px;height:24px}.headicon{font-size:33px;margin-top:14px}.headicon .ico{width:34px;height:34px}.page-head h1{font-size:14px;text-transform:uppercase;margin:12px 0 0;font-weight:1000;letter-spacing:.2px}.toolbar{height:57px;margin:-24px 10px 8px;background:white;border-radius:8px;box-shadow:var(--shadow);display:flex;align-items:center;gap:18px;padding:0 14px;color:var(--g);position:relative;z-index:2}.toolbtn{border:0;background:transparent;color:var(--g);font-size:24px}.empty{margin:8px 10px;background:white;border-radius:9px;box-shadow:var(--shadow);padding:22px;text-align:center;font-weight:500;color:#495163}.bottom{position:fixed;bottom:0;left:50%;transform:translateX(-50%);width:100%;max-width:390px;height:58px;background:#fff;border-top:1px solid #e4ebe6;display:flex;z-index:40}.bottom a{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#6d7b72;font-size:10px;font-weight:900}.bottom a.on{color:var(--g)}.bottom .ico{width:21px;height:21px;margin-bottom:3px}.att-card{margin:-21px 10px 0;background:white;border-radius:10px;box-shadow:var(--shadow);padding:12px 10px 14px;position:relative;z-index:2}.sys{text-align:center;color:#008c3e;font-size:10px;font-weight:1000}.clock{text-align:center;font-size:32px;color:#05091f;font-weight:1000;margin:7px 0 3px}.date{text-align:center;color:#8690a2;font-size:10px;margin-bottom:10px}.mark{width:100%;height:42px;border:0;border-radius:7px;background:#eef1f3;color:#405064;font-size:13px;font-weight:1000;margin-bottom:8px}.mark.main{background:#22b45e;color:white}.mark:hover,.mark.active{background:var(--g);color:#fff}.row2{display:grid;grid-template-columns:1fr 1fr;gap:8px}.methods{font-size:10px;color:#667085;text-transform:uppercase;font-weight:1000;margin:12px 0 8px}.method-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}.method{height:74px;border-radius:8px;border:1px solid #e0e7e3;background:white;box-shadow:0 5px 13px rgba(0,0,0,.08);display:flex;flex-direction:column;align-items:center;justify-content:center;color:var(--ink);font-size:11px;font-weight:1000}.method .ico{width:34px;height:34px;color:var(--g);margin-bottom:4px}.map-title{display:flex;justify-content:space-between;align-items:center;margin:12px 0 8px;color:#838da0;font-size:10px;font-weight:1000}.signal{border:1px solid var(--g);background:#d9f8df;color:#077d3a;border-radius:15px;padding:6px 10px;font-size:10px;font-weight:1000}.map{height:108px;border:1px solid #d6dde6;border-radius:10px;background:#f2f4f6;position:relative;overflow:hidden}.road1{position:absolute;width:300px;height:12px;background:#fbd577;left:-30px;top:37px;transform:rotate(38deg)}.road2{position:absolute;width:290px;height:9px;background:#cfd5dd;left:-30px;top:70px;transform:rotate(-2deg)}.road3{position:absolute;width:13px;height:190px;background:#cfd5dd;left:175px;top:-43px;transform:rotate(24deg)}.pin{position:absolute;left:50%;top:55%;transform:translate(-50%,-50%);font-size:37px;color:#1f76ff}.bubble{position:absolute;left:50%;top:42%;transform:translate(-50%,-50%);background:white;border-radius:6px;padding:7px 10px;box-shadow:0 5px 12px rgba(0,0,0,.15);font-size:10px;white-space:nowrap}.loc{display:flex;justify-content:space-between;gap:6px;font-size:10px;color:#697386;font-weight:900;margin:7px 2px 0}.loc a{color:#006fff}.selected{height:38px;border-radius:8px;background:#eef1f3;display:flex;justify-content:center;align-items:center;color:#405064;font-size:12px;font-weight:1000;margin-top:8px}.history{background:#fbfffb;border:1px solid #dcebdd;border-radius:9px;padding:8px;margin-top:10px}.histrow{display:flex;justify-content:space-between;border-bottom:1px solid #e7f1e7;padding:6px 2px;font-size:11px}.histrow:last-child{border-bottom:0}.histrow b{color:var(--g)}.doc-list{padding:18px 10px 76px;background:#fafafa;min-height:calc(100dvh - 132px)}.doc-row{height:64px;border-radius:9px;background:white;border:1px solid #ebeff2;box-shadow:0 5px 13px rgba(0,0,0,.05);display:grid;grid-template-columns:42px 1fr 18px;align-items:center;padding:0 10px;margin-bottom:10px}.doc-ico{width:32px;height:32px;border-radius:7px;border:1px solid #cae9d1;color:var(--g);display:grid;place-items:center}.doc-title{font-size:13px;font-weight:1000}.doc-sub{font-size:10px;color:#8b94a3;margin-top:3px}.chev{font-size:23px;color:#9aa2ad}.form-card{margin:12px 10px 76px;background:white;border-radius:10px;box-shadow:var(--shadow);padding:12px}.field{margin-bottom:9px}.field label{font-size:10px;font-weight:1000;color:var(--g);display:block;margin-bottom:4px}.field input,.field select{width:100%;height:36px;border:1px solid #dbe5df;border-radius:7px;padding:0 9px;font-weight:800}.two{display:grid;grid-template-columns:1fr 1fr;gap:8px}.cfg{width:100%;min-height:42px;border-radius:7px;border:1px solid var(--g);color:var(--g);background:white;font-weight:1000;margin-bottom:9px}.cfg.primary{background:var(--g);color:white}.profile-card{text-align:center}.photo{width:78px;height:78px;border-radius:999px;background:#edf4ef;color:var(--g);display:grid;place-items:center;margin:0 auto 12px}.photo .ico{width:46px;height:46px}.kv{display:grid;gap:8px;text-align:left;margin-top:14px}.kv div{background:#f4f6f7;border-radius:8px;padding:10px}.kv small{display:block;color:#7e8794;font-size:10px}.kv b{display:block;margin-top:4px;font-size:11px}.signbox{height:190px;border:2px dashed #dbe3ef;border-radius:10px;margin:12px 10px;background:white}#canvas{width:100%;height:100%;touch-action:none}.grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:0 10px 10px}.mini{height:39px;border:0;border-radius:7px;background:#f1f4f7;font-weight:1000}.danger{background:#feecec;color:#e93030}.save{background:var(--g);color:white}.toast{position:fixed;left:50%;top:12px;transform:translateX(-50%);background:#10251a;color:white;border-radius:999px;padding:10px 15px;font-size:12px;font-weight:1000;z-index:90}.wheel-modal{position:fixed;inset:auto 0 0;z-index:80;background:rgba(0,0,0,.25);display:none;align-items:end;justify-content:center}.wheel-modal.show{display:flex}.wheel-card{width:100%;max-width:390px;background:white;border-radius:20px 20px 0 0;padding:14px 18px 22px;box-shadow:0 -12px 35px rgba(0,0,0,.22)}.wheel-head{display:flex;justify-content:space-between;font-weight:1000;margin-bottom:8px}.wheel-head button{border:0;background:white;color:#0aa25b;font-weight:1000;font-size:16px}.wheels{height:150px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;position:relative}.wheels:before{content:'';position:absolute;left:0;right:0;top:53px;height:44px;background:rgba(0,0,0,.04);border-radius:12px}.wheel{height:150px;overflow:auto;scroll-snap-type:y mandatory;text-align:center;padding:53px 0;scrollbar-width:none}.wheel::-webkit-scrollbar{display:none}.wheel div{height:44px;line-height:44px;font-size:22px;color:#b6bdc7;scroll-snap-align:center;font-weight:500}.wheel div.sel{color:#303642;font-size:25px}.scanner{min-height:100dvh;background:#111;color:white}.scan-head{height:58px;background:var(--g);display:flex;align-items:center;justify-content:space-between;padding:0 14px;font-size:13px;font-weight:1000}.camera{height:calc(100dvh - 116px);position:relative;background:#222;display:grid;place-items:center;overflow:hidden}.camera video,.camera #reader{width:100%;min-height:100%;object-fit:cover}.scan-frame{position:absolute;width:240px;height:240px;border:3px solid #35c970;box-shadow:0 0 0 999px rgba(0,0,0,.25)}.scan-frame:after{content:'';position:absolute;left:-30px;right:-30px;top:50%;border-top:2px solid #22c55e}.scan-help{position:absolute;bottom:78px;left:0;right:0;text-align:center;color:white;font-size:12px;font-weight:900}.cancel{position:absolute;left:34px;right:34px;bottom:25px;height:45px;border:0;border-radius:8px;background:white;color:#111;font-weight:1000}.finger-wrap{height:calc(100dvh - 116px);display:flex;align-items:center;justify-content:center;flex-direction:column;background:white;color:#667085}.finger-big{width:160px;height:160px;border-radius:999px;border:8px solid #e5e7eb;color:#2f9c52;display:grid;place-items:center;margin-bottom:20px}.finger-big .ico{width:115px;height:115px}.mass-preview{max-height:170px;overflow:auto;background:white;border:1px solid #e5eee6;border-radius:8px;padding:8px;font-size:11px}.ok{background:#e7f8eb;color:#166534;border:1px solid #bfe7c7;border-radius:8px;padding:9px;font-size:12px;font-weight:900;margin-bottom:10px}.err{background:#fee2e2;color:#991b1b;border:1px solid #fecaca;border-radius:8px;padding:9px;font-size:12px;font-weight:900;margin-bottom:10px}@media(min-width:800px){body{display:block}.phone{margin:0 auto}.sync{left:10px;transform:none}.exit{right:13px}}
'''

JS = r'''
function tone(ok=true){try{const C=window.AudioContext||window.webkitAudioContext,ctx=new C(),o=ctx.createOscillator(),g=ctx.createGain();o.type='sine';o.frequency.value=ok?880:220;g.gain.value=.07;o.connect(g);g.connect(ctx.destination);o.start();setTimeout(()=>{o.stop();ctx.close()},150)}catch(e){}}
function toast(msg,ok=true){tone(ok);let t=document.createElement('div');t.className='toast';t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.remove(),2300)}
function clock(){let e=document.getElementById('liveClock');if(!e)return;let d=new Date(),h=d.getHours(),a=h<12?'a. m.':'p. m.';h=h%12||12;e.textContent=String(h).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0')+' '+a}setInterval(clock,1000);document.addEventListener('DOMContentLoaded',()=>{clock();loc(false);initSign();})
function fakeOk(m){toast(m||'Opción habilitada')}
function loc(n=true){let info=document.getElementById('locText'),link=document.getElementById('mapLink'),pin=document.getElementById('pin');if(!info)return;if(!navigator.geolocation){info.textContent='Navegador sin GPS';return}navigator.geolocation.getCurrentPosition(p=>{let lat=p.coords.latitude.toFixed(6),lng=p.coords.longitude.toFixed(6),acc=Math.round(p.coords.accuracy);info.textContent='Ubicación real detectada con precisión aproximada de '+acc+' m.';if(link)link.href='https://www.google.com/maps?q='+lat+','+lng;if(pin){pin.dataset.lat=lat;pin.dataset.lng=lng;pin.dataset.acc=acc}if(n)toast('Ubicación actualizada')},()=>{info.textContent='Permite ubicación en el navegador.';if(n)toast('No se pudo obtener GPS',false)},{enableHighAccuracy:true,timeout:9000,maximumAge:0})}
function marcar(tipo,metodo='BOTON',lectura=''){let b=document.querySelector('[data-tipo="'+tipo+'"]');if(b){b.classList.add('active');setTimeout(()=>b.classList.remove('active'),1000)}let p=document.getElementById('pin')||{};let payload={tipo,metodo,lectura,hora_manual:(document.getElementById('selectedTime')?.dataset.value||''),lat:p.dataset?.lat||'',lng:p.dataset?.lng||'',precision:p.dataset?.acc||''};function send(){fetch('/api/marcar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).then(r=>r.json()).then(j=>{toast(j.msg||'Registrado',!!j.ok);setTimeout(()=>location.reload(),900)}).catch(()=>toast('Error registrando',false))}if(navigator.geolocation){navigator.geolocation.getCurrentPosition(pos=>{payload.lat=pos.coords.latitude;payload.lng=pos.coords.longitude;payload.precision=Math.round(pos.coords.accuracy);send()},()=>send(),{enableHighAccuracy:true,timeout:5000})}else send()}
function openWheel(){document.getElementById('wheelModal')?.classList.add('show');setNow();syncWheel()}function closeWheel(){document.getElementById('wheelModal')?.classList.remove('show')}function syncWheel(){document.querySelectorAll('.wheel').forEach(w=>{let mid=w.scrollTop+75;[...w.children].forEach(c=>c.classList.toggle('sel',Math.abs((c.offsetTop+22)-mid)<24))})}function setWheel(id,val){let w=document.getElementById(id);if(!w)return;[...w.children].forEach(c=>{if(String(c.dataset.v)===String(val))w.scrollTop=c.offsetTop-53})}function setNow(){let d=new Date();setWheel('wh',d.getHours()%12||12);setWheel('wm',d.getMinutes());setWheel('wa',d.getHours()<12?'AM':'PM')}function pick(id){let w=document.getElementById(id),mid=w.scrollTop+75,b=w.children[0],bd=999;[...w.children].forEach(c=>{let d=Math.abs((c.offsetTop+22)-mid);if(d<bd){bd=d;b=c}});return b.dataset.v}function applyWheel(){let h=pick('wh'),m=pick('wm'),a=pick('wa'),val=String(h).padStart(2,'0')+':'+String(m).padStart(2,'0')+' '+(a==='AM'?'a. m.':'p. m.');let e=document.getElementById('selectedTime');if(e){e.dataset.value=val;e.textContent='● Ajustar hora táctil: '+val}closeWheel();toast('Hora seleccionada')}document.addEventListener('scroll',e=>{if(e.target.classList&&e.target.classList.contains('wheel'))syncWheel()},true)
function initSign(){let c=document.getElementById('canvas');if(!c)return;let ctx=c.getContext('2d'),draw=false;function rz(){let r=c.getBoundingClientRect();c.width=r.width*devicePixelRatio;c.height=r.height*devicePixelRatio;ctx.setTransform(devicePixelRatio,0,0,devicePixelRatio,0,0);ctx.lineWidth=2;ctx.lineCap='round'}rz();function p(ev){let r=c.getBoundingClientRect(),t=ev.touches?ev.touches[0]:ev;return{x:t.clientX-r.left,y:t.clientY-r.top}}c.onpointerdown=e=>{draw=true;let q=p(e);ctx.beginPath();ctx.moveTo(q.x,q.y)};c.onpointermove=e=>{if(!draw)return;let q=p(e);ctx.lineTo(q.x,q.y);ctx.stroke()};c.onpointerup=()=>draw=false;c.onpointerleave=()=>draw=false;window.clearSign=()=>ctx.clearRect(0,0,c.width,c.height);window.saveSign=()=>fetch('/api/firma',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({data:c.toDataURL('image/png')})}).then(r=>r.json()).then(j=>toast(j.msg||'Firma guardada'))}
async function startScan(mode){let reader=document.getElementById('reader'),video=document.getElementById('camvideo');if(!reader&&!video)return;let q=(new URLSearchParams(location.search)).get('tipo')||'INGRESO';let stream=null;try{stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'}}); if(video){video.srcObject=stream;video.play();} if('BarcodeDetector' in window){let formats=mode==='barcode'?['code_128','code_39','ean_13','ean_8','upc_a','upc_e','qr_code']:['qr_code'];let det=new BarcodeDetector({formats});let run=async()=>{try{let codes=await det.detect(video);if(codes&&codes.length){let val=codes[0].rawValue||'';stream.getTracks().forEach(t=>t.stop());marcar(q,mode.toUpperCase(),val);return}}catch(e){}requestAnimationFrame(run)};run();}else{toast('Usa Chrome/Edge para lectura por cámara',false)}}catch(e){toast('No se pudo abrir cámara',false)}}
function simFinger(){setTimeout(()=>{marcar((new URLSearchParams(location.search)).get('tipo')||'INGRESO','HUELLA','SIMULADA')},1400)}
function previewCsv(input){let box=document.getElementById('preview');if(!box||!input.files[0])return;let r=new FileReader();r.onload=()=>{let lines=r.result.split(/\r?\n/).filter(Boolean).slice(0,6);box.innerHTML='<b>Vista previa:</b><br>'+lines.map(x=>'<div>'+x.replaceAll('<','&lt;')+'</div>').join('')};r.readAsText(input.files[0])}
'''

def shell(body,title='P&A Mobile',cls=''):
    return f"""<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1,viewport-fit=cover'><meta name='theme-color' content='{GREEN}'><title>{title}</title><style>{CSS}</style></head><body><div class='phone {cls}'>{body}</div><script>{JS}</script></body></html>"""

# ---------------- routes ----------------
@app.route('/')
def index():
    return redirect(url_for('home' if session.get('usuario') else 'login'))

@app.route('/login', methods=['GET','POST'])
def login():
    err=''
    if request.method=='POST':
        u=dni(request.form.get('usuario')); p=request.form.get('password','')
        c=conn(); r=c.execute('SELECT * FROM users WHERE usuario=?',(u,)).fetchone(); c.close()
        if r and check_password_hash(r['password_hash'],p): session['usuario']=u; return redirect(url_for('home'))
        err='<div class="err">Usuario o contraseña incorrectos.</div>'
    body=f"""<div class='login'><div class='login-hero'><div class='topmini'><button>{icon('support')} Soporte</button><button>{icon('settings')} Config.</button></div><div class='round-logo'>{icon('doc')}</div><div class='brand'>P&A MOBILE</div><div class='subbrand'>INICIAR SESIÓN</div></div><form class='card login-card' method='post'><div class='role'><button type='button' class='on'>USUARIO</button><button type='button'>ADMINISTRADOR</button></div><div class='form-title'>Inicia sesión para continuar</div>{err}<div class='label'>Usuario</div><input class='input' name='usuario' maxlength='8' inputmode='numeric' value='11223344' required><div class='label'>Empresa</div><select class='input'><option>{COMPANY}</option></select><div class='label'>Contraseña</div><input class='input' name='password' type='password' value='123456' required><br><br><button class='btn'>{icon('logout')} INGRESAR</button><div class='link'>¿Olvidaste tu contraseña?</div></form><div class='mini-tiles'><a class='mini-tile'>{icon('list')}ASISTENCIA</a><a class='mini-tile'>{icon('doc')}DOCUMENTOS</a></div>{leaves()}<div class='help'>¿Problemas para acceder?<br>Contactar a Mesa</div></div>"""
    return shell(body,'Login')

@app.route('/home')
@login_required
def home():
    u=current_user() or {}; c=conn(); hoy=date.today().isoformat(); rows=c.execute('SELECT tipo FROM marcas WHERE usuario=? AND fecha=? ORDER BY id',(session['usuario'],hoy)).fetchall(); c.close(); ult=rows[-1]['tipo'].replace('_',' ') if rows else 'SIN REG.'
    body=f"""<div class='home-hero'><div class='topmini'><button>{icon('support')} Soporte</button><a href='/config'>{icon('settings')} Config.</a></div><div class='avatar'>{icon('profile')}</div><div class='name'>{esc(u.get('nombres','NISEDM01'))}</div><div class='welcome'>Bienvenido(a)</div></div><div class='content'><div class='stats'><div class='stat'>Marcaciones<b>{len(rows)}</b></div><div class='stat'>Último<b>{esc(ult)}</b></div></div><div class='grid'><a class='tile' href='/asistencia'>{icon('list')}ASISTENCIA</a><a class='tile' href='/documentos'>{icon('doc')}DOCUMENTOS</a><a class='tile' href='/firma'>{icon('sign')}FIRMA</a><a class='tile' href='/perfil'>{icon('profile')}PERFIL</a></div>{leaves()}</div><div class='sync'><button class='syncbtn' onclick="fakeOk('Sincronizado')">↻</button><div><div class='t'>Sincronizar Tablas Maestras</div><div class='s'>Actualizado hasta: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div></div></div><a class='exit' href='/logout'>{icon('logout')}</a>"""
    return shell(body,'Home')

@app.route('/asistencia')
@login_required
def asistencia():
    c=conn(); rows=c.execute('SELECT * FROM marcas WHERE usuario=? AND fecha=? ORDER BY id',(session['usuario'],date.today().isoformat())).fetchall(); c.close()
    hist=''.join([f"<div class='histrow'><span>{esc(r['tipo'].replace('_',' '))}</span><b>{esc(r['hora_manual'] or r['hora'])}</b></div>" for r in rows]) or "<div class='histrow'><span>Sin marcaciones</span><b>--:--</b></div>"
    body=f"""<div class='page-head'><a class='back' href='/home'>{icon('back')}</a><div class='headicon'>{icon('list')}</div><h1>REGISTRA TU ASISTENCIA FÁCIL</h1></div><div class='att-card'><div class='sys'>HORA ACTUAL DEL SISTEMA</div><div class='clock' id='liveClock'>{h12()}</div><div class='date'>{today_es()}</div><button class='mark main' data-tipo='INGRESO' onclick="marcar('INGRESO')">↪ Registrar Ingreso</button><div class='row2'><button class='mark' data-tipo='SALIDA_REFRIGERIO' onclick="marcar('SALIDA_REFRIGERIO')">Salida a<br>Refrigerio</button><button class='mark' data-tipo='RETORNO_REFRIGERIO' onclick="marcar('RETORNO_REFRIGERIO')">Retorno de<br>Refrigerio</button></div><button class='mark' data-tipo='SALIDA' onclick="marcar('SALIDA')">↩ Registrar Salida</button><div id='selectedTime' class='selected' onclick='openWheel()' data-value=''>● Ajustar hora táctil: Ahora</div><div class='methods'>Métodos de marcación</div><div class='method-grid'><a class='method' href='/scanner/qr?tipo=INGRESO'>{icon('qr')}QR</a><a class='method' href='/huella?tipo=INGRESO'>{icon('finger')}HUELLA</a><a class='method' href='/scanner/barcode?tipo=INGRESO'>{icon('barcode')}CÓDIGO</a></div><div class='map-title'><span>UBICACIÓN ACTUAL</span><button class='signal' onclick='loc(true)'>SEÑAL FUERTE</button></div><div class='map'><div class='road1'></div><div class='road2'></div><div class='road3'></div><div class='bubble'>Tu ubicación actual</div><div class='pin' id='pin'>📍</div></div><div class='loc'><span id='locText'>Detectando ubicación real...</span><a id='mapLink' target='_blank'>Ver mapa</a></div><div class='history'><div class='label'>Historial del día</div>{hist}</div></div>{wheel_modal()}{bottom('asistencia')}"""
    return shell(body,'Asistencia')

def wheel_modal():
    hours=''.join([f"<div data-v='{i}'>{i:02d}</div>" for i in range(1,13)]); mins=''.join([f"<div data-v='{i}'>{i:02d}</div>" for i in range(60)]); ampm="<div data-v='AM'>AM</div><div data-v='PM'>PM</div>"
    return f"<div id='wheelModal' class='wheel-modal'><div class='wheel-card'><div class='wheel-head'><button onclick='closeWheel()'>Cancelar</button><span>Seleccionar hora</span><button onclick='applyWheel()'>Aceptar</button></div><div class='wheels'><div class='wheel' id='wh'>{hours}</div><div class='wheel' id='wm'>{mins}</div><div class='wheel' id='wa'>{ampm}</div></div></div></div>"

@app.route('/scanner/<mode>')
@login_required
def scanner(mode):
    title='LECTURA QR' if mode=='qr' else 'LECTURA CÓDIGO DE BARRAS'
    extra="<div class='scan-frame'></div><div class='scan-help'>Alinea el código dentro del marco</div>"
    body=f"""<div class='scanner'><div class='scan-head'><a href='/asistencia'>{icon('back')}</a><b>{title}</b><span>⚡</span></div><div class='camera'><video id='camvideo' autoplay muted playsinline></video>{extra}<button class='cancel' onclick="location.href='/asistencia'">CANCELAR</button></div>{bottom('asistencia')}</div><script>document.addEventListener('DOMContentLoaded',()=>startScan('{ 'barcode' if mode=='barcode' else 'qr' }'));</script>"""
    return shell(body,title)

@app.route('/huella')
@login_required
def huella():
    body=f"""<div class='scanner'><div class='scan-head'><a href='/asistencia'>{icon('back')}</a><b>LECTURA DE HUELLA</b><span>⚡</span></div><div class='finger-wrap'><div class='finger-big'>{icon('finger')}</div><b>Coloca tu dedo en el sensor</b><p>Leyendo huella...</p><button class='cancel' onclick="location.href='/asistencia'">CANCELAR</button></div>{bottom('asistencia')}</div><script>document.addEventListener('DOMContentLoaded',simFinger)</script>"""
    return shell(body,'Huella')

@app.route('/documentos')
@login_required
def documentos():
    docs=[('boletas','Boletas Normales'),('utilidades','Constancias de Utilidades'),('vacaciones','Boletas de Vacaciones'),('cts','Constancias de CTS'),('liquidaciones','Constancias de Liquidaciones'),('gratificaciones','Boletas de Gratificaciones'),('constancia-grati','Constancias de Gratificaciones')]
    rows=''.join([f"<a class='doc-row' href='/api/documento/{k}'><div class='doc-ico'>{icon('doc')}</div><div><div class='doc-title'>{t}</div><div class='doc-sub'>Disponible para descarga</div></div><div class='chev'>›</div></a>" for k,t in docs])
    return shell(page_head('GESTIÓN DE DOCUMENTOS','doc')+f"<div class='doc-list'>{rows}</div>{bottom('documentos')}",'Documentos')

@app.route('/firma')
@login_required
def firma():
    return shell(page_head('FIRMA DIGITAL','sign')+"<div class='signbox'><canvas id='canvas'></canvas></div><div class='grid2'><button class='mini danger' onclick='clearSign()'>Limpiar</button><button class='mini save' onclick='saveSign()'>Guardar →</button></div>"+bottom('home'),'Firma')

@app.route('/perfil')
@login_required
def perfil():
    u=current_user() or {}; items=[('DNI',u.get('usuario')),('Empresa',u.get('empresa')),('Área',u.get('area')),('Régimen',u.get('regimen')),('Ingreso',u.get('fecha_ingreso')),('Banco',u.get('banco')),('Cuenta',u.get('cuenta')),('Celular',u.get('celular'))]
    kv=''.join([f"<div><small>{a}</small><b>{esc(b)}</b></div>" for a,b in items])
    return shell(page_head('PERFIL','profile')+f"<div class='form-card profile-card'><div class='photo'>{icon('profile')}</div><h3>{esc(u.get('nombres'))}</h3><p>TRABAJADOR</p><div class='kv'>{kv}</div></div>{bottom('perfil')}",'Perfil')

@app.route('/config')
@login_required
def config():
    return shell(page_head('CONFIGURACIÓN','settings')+"""<div class='form-card'><a class='cfg primary' style='display:grid;place-items:center' href='/config/trabajador'>👥 Cargar trabajador</a><a class='cfg primary' style='display:grid;place-items:center' href='/config/masiva'>📥 Importación masiva trabajadores</a><a class='cfg' style='display:grid;place-items:center' href='/config/plantilla-trabajadores'>Plantilla trabajadores CSV</a><a class='cfg' style='display:grid;place-items:center' href='/config/db'>Conexión SQL Server futura</a><a class='cfg' style='display:grid;place-items:center' href='/home'>Volver</a></div>""",'Config')

@app.route('/config/trabajador', methods=['GET','POST'])
@login_required
def config_trabajador():
    msg=''
    if request.method=='POST':
        d=dni(request.form.get('dni'))
        if len(d)!=8: msg="<div class='err'>DNI inválido.</div>"
        else:
            vals=(d,generate_password_hash('123456'),request.form.get('nombres',''),request.form.get('apellidos',''),COMPANY,request.form.get('cargo','TRABAJADOR'),request.form.get('area','OPERACIONES'),request.form.get('regimen','RÉGIMEN GENERAL'),request.form.get('fecha_ingreso',''),request.form.get('tipo_doc','DNI'),d,request.form.get('banco',''),request.form.get('cuenta',''),request.form.get('correo',''),request.form.get('celular',''),request.form.get('direccion',''),'trabajador')
            c=conn(); c.execute("""INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,rol) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(usuario) DO UPDATE SET nombres=excluded.nombres,apellidos=excluded.apellidos,cargo=excluded.cargo,area=excluded.area,regimen=excluded.regimen,fecha_ingreso=excluded.fecha_ingreso,banco=excluded.banco,cuenta=excluded.cuenta,correo=excluded.correo,celular=excluded.celular,direccion=excluded.direccion""",vals); c.commit(); c.close(); msg="<div class='ok'>Trabajador guardado. Clave: 123456</div>"
    form="""<div class='field'><label>DNI *</label><input name='dni' maxlength='8' required></div><div class='field'><label>Nombres *</label><input name='nombres' required></div><div class='field'><label>Apellidos</label><input name='apellidos'></div><div class='two'><div class='field'><label>Cargo</label><input name='cargo'></div><div class='field'><label>Área</label><input name='area'></div></div><div class='two'><div class='field'><label>Régimen</label><select name='regimen'><option>RÉGIMEN GENERAL</option><option>RÉGIMEN AGRARIO</option></select></div><div class='field'><label>Ingreso</label><input type='date' name='fecha_ingreso'></div></div><div class='two'><div class='field'><label>Banco</label><input name='banco'></div><div class='field'><label>Cuenta</label><input name='cuenta'></div></div><div class='field'><label>Correo</label><input type='email' name='correo'></div><div class='field'><label>Celular</label><input name='celular'></div><button class='btn'>Guardar trabajador</button>"""
    return shell(page_head('CARGAR TRABAJADOR','profile','/config')+f"<form class='form-card' method='post'>{msg}{form}</form>",'Cargar')

@app.route('/config/masiva', methods=['GET','POST'])
@login_required
def masiva():
    msg=''
    if request.method=='POST':
        f=request.files.get('archivo')
        if not f: msg="<div class='err'>Seleccione un archivo CSV.</div>"
        else:
            text=f.read().decode('utf-8-sig','ignore'); rd=csv.DictReader(StringIO(text)); count=0; bad=0; c=conn()
            for r in rd:
                d=dni(r.get('DNI') or r.get('dni') or r.get('USUARIO') or '')
                if len(d)!=8: bad+=1; continue
                vals=(d,generate_password_hash('123456'),r.get('NOMBRES',''),r.get('APELLIDOS',''),COMPANY,r.get('CARGO','TRABAJADOR'),r.get('AREA','OPERACIONES'),r.get('REGIMEN','RÉGIMEN GENERAL'),r.get('FECHA_INGRESO',''),r.get('TIPO_DOC','DNI'),d,r.get('BANCO',''),r.get('CUENTA',''),r.get('CORREO',''),r.get('CELULAR',''),r.get('DIRECCION',''),'trabajador')
                c.execute("""INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,rol) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(usuario) DO UPDATE SET nombres=excluded.nombres,apellidos=excluded.apellidos,cargo=excluded.cargo,area=excluded.area,regimen=excluded.regimen,fecha_ingreso=excluded.fecha_ingreso,banco=excluded.banco,cuenta=excluded.cuenta,correo=excluded.correo,celular=excluded.celular,direccion=excluded.direccion""",vals); count+=1
            c.commit(); c.close(); msg=f"<div class='ok'>Importación finalizada: {count} trabajador(es). Errores: {bad}.</div>"
    body=page_head('IMPORTACIÓN MASIVA','doc','/config')+f"""<form class='form-card' method='post' enctype='multipart/form-data'>{msg}<div class='field'><label>Archivo CSV</label><input type='file' name='archivo' accept='.csv' onchange='previewCsv(this)' required></div><div id='preview' class='mass-preview'>Sube la plantilla CSV para revisar los primeros registros.</div><br><button class='btn'>Importar trabajadores</button><br><br><a class='cfg' style='display:grid;place-items:center' href='/config/plantilla-trabajadores'>Descargar plantilla</a></form>"""
    return shell(body,'Masiva')

@app.route('/config/plantilla-trabajadores')
@login_required
def plantilla():
    csvtxt='DNI,NOMBRES,APELLIDOS,CARGO,AREA,REGIMEN,FECHA_INGRESO,BANCO,CUENTA,CORREO,CELULAR,DIRECCION\n11223344,NISEDM01,,TRABAJADOR,OPERACIONES,RÉGIMEN GENERAL,2026-06-22,BCP,001-000000000,trabajador@pa.com,999999999,TRUJILLO\n'
    return Response(csvtxt,mimetype='text/csv; charset=utf-8',headers={'Content-Disposition':'attachment; filename=plantilla_trabajadores_pa.csv'})

@app.route('/config/db', methods=['GET','POST'])
@login_required
def config_db():
    msg=''
    if request.method=='POST':
        vals=(request.form.get('servidor'),request.form.get('base_datos'),request.form.get('usuario'),request.form.get('password'),request.form.get('puerto','1433'),datetime.now().isoformat())
        c=conn(); c.execute("INSERT INTO config_db(motor,servidor,base_datos,usuario,password,puerto,actualizado) VALUES('SQLSERVER',?,?,?,?,?,?)", vals); c.commit(); c.close(); msg="<div class='ok'>Configuración guardada para SQL Server futuro.</div>"
    form="""<div class='field'><label>Servidor</label><input name='servidor' placeholder='192.168.1.10'></div><div class='field'><label>Base de Datos</label><input name='base_datos' placeholder='RRHH_PA'></div><div class='field'><label>Usuario</label><input name='usuario' placeholder='sa'></div><div class='field'><label>Contraseña</label><input name='password' type='password'></div><div class='field'><label>Puerto</label><input name='puerto' value='1433'></div><button class='btn'>Guardar Configuración</button>"""
    return shell(page_head('CONEXIÓN BD','settings','/config')+f"<form class='form-card' method='post'>{msg}{form}</form>",'BD')

@app.route('/api/marcar', methods=['POST'])
@login_required
def api_marcar():
    data=request.get_json(silent=True) or {}; tipo=(data.get('tipo') or '').upper()
    if tipo not in ['INGRESO','SALIDA_REFRIGERIO','RETORNO_REFRIGERIO','SALIDA']: return jsonify(ok=False,msg='Tipo inválido')
    now=datetime.now(); c=conn(); c.execute("""INSERT INTO marcas(usuario,tipo,fecha,hora,fecha_hora,lat,lng,precision,metodo,lectura,hora_manual) VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (session['usuario'],tipo,date.today().isoformat(),h12(now),now.isoformat(),str(data.get('lat','')),str(data.get('lng','')),str(data.get('precision','')),str(data.get('metodo','BOTON')),str(data.get('lectura','')),str(data.get('hora_manual','')))); c.commit(); c.close()
    return jsonify(ok=True,msg=tipo.replace('_',' ').title()+' registrado')

@app.route('/api/firma', methods=['POST'])
@login_required
def api_firma():
    data=request.get_json(silent=True) or {}; c=conn(); c.execute("INSERT INTO firmas(usuario,data_url,actualizado) VALUES(?,?,?) ON CONFLICT(usuario) DO UPDATE SET data_url=excluded.data_url, actualizado=excluded.actualizado", (session['usuario'],data.get('data',''),datetime.now().isoformat())); c.commit(); c.close(); return jsonify(ok=True,msg='Firma guardada')

@app.route('/api/documento/<tipo>')
@login_required
def api_doc(tipo):
    u=current_user() or {}; txt=f"P&A MOBILE - {tipo}\nTrabajador: {u.get('nombres','')}\nDNI: {u.get('usuario','')}\nEmpresa: {u.get('empresa','')}\nFecha: {today_es()}\n"
    return Response(txt,mimetype='text/plain; charset=utf-8',headers={'Content-Disposition':f'attachment; filename={tipo}_pa.txt'})

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT','5000')), debug=False)
