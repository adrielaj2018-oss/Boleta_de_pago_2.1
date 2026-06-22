# -*- coding: utf-8 -*-
"""
P&A Mobile - Asistencia, Documentos, Firma y Configuración
Render 5 archivos: todo el HTML/CSS/JS está embebido en app.py.
"""
import os, re, sqlite3, json, csv
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

BRAND = "P&A"
COMPANY = "P&A Systems S.A.C."
VERSION = "V.3.2.15"
GREEN = "#28763A"
GREEN_DARK = "#1f5f31"
LIME = "#b8d81b"

# ------------------------- DB -------------------------
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
        banco TEXT, cuenta TEXT, correo TEXT, celular TEXT, direccion TEXT, avatar TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS marcas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT, tipo TEXT, fecha TEXT, hora TEXT, fecha_hora TEXT,
        lat TEXT, lng TEXT, precision TEXT, direccion TEXT, hora_manual TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS firmas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE, data_url TEXT, almacenamiento TEXT DEFAULT 'LOCAL', actualizado TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS config_db(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        motor TEXT DEFAULT 'SQLSERVER', servidor TEXT, base_datos TEXT, usuario TEXT,
        password TEXT, puerto TEXT DEFAULT '1433', actualizado TEXT
    )""")
    cur.execute("SELECT id FROM users WHERE usuario=?", ("11223344",))
    if not cur.fetchone():
        cur.execute("""INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,avatar)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    ("11223344", generate_password_hash("123456"), "NISEDM01", "", COMPANY,
                     "TRABAJADOR", "OPERACIONES", "RÉGIMEN GENERAL", "2026-06-22", "DNI", "11223344",
                     "BCP", "001-000000000", "trabajador@pa.com", "999999999", "TRUJILLO", ""))
    c.commit(); c.close()

init_db()

# ------------------------- Helpers -------------------------
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

def today_es():
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    d = date.today()
    return f"{dias[d.weekday()]}, {d.day:02d} de {meses[d.month-1]} de {d.year}"

def h12(dt=None):
    dt = dt or datetime.now(); h = dt.hour; ampm = "a. m." if h < 12 else "p. m."; hh = h % 12 or 12
    return f"{hh:02d}:{dt.minute:02d} {ampm}"

def esc(s):
    return str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def dni(v):
    return re.sub(r"\D", "", str(v or ""))[-8:]

# ------------------------- CSS/JS -------------------------
CSS = f"""
:root{{--green:{GREEN};--green-dark:{GREEN_DARK};--lime:{LIME};--ink:#071126;--muted:#8b94a3;--soft:#f5f7f8;--line:#e6eaee;--shadow:0 9px 24px rgba(0,0,0,.13)}}
*{{box-sizing:border-box;-webkit-tap-highlight-color:transparent}}html,body{{margin:0;width:100%;min-height:100%;font-family:Inter,Segoe UI,Arial,sans-serif;background:#eef2f1;color:var(--ink)}}body{{display:flex;justify-content:center}}a{{text-decoration:none;color:inherit}}
.phone{{width:100%;max-width:390px;min-height:100dvh;background:#fff;position:relative;overflow-x:hidden;box-shadow:0 0 0 1px rgba(0,0,0,.04)}}
.green-grad{{background:radial-gradient(circle at 55% 20%,#2e8c4d 0%,var(--green) 38%,var(--green-dark) 100%);color:white}}
.splash{{height:100dvh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px;position:relative}}.logo-circle{{width:150px;height:150px;border-radius:50%;background:#fff;border:8px solid var(--lime);display:grid;place-items:center;box-shadow:0 13px 28px rgba(0,0,0,.22);position:relative;overflow:hidden}}.logo-circle:before{{content:"";position:absolute;inset:92px -8px -8px;background:linear-gradient(135deg,#69a62d,#c7db22 45%,#2a8f3d 46%,#1f7335);transform:skewY(-18deg)}}.logo-inner{{position:relative;z-index:2;color:#1f783c;font-size:56px;font-weight:1000}}.brand-title{{font-size:40px;font-weight:1000;margin-top:23px;letter-spacing:.5px}}.brand-title span{{color:var(--lime)}}.progress{{position:absolute;bottom:112px;width:142px;height:7px;background:#07592f;border-radius:999px;overflow:hidden}}.progress i{{display:block;width:55%;height:100%;background:var(--lime)}}.splash-foot{{position:absolute;bottom:35px;text-align:center;font-size:13px;line-height:1.45;color:#fff}}
.login-wrap{{min-height:100dvh;background:#fff;padding:38px 24px 20px;display:flex;flex-direction:column;align-items:center}}.login-logo{{height:118px;display:flex;align-items:center}}.login-card{{width:100%;background:white;border:1px solid #e1e6eb;border-radius:9px;box-shadow:0 5px 14px rgba(0,0,0,.13);padding:18px}}.login-card h2{{text-align:center;font-size:12px;color:#7b8290;margin:0 0 16px}}.label{{font-size:10px;font-weight:800;color:#7d8594;margin:9px 0 5px}}.input{{width:100%;height:38px;border:1px solid #e1e6eb;border-radius:6px;padding:0 10px;font-weight:700}}.login-btn,.btn-green{{width:100%;height:42px;border:0;border-radius:7px;background:linear-gradient(90deg,#128953,#0ca060);color:white;font-weight:1000;font-size:14px;box-shadow:0 6px 12px rgba(0,120,60,.18)}}.login-btn{{margin-top:16px}}.link,.help{{text-align:center;color:#2182ff;font-size:10px;font-weight:800;margin-top:13px}}.help{{color:#8791a0;margin-top:auto}}
.top-green{{height:230px;border-radius:0 0 30px 30px;padding:10px 18px;background:radial-gradient(circle at 60% 15%,#2f8c4f 0%,var(--green) 36%,var(--green-dark) 100%);color:white;position:relative}}.pill-row{{display:flex;justify-content:space-between;align-items:center;margin-top:8px}}.pill{{background:white;color:var(--green);border-radius:24px;padding:8px 12px;font-weight:1000;font-size:13px;box-shadow:0 8px 18px rgba(0,0,0,.15)}}.pill:hover,.pill:active{{background:#eef9f1;transform:scale(.98)}}.avatar{{width:72px;height:72px;border-radius:50%;background:#fff;color:var(--green);display:grid;place-items:center;margin:18px auto 8px;font-size:48px;font-weight:900}}.user-name{{text-align:center;font-size:17px;font-weight:1000}}.welcome-card{{background:white;color:var(--ink);border-radius:18px;padding:14px 12px;text-align:center;box-shadow:var(--shadow);position:absolute;left:22px;right:22px;bottom:-36px}}.welcome-card b{{font-size:15px}}.welcome-card p{{margin:6px 0 0;color:#535b6b;font-size:14px}}.content{{padding:55px 19px 82px;min-height:calc(100dvh - 230px);position:relative;background:#fff}}.tile-grid{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-top:9px}}.tile{{height:112px;border-radius:16px;background:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;box-shadow:0 10px 23px rgba(0,0,0,.12);font-weight:1000;font-size:14px;border:1px solid #f0f1f2}}.tile .ico{{font-size:32px;color:var(--green)}}.tile:hover,.doc-row:hover,.mark-btn:hover{{background:#eef9f1;transform:translateY(-1px);transition:.12s}}.leaves{{width:136px;height:136px;margin:28px auto 26px;position:relative;opacity:.45}}.leaves i{{position:absolute;width:68px;height:112px;border-radius:70% 25% 70% 25%;filter:drop-shadow(0 4px 3px rgba(0,0,0,.12))}}.leaves i:nth-child(1){{background:#ffc9b3;left:14px;transform:rotate(26deg)}}.leaves i:nth-child(2){{background:#bdd2a4;left:56px;transform:rotate(26deg)}}.leaves i:nth-child(3){{background:#fff0a4;left:82px;top:32px;transform:rotate(26deg)}}.sync-card{{position:absolute;left:12px;right:12px;bottom:10px;background:#fff;border-radius:16px;padding:10px 13px;box-shadow:var(--shadow);display:grid;grid-template-columns:38px 1fr 38px;gap:8px;align-items:center}}.sync-btn{{width:36px;height:36px;border-radius:50%;background:var(--green);display:grid;place-items:center;color:white;font-size:24px}}.exit{{font-size:32px;color:#ff3030;text-align:center}}.sync-title{{font-weight:1000;font-size:13px}}.sync-sub{{color:#008a3f;font-weight:900;font-size:10px;margin-top:2px}}
.page-head{{height:132px;border-radius:0 0 28px 28px;background:radial-gradient(circle at 50% 0%,#2f8c4f 0%,var(--green) 48%,var(--green-dark) 100%);color:white;position:relative;text-align:center;padding:18px 16px}}.back{{position:absolute;left:17px;top:18px;color:white;font-size:30px;line-height:1;font-weight:300}}.back:hover{{color:var(--lime)}}.page-head h1{{font-size:22px;margin:44px 0 0;font-weight:1000}}.head-ico{{position:absolute;top:24px;left:50%;transform:translateX(-50%);color:var(--lime);font-size:28px}}
.att-page{{min-height:100dvh;background:#fff}}.hero-att{{height:214px;padding:24px 24px 0;text-align:center;background:radial-gradient(circle at 50% 0%,#159a8e 0%,#0b827b 44%,#006b3b 100%);color:white}}.hero-att .back{{top:16px}}.hero-att .cal{{font-size:30px;color:var(--lime);margin-bottom:8px}}.hero-att h1{{margin:0;font-size:25px;line-height:1.05;font-weight:1000}}.hero-att p{{font-size:14px;line-height:1.4;font-weight:700;margin:14px 6px 0}}.att-card{{margin:-38px 18px 0;background:white;border-radius:21px;padding:18px 15px 16px;box-shadow:var(--shadow);position:relative;z-index:2}}.sys{{font-size:11px;font-weight:1000;color:#009b4a;text-align:center}}.big-clock{{font-size:42px;font-weight:1000;text-align:center;letter-spacing:-1px;margin-top:8px;color:#070b22}}.date-line{{text-align:center;color:#979daf;font-size:13px;margin:5px 0 15px}}.mark-btn{{height:50px;border:0;border-radius:9px;background:#f1f3f5;color:#405064;font-weight:1000;font-size:15px;width:100%;cursor:pointer}}.btn-main{{background:#21b45d;color:white;font-size:17px}}.btn-row{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:12px 0}}.mark-btn.done{{background:#d9f8e2!important;color:#087a38!important}}.mark-btn.active{{background:#c9f3d6!important;outline:2px solid #37b96a}}.map-title{{display:flex;align-items:center;justify-content:space-between;margin-top:20px;color:#838da0;font-size:12px;font-weight:1000}}.signal{{background:#d8f8df;color:#0c923e;border-radius:18px;padding:8px 13px;font-size:11px;font-weight:1000}}.map{{height:142px;border-radius:14px;border:1px solid #d6dde6;margin-top:10px;position:relative;overflow:hidden;background:#f4f6f9}}.road-a,.road-b{{position:absolute;background:#fbd577}}.road-a{{width:420px;height:15px;left:-35px;top:50px;transform:rotate(38deg)}}.road-b{{width:320px;height:10px;left:-35px;top:92px;background:#cfd5dd;transform:rotate(-2deg)}}.road-c{{position:absolute;width:14px;height:220px;background:#cfd5dd;left:184px;top:-35px;transform:rotate(24deg)}}.pin{{position:absolute;left:50%;top:54%;transform:translate(-50%,-50%);font-size:48px;color:#2376ff;filter:drop-shadow(0 5px 4px rgba(0,0,0,.15))}}.bubble{{position:absolute;left:50%;top:42%;transform:translate(-50%,-50%);background:white;border-radius:7px;padding:8px 11px;box-shadow:0 5px 12px rgba(0,0,0,.16);font-size:12px;white-space:nowrap}}.loc-foot{{font-size:12px;color:#697386;font-weight:900;margin:11px 4px 0;display:flex;justify-content:space-between;gap:10px}}.loc-foot a{{color:#006fff;font-size:15px}}.selected-time{{margin-top:12px;height:44px;border-radius:10px;background:#f1f3f5;display:flex;align-items:center;justify-content:center;font-weight:1000;color:#405064;cursor:pointer}}.selected-time:hover{{background:#e3f6e9;color:#0b783d}}
.doc-list{{padding:23px 18px 82px;background:#fafafa;min-height:calc(100dvh - 132px)}}.doc-row{{height:73px;background:white;border:1px solid #ebeff2;border-radius:14px;margin-bottom:14px;display:grid;grid-template-columns:50px 1fr 18px;align-items:center;padding:0 14px;box-shadow:0 5px 14px rgba(0,0,0,.05)}}.doc-ico{{width:36px;height:36px;border:1px solid #58d2d0;border-radius:10px;display:grid;place-items:center;color:#00a8a6}}.doc-title{{font-size:16px;font-weight:1000}}.doc-sub{{font-size:12px;color:#999fb0;margin-top:5px}}.chev{{font-size:26px;color:#9aa2ad}}.bottom-nav{{position:absolute;left:0;right:0;bottom:0;height:58px;background:white;border-top:1px solid #e7eaee;display:flex;justify-content:space-around;align-items:center;color:#7b8491;font-size:11px;font-weight:800}}.bottom-nav a.active{{color:#009b8e}}.bottom-nav div{{font-size:22px;margin-bottom:2px;text-align:center}}
.config-card,.form-card{{margin:18px;background:#fff;border-radius:15px;box-shadow:var(--shadow);padding:15px}}.cfg-btn{{width:100%;min-height:48px;border-radius:8px;border:1px solid var(--green);font-size:16px;font-weight:1000;margin-bottom:10px;background:white;color:var(--green)}}.cfg-btn.primary{{background:var(--green);color:white}}.cfg-btn.disabled{{border-color:#9aa2ad;color:#6b7280}}.cfg-btn:hover{{background:#eaf7ef}}.cfg-btn.primary:hover{{background:#1d6539}}.form-title{{font-size:12px;font-weight:1000;color:var(--green);margin:4px 0 9px}}.field{{margin-bottom:9px}}.field label{{font-size:11px;font-weight:900;color:#333;display:block;margin-bottom:4px}}.field input,.field select{{width:100%;height:35px;border:1px solid #dde3e9;border-radius:6px;padding:0 9px;font-size:12px}}.two{{display:grid;grid-template-columns:1fr 1fr;gap:8px}}.ok-box{{background:#e3f5e8;color:#0c6d35;border-radius:10px;padding:13px;font-weight:800;font-size:13px;display:none}}.ok-box.show{{display:block}}.mini-help{{font-size:11px;color:#788292;line-height:1.35;margin-bottom:10px}}
.sign-box{{margin:18px;background:#fff;border:2px dashed #dbe3ef;border-radius:14px;height:190px;position:relative}}#canvas{{width:100%;height:100%;touch-action:none}}.mini-btn{{border:0;border-radius:10px;background:#f1f4f7;height:42px;font-weight:900;color:#39424e}}.danger{{background:#feecec;color:#e93030}}.save{{background:#22b45e;color:#fff}}.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:0 18px 12px}}.grid1{{margin:0 18px 12px}}.grid1 button{{width:100%}}.profile-card{{margin:18px;border-radius:20px;background:white;box-shadow:var(--shadow);padding:22px;text-align:center}}.profile-photo{{width:92px;height:92px;border-radius:50%;background:#edf4ef;color:var(--green);display:grid;place-items:center;font-size:60px;margin:0 auto 14px}}.kv{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:18px;text-align:left}}.kv div{{background:#f6f8f8;border-radius:13px;padding:13px}}.kv small{{display:block;color:#8791a0;font-weight:800}}.kv b{{display:block;margin-top:5px;font-size:12px}}
.toast{{position:fixed;left:50%;top:14px;transform:translateX(-50%);background:#10251a;color:white;border-radius:999px;padding:11px 18px;font-weight:900;z-index:99;box-shadow:var(--shadow);font-size:13px;white-space:nowrap}}
.wheel-modal{{position:fixed;inset:auto 0 0;z-index:50;background:rgba(0,0,0,.25);display:none;align-items:end;justify-content:center}}.wheel-modal.show{{display:flex}}.wheel-card{{width:100%;max-width:390px;background:white;border-radius:22px 22px 0 0;padding:14px 18px 22px;box-shadow:0 -12px 35px rgba(0,0,0,.22)}}.wheel-head{{display:flex;justify-content:space-between;align-items:center;font-weight:1000;margin-bottom:8px}}.wheel-head button{{border:0;background:white;color:#0aa25b;font-weight:1000;font-size:16px}}.wheels{{height:160px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;position:relative}}.wheels:before{{content:"";position:absolute;left:0;right:0;top:58px;height:44px;background:rgba(0,0,0,.04);border-radius:12px;pointer-events:none}}.wheel{{height:160px;overflow-y:auto;scroll-snap-type:y mandatory;text-align:center;padding:58px 0;scrollbar-width:none}}.wheel::-webkit-scrollbar{{display:none}}.wheel div{{height:44px;line-height:44px;font-size:22px;color:#b6bdc7;scroll-snap-align:center;font-weight:500}}.wheel div.sel{{color:#303642;font-size:26px}}.err{{background:#fee2e2;color:#991b1b;border-radius:8px;padding:10px 12px;font-size:12px;margin-bottom:10px;font-weight:800}}
@media(min-width:800px){{.phone{{margin:10px 0;min-height:780px;border-radius:20px}}}}

/* ===== V3 OMAR: compacto estilo referencia, sin tareo, botones no quedan pintados ===== */
.phone{{max-width:390px!important;min-height:100dvh!important;box-shadow:0 0 0 1px rgba(0,0,0,.05)!important}}
@media(min-width:800px){{body{{background:#fff!important}}.phone{{margin:8px auto!important;border-radius:0!important;min-height:100dvh!important;box-shadow:0 0 0 1px #e1e5e8!important}}}}
.top-green{{height:214px!important;border-radius:0 0 28px 28px!important;padding:8px 14px!important}}
.pill-row{{margin-top:5px!important}}.pill{{font-size:12px!important;padding:7px 11px!important;border-radius:20px!important}}
.avatar{{width:68px!important;height:68px!important;margin:16px auto 8px!important;font-size:43px!important}}.user-name{{font-size:16px!important}}
.welcome-card{{left:20px!important;right:20px!important;bottom:-32px!important;border-radius:16px!important;padding:12px!important}}
.content{{padding:49px 18px 76px!important}}.tile-grid{{gap:14px!important}}.tile{{height:96px!important;border-radius:13px!important;font-size:12px!important;gap:8px!important}}.tile .ico{{font-size:27px!important}}
.leaves{{width:112px!important;height:112px!important;margin:20px auto!important}}
.sync-card{{bottom:8px!important;border-radius:13px!important;padding:8px 10px!important;grid-template-columns:34px 1fr 32px!important}}.sync-btn{{width:32px!important;height:32px!important;font-size:20px!important}}.sync-title{{font-size:12px!important}}.sync-sub{{font-size:9px!important}}.exit{{font-size:26px!important}}
.hero-att{{height:184px!important;padding:17px 22px 0!important}}.hero-att .cal{{font-size:25px!important;margin-bottom:5px!important}}.hero-att h1{{font-size:23px!important}}.hero-att p{{font-size:13px!important;margin-top:10px!important}}
.att-card{{margin:-33px 15px 0!important;border-radius:20px!important;padding:17px 15px 14px!important}}.big-clock{{font-size:39px!important}}.date-line{{font-size:12px!important;margin-bottom:13px!important}}
.mark-btn{{height:48px!important;border-radius:9px!important;background:#f0f2f4!important;color:#405064!important;transition:.15s!important}}
.mark-btn:hover,.mark-btn.active{{background:#21b45d!important;color:#fff!important;outline:none!important;transform:translateY(-1px)!important}}
.mark-btn.done{{background:#f0f2f4!important;color:#405064!important}}
.btn-main{{background:#f0f2f4!important;color:#405064!important}}.btn-main:hover,.btn-main.active{{background:#21b45d!important;color:#fff!important}}
.btn-row{{gap:11px!important;margin:11px 0!important}}.selected-time{{height:44px!important;margin-top:11px!important}}
.map-title{{margin-top:18px!important}}.map{{height:137px!important}}.loc-foot{{font-size:11px!important;align-items:flex-start}}.loc-foot a{{font-size:13px!important;font-weight:1000}}
.page-head{{height:122px!important;border-radius:0 0 27px 27px!important}}.page-head h1{{font-size:20px!important;margin-top:43px!important}}.back{{font-size:29px!important;top:15px!important}}
.doc-list{{padding:22px 17px 76px!important}}.doc-row{{height:70px!important;border-radius:13px!important;margin-bottom:12px!important}}.doc-title{{font-size:15px!important}}.doc-sub{{font-size:11px!important}}
.config-card,.form-card,.profile-card{{margin:15px!important;border-radius:14px!important}}.cfg-btn{{min-height:45px!important;font-size:15px!important;margin-bottom:9px!important}}
.login-wrap{{padding:28px 24px 20px!important}}.login-logo{{height:106px!important}}.brand-title{{font-size:36px!important}}
.sign-box{{height:180px!important}}.bottom-nav{{height:56px!important}}

"""


CSS += """
/* ===== V4 P&A visual: verde referencia, login verde, iconos SVG limpios ===== */
:root{--green:#28763A!important;--green-dark:#1f5f31!important;--lime:#b8d81b!important;--ink:#071126!important;--shadow:0 12px 28px rgba(0,0,0,.14)!important}
html,body{background:#fff!important;font-family:'Segoe UI',Arial,sans-serif!important;font-weight:700!important}
.phone{max-width:390px!important;background:#fff!important;border-left:1px solid #dfe7e3!important;border-right:1px solid #dfe7e3!important;box-shadow:none!important;overflow-x:hidden!important}
.svgico{display:inline-grid;place-items:center;width:1.15em;height:1.15em;vertical-align:-.17em}.svgico svg{width:100%;height:100%;fill:currentColor;display:block}.pill .svgico,.login-btn .svgico{margin-right:5px}.back .svgico{width:26px;height:26px}.head-ico .svgico{width:30px;height:30px}.tile .svgico{width:34px;height:34px}.bottom-nav .svgico{width:23px;height:23px}.doc-ico .svgico{width:19px;height:19px}.exit .svgico{width:29px;height:29px}.pin .svgico{width:46px;height:46px;color:#1f76ff;filter:drop-shadow(0 5px 4px rgba(0,0,0,.16))}
.login-pro{padding:0 22px 20px!important;background:#fff!important;position:relative;min-height:100dvh!important}.login-green{height:270px;margin:0 -22px 0;background:#28763A;color:#fff;text-align:center;padding:18px 22px 0}.mini-top{display:flex;justify-content:space-between}.mini-top button{border:0;background:transparent;color:#fff;font-weight:1000;font-size:12px}.mini-top .svgico{width:14px;height:14px;margin-right:3px}.login-mark{width:110px;height:110px;margin:35px auto 12px;border-radius:50%;background:#fff;border:7px solid var(--lime);box-shadow:0 8px 18px rgba(0,0,0,.22);display:grid;place-items:center;color:#28763A}.login-mark div{font-size:33px;font-weight:1000}.login-mark span{color:var(--lime)}.login-green h1{margin:0;font-size:22px;text-transform:uppercase;font-weight:1000;letter-spacing:.3px}.login-green p{margin:5px 0 0;font-size:12px;font-weight:1000}.login-float{margin:-36px 0 0!important;border-radius:12px!important;box-shadow:0 14px 35px rgba(0,0,0,.15)!important;border:0!important}.login-card h2{font-size:13px!important;color:#6e7787!important;font-weight:1000!important}.input{height:43px!important;border-radius:9px!important;font-size:15px!important}.label{font-size:11px!important;color:#737b8c!important}.login-btn{height:48px!important;background:#28763A!important;border-radius:9px!important;font-size:17px!important;display:flex;align-items:center;justify-content:center}.link{font-size:11px!important;color:#006fff!important}.help{font-size:11px!important;color:#7c8797!important;margin-top:auto!important;padding-bottom:10px!important}
.top-green,.page-head{background:#28763A!important;background-image:none!important}.top-green{height:224px!important;border-radius:0!important}.pill{background:transparent!important;color:#fff!important;box-shadow:none!important;padding:6px 2px!important;font-size:12px!important}.pill:hover{color:var(--lime)!important;background:transparent!important}.avatar-pa{background:#fff!important;color:#28763A!important;border:7px solid #fff!important}.avatar-pa .svgico{width:45px;height:45px}.welcome-card{bottom:-28px!important;border-radius:9px!important;min-height:46px!important;padding:13px!important}.welcome-card p{display:none!important}.content{padding-top:47px!important}.tile{height:98px!important;border-radius:9px!important;color:#00122b!important;box-shadow:0 12px 28px rgba(0,0,0,.10)!important}.tile:hover{background:#eaf7ef!important}.tile .ico{color:#28763A!important}.tile-grid{gap:14px!important}.leaves{opacity:.22!important;margin-top:18px!important}.sync-card{border-radius:10px!important}.sync-btn{background:#28763A!important}.exit{color:#ff3434!important}.sync-sub{color:#087e42!important}.brand-title{font-size:40px!important}
.hero-att{background:#0b8f84!important;background:linear-gradient(180deg,#108b84 0%,#087447 100%)!important;height:184px!important}.hero-att .cal{color:var(--lime)!important}.hero-att h1{font-size:22px!important}.att-card{border-radius:18px!important}.sys{color:#008e42!important}.big-clock{font-size:40px!important;color:#05091f!important}.mark-btn{border-radius:9px!important;font-weight:1000!important}.mark-btn:hover,.mark-btn.active{background:#28763A!important;color:#fff!important}.btn-main{background:#22b45e!important;color:#fff!important}.btn-main:hover{background:#28763A!important}.signal{border:2px solid #28763A!important;background:#d8f9df!important;color:#087b3e!important}.map{border-radius:13px!important}.selected-time .svgico{margin-right:8px;color:#7a8290}
.page-head{height:126px!important;border-radius:0 0 28px 28px!important}.page-head h1{font-size:20px!important;color:#fff!important;text-shadow:0 1px 1px rgba(0,0,0,.08)}.head-ico{color:var(--lime)!important}.doc-list{background:#f7f8f7!important}.doc-row{height:72px!important;border-radius:13px!important;box-shadow:0 5px 13px rgba(0,0,0,.045)!important}.doc-row:hover{background:#eaf7ef!important}.doc-ico{color:#00a9a5!important;border-color:#59d5d1!important}.doc-title{font-weight:1000!important;color:#03112b!important}.doc-sub{font-weight:500!important}.chev{font-size:28px!important}.bottom-nav{position:fixed!important;max-width:390px;margin:0 auto;height:58px!important;color:#7e8794!important}.bottom-nav a.active{color:#008f82!important}.bottom-nav div{height:25px!important}.config-card,.form-card,.profile-card{box-shadow:0 14px 30px rgba(0,0,0,.10)!important}.cfg-btn.primary,.btn-green,.save{background:#28763A!important}.cfg-btn{color:#28763A!important;border-color:#28763A!important}.cfg-btn.primary{color:#fff!important}.profile-photo{color:#28763A!important}.sign-box{border-color:#d8e4dc!important}
@media(min-width:800px){.phone{min-height:100dvh!important;border-radius:0!important;margin:0 auto!important}.login-green{height:270px!important}}
"""


CSS += """
/* ===== V5 FINAL: estética Tareo Móvil adaptada a P&A, sin módulos de tareo ===== */
:root{--green:#2f773b!important;--green-dark:#23713b!important;--lime:#a8d32a!important;--ink:#00112b!important;--shadow:0 8px 18px rgba(0,0,0,.15)!important}
html,body{background:#fff!important;font-family:Inter,Segoe UI,Arial,sans-serif!important;font-weight:800!important}.phone{max-width:390px!important;min-height:100dvh!important;background:#fff!important;border-left:1px solid #dfe7e3!important;border-right:1px solid #dfe7e3!important;overflow-x:hidden!important;box-shadow:none!important}
.login-pro{padding:0 20px 22px!important;background:#fff!important;min-height:100dvh!important;position:relative!important;overflow:hidden!important}.login-green{height:292px!important;margin:0 -20px!important;background:#2f773b!important;color:#fff!important;text-align:center!important;padding:15px 20px 0!important;border-radius:0!important;box-shadow:none!important}.mini-top{display:flex!important;justify-content:space-between!important;align-items:center!important}.mini-top button{border:0!important;background:transparent!important;color:#fff!important;font-weight:1000!important;font-size:12px!important;padding:0!important}.mini-top .svgico{width:14px!important;height:14px!important;margin-right:3px!important}.login-mark{width:112px!important;height:112px!important;border-radius:999px!important;background:#fff!important;border:7px solid var(--lime)!important;margin:40px auto 14px!important;display:grid!important;place-items:center!important;box-shadow:0 4px 12px rgba(0,0,0,.26)!important;color:#2f773b!important}.login-mark .svgico{width:54px!important;height:54px!important}.login-green h1{font-size:21px!important;margin:0!important;font-weight:1000!important;text-transform:uppercase!important;letter-spacing:.3px!important}.login-green p{font-size:12px!important;margin:5px 0 0!important;font-weight:1000!important;text-transform:uppercase!important}.login-float{margin:-36px 0 0!important;border-radius:11px!important;border:0!important;box-shadow:0 13px 30px rgba(0,0,0,.15)!important;padding:14px 15px 18px!important}.login-card h2{font-size:13px!important;color:#6e7787!important;font-weight:1000!important}.input{height:42px!important;border-radius:9px!important;font-size:15px!important}.login-btn{height:47px!important;border-radius:9px!important;background:#2f773b!important;color:#fff!important;font-size:18px!important;text-transform:uppercase!important}.login-mini-tiles{display:flex!important;gap:18px!important;justify-content:center!important;margin:30px auto 0!important}.mini-login-tile{width:76px!important;height:76px!important;border-radius:8px!important;background:#fff!important;box-shadow:0 8px 18px rgba(0,0,0,.15)!important;color:#2f773b!important;display:flex!important;flex-direction:column!important;align-items:center!important;justify-content:center!important;font-size:10px!important;font-weight:1000!important;text-align:center!important}.mini-login-tile .svgico{width:27px!important;height:27px!important;margin-bottom:5px!important}.login-leaf{width:83px!important;height:112px!important;border-radius:70% 30% 70% 30%!important;background:linear-gradient(135deg,#fff9c6,#e7eccd,#ccdcb1)!important;opacity:.62!important;margin:18px auto 0!important;transform:rotate(-18deg)!important}
.top-green{height:232px!important;background:#2f773b!important;border-radius:0!important;padding:11px 14px 0!important;color:#fff!important}.pill{background:transparent!important;color:white!important;box-shadow:none!important;padding:5px 0!important;font-size:12px!important}.pill:hover{color:var(--lime)!important}.avatar-pa{width:82px!important;height:82px!important;margin:12px auto 7px!important;border:0!important;background:white!important;color:#2f773b!important;box-shadow:0 8px 18px rgba(0,0,0,.13)!important}.avatar-pa .svgico{width:52px!important;height:52px!important}.user-name{font-size:15px!important;font-weight:1000!important;text-transform:uppercase!important}.welcome-card{left:18px!important;right:18px!important;bottom:-23px!important;border-radius:9px!important;min-height:45px!important;padding:12px!important;box-shadow:0 8px 18px rgba(0,0,0,.16)!important}.welcome-card b{font-size:13px!important}.welcome-card p{display:none!important}.content{padding:50px 16px 78px!important}.tile-grid{grid-template-columns:1fr 1fr!important;gap:14px!important;margin-top:0!important}.tile{height:92px!important;border-radius:9px!important;box-shadow:0 8px 18px rgba(0,0,0,.13)!important;font-size:12px!important;color:#00112b!important;gap:7px!important}.tile .svgico{width:33px!important;height:33px!important}.tile:hover{background:#eaf5eb!important;transform:translateY(-1px)!important}.leaves{width:112px!important;height:112px!important;margin:22px auto 18px!important;opacity:.40!important}.sync-card{bottom:8px!important;border-radius:10px!important;padding:8px 10px!important;box-shadow:0 8px 18px rgba(0,0,0,.15)!important}.sync-btn{background:#2f773b!important}.sync-title{font-size:12px!important}.sync-sub{font-size:9px!important;color:#16823c!important}.exit{color:#e83434!important}.page-head{height:128px!important;background:#2f773b!important;border-radius:0 0 28px 28px!important}.page-head h1{font-size:21px!important;line-height:1.08!important;margin-top:42px!important;text-transform:none!important}.head-ico{color:var(--lime)!important}.back{top:15px!important;left:16px!important;color:white!important}.back .svgico{width:25px!important;height:25px!important}
.doc-list{background:#f7f8f7!important;padding:22px 17px 82px!important}.doc-row{height:72px!important;border-radius:13px!important;margin-bottom:13px!important;box-shadow:0 5px 13px rgba(0,0,0,.055)!important}.doc-title{font-size:15px!important;color:#00112b!important;font-weight:1000!important}.doc-sub{font-size:11px!important;font-weight:500!important}.doc-ico{border-color:#59d5d1!important;color:#009f9d!important}.chev{font-size:27px!important;color:#929aa6!important}.hero-att{background:#0d887f!important;background:linear-gradient(180deg,#108b84 0%,#0a7147 100%)!important;height:185px!important}.hero-att h1{font-size:22px!important}.hero-att p{font-size:13px!important}.att-card{margin:-33px 15px 0!important;border-radius:18px!important;box-shadow:0 9px 23px rgba(0,0,0,.14)!important}.btn-main{background:#23b662!important;color:#fff!important}.mark-btn{border-radius:9px!important}.mark-btn:hover,.mark-btn.active{background:#2f773b!important;color:#fff!important}.mark-btn.done{background:#f0f2f4!important;color:#405064!important}.signal{background:#d9f8df!important;color:#077d3a!important;border:2px solid #2f773b!important}.selected-time{height:44px!important}.bottom-nav{position:fixed!important;left:50%!important;transform:translateX(-50%)!important;width:100%!important;max-width:390px!important;height:58px!important;background:#fff!important;border-top:1px solid #e5ece7!important;color:#73809a!important;z-index:40!important}.bottom-nav a.active{color:#008f82!important}.bottom-nav .svgico{width:23px!important;height:23px!important}.config-card,.form-card,.profile-card{box-shadow:0 12px 28px rgba(0,0,0,.12)!important;border-radius:13px!important}.cfg-btn.primary,.btn-green,.save{background:#2f773b!important;color:white!important}.cfg-btn{border-color:#2f773b!important;color:#2f773b!important}.profile-photo{color:#2f773b!important}.sign-box{border-color:#dce8df!important}.stats-row{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:0 0 14px}.stat-card{background:#f7fbf8;border:1px solid #e1ece4;border-radius:10px;padding:10px;text-align:center;color:#2f773b;font-weight:1000}.stat-card b{display:block;color:#00112b;font-size:18px;margin-top:3px}.history-card{margin-top:14px;background:#fbfffb;border:1px solid #dcebdd;border-radius:12px;padding:10px}.history-row{display:flex;justify-content:space-between;border-bottom:1px solid #e7f1e7;padding:8px 2px;font-size:12px}.history-row:last-child{border-bottom:0}.history-row b{color:#2f773b}.pdf-badge{display:inline-flex;gap:6px;align-items:center;background:#eaf5eb;border:1px solid #cfe6d4;border-radius:999px;padding:5px 9px;font-size:11px;color:#2f773b;font-weight:1000}@media(min-width:800px){body{background:#fff!important}.phone{margin:0 auto!important;border-radius:0!important;min-height:100dvh!important}}
"""


CSS += """
/* ===== V6 DEFINITIVO: distribución tipo Tareo Móvil en todo el lienzo ===== */
html,body{background:#fff!important;min-height:100dvh!important;overflow-x:hidden!important}.phone{width:390px!important;max-width:390px!important;margin:0 auto!important;min-height:100dvh!important;border:0!important;box-shadow:none!important;background:#fff!important}.login-pro{width:390px!important;max-width:390px!important;margin:0 auto!important;padding:0 20px 22px!important}.login-green{height:270px!important;margin:0 -20px!important;background:#2f773b!important}.login-mark{width:112px!important;height:112px!important;margin-top:38px!important}.login-float{margin-top:-32px!important}.login-mini-tiles{margin-top:26px!important}.mini-login-tile{width:76px!important;height:76px!important}.login-leaf{margin-top:10px!important}.help{position:fixed!important;left:50%!important;transform:translateX(-50%)!important;bottom:18px!important;width:390px!important;max-width:390px!important;text-align:center!important}.top-green{background:#2f773b!important;height:210px!important;border-radius:0 0 16px 16px!important}.content{min-height:calc(100dvh - 210px)!important}.sync-card{position:fixed!important;left:50%!important;right:auto!important;transform:translateX(-50%)!important;width:370px!important;max-width:calc(100vw - 20px)!important;bottom:8px!important}.exit{position:fixed!important;right:12px!important;bottom:12px!important;z-index:80!important}.page-head,.hero-att{background:#2f773b!important;background-image:none!important}.bottom-nav{left:50%!important;right:auto!important;transform:translateX(-50%)!important}.mass-preview{max-height:180px;overflow:auto;background:white;border:1px solid #e5eee6;border-radius:9px;padding:8px;font-size:11px}.mass-preview table{width:100%;border-collapse:collapse}.mass-preview th,.mass-preview td{border-bottom:1px solid #eef3ef;padding:5px;text-align:left}.mass-ok{background:#e7f8eb;color:#166534;border:1px solid #bfe7c7;border-radius:10px;padding:10px;font-size:12px;font-weight:900;margin-bottom:10px}.mass-bad{background:#fee2e2;color:#991b1b;border:1px solid #fecaca;border-radius:10px;padding:10px;font-size:12px;font-weight:900;margin-bottom:10px}
@media(min-width:800px){body{display:block!important}.phone{margin:0 auto!important}.login-pro{margin:0 auto!important}.sync-card{left:10px!important;transform:none!important;width:330px!important;bottom:10px!important}.login-pro .help{left:50%!important}.login-pro:after{content:'';position:fixed;right:16px;bottom:13px;width:20px;height:20px;border:2px solid #ef4444;border-left:0;border-radius:2px}.login-pro:before{content:'➜';position:fixed;right:13px;bottom:8px;color:#ef4444;font-size:24px;font-weight:900;z-index:1}}
"""


CSS += """
/* ===== V7 CLÁSICO: P&A con apariencia Tareo Móvil real, compacto y centrado ===== */
:root{
  --green:#2f773b!important;
  --green-dark:#286b34!important;
  --lime:#a7d129!important;
  --ink:#05200c!important;
  --shadow:0 6px 16px rgba(0,0,0,.16)!important;
}
html,body{
  width:100%!important;
  min-height:100dvh!important;
  margin:0!important;
  background:#fff!important;
  overflow-x:hidden!important;
  font-family:Inter,Segoe UI,Arial,sans-serif!important;
  font-weight:800!important;
}
body{
  display:block!important;
  justify-content:initial!important;
}
.phone{
  width:390px!important;
  max-width:390px!important;
  min-height:100dvh!important;
  margin:0 auto!important;
  background:#fff!important;
  border-left:1px solid #e1e7e1!important;
  border-right:1px solid #e1e7e1!important;
  box-shadow:none!important;
  overflow-x:hidden!important;
  position:relative!important;
}
@media(max-width:430px){
  .phone{width:100%!important;max-width:100%!important;border-left:0!important;border-right:0!important}
}

/* Splash clásico */
.splash.green-grad,.green-grad{
  background:#23773f!important;
  background-image:radial-gradient(circle at 50% 36%,#2d8547 0%,#23773f 44%,#156430 100%)!important;
}
.splash{height:100dvh!important;justify-content:center!important}
.logo-circle{
  width:145px!important;height:145px!important;border:6px solid #92bd33!important;
  box-shadow:0 4px 12px rgba(0,0,0,.25)!important;
}
.logo-inner{font-size:45px!important;color:#23773f!important}
.brand-title{font-size:24px!important;color:#fff!important;text-transform:uppercase!important;margin-top:18px!important}
.splash-foot{font-size:11px!important;color:#e3f4e5!important;bottom:26px!important}

/* Login igual a referencia: verde arriba, tarjeta flotante, letras pequeñas */
.login-pro{
  width:390px!important;max-width:390px!important;margin:0 auto!important;
  padding:0 18px 20px!important;background:#fff!important;min-height:100dvh!important;overflow:hidden!important;
}
.login-green{
  height:260px!important;margin:0 -18px!important;background:#2f773b!important;background-image:none!important;
  border-radius:0!important;padding:16px 16px 0!important;text-align:center!important;color:#fff!important;
}
.mini-top{display:flex!important;justify-content:space-between!important;align-items:center!important}
.mini-top button{font-size:11px!important;font-weight:1000!important;color:#fff!important;background:transparent!important;border:0!important;padding:0!important}
.mini-top .svgico{width:13px!important;height:13px!important;margin-right:3px!important}
.login-mark{
  width:108px!important;height:108px!important;border-radius:999px!important;background:#fff!important;
  border:6px solid #92bd33!important;margin:36px auto 13px!important;color:#23773f!important;
  box-shadow:0 4px 12px rgba(0,0,0,.24)!important;
}
.login-mark .svgico{width:52px!important;height:52px!important}
.login-green h1{font-size:20px!important;font-weight:1000!important;letter-spacing:.2px!important;margin:0!important;text-transform:uppercase!important}
.login-green p{font-size:11px!important;font-weight:1000!important;margin:4px 0 0!important;text-transform:uppercase!important}
.login-float{
  margin:-28px auto 0!important;border-radius:9px!important;padding:12px 14px 16px!important;border:0!important;
  box-shadow:0 10px 24px rgba(0,0,0,.16)!important;background:#fff!important;
}
.login-card h2{font-size:12px!important;color:#2f2f2f!important;margin-bottom:13px!important;font-weight:900!important}
.label{font-size:10px!important;color:#334155!important;margin:8px 0 4px!important;font-weight:900!important}
.input{
  height:36px!important;border-radius:8px!important;border:1px solid #dce7dc!important;
  background:#fff!important;font-size:13px!important;color:#0f172a!important;font-weight:900!important;
}
.login-btn{
  height:41px!important;border-radius:8px!important;background:#2f773b!important;
  font-size:15px!important;text-transform:uppercase!important;box-shadow:none!important;margin-top:14px!important;
}
.link{font-size:10px!important;margin-top:11px!important}
.login-mini-tiles{
  display:flex!important;justify-content:center!important;gap:18px!important;margin:26px auto 0!important;
}
.mini-login-tile{
  width:72px!important;height:70px!important;border-radius:8px!important;background:#fff!important;color:#2f773b!important;
  box-shadow:0 7px 17px rgba(0,0,0,.14)!important;font-size:9px!important;font-weight:1000!important;
  display:flex!important;flex-direction:column!important;align-items:center!important;justify-content:center!important;text-align:center!important;
}
.mini-login-tile .svgico{width:25px!important;height:25px!important;margin-bottom:5px!important}
.login-leaf{
  width:78px!important;height:104px!important;margin:12px auto 0!important;opacity:.55!important;
  border-radius:70% 30% 70% 30%!important;background:linear-gradient(135deg,#fff9c5,#e9edca,#cad9b0)!important;transform:rotate(-18deg)!important;
}
.help{
  position:fixed!important;left:50%!important;bottom:15px!important;transform:translateX(-50%)!important;
  width:390px!important;max-width:100%!important;text-align:center!important;font-size:10px!important;color:#6b7280!important;padding:0!important;
}

/* Home clásico */
.top-green{
  height:205px!important;background:#2f773b!important;background-image:none!important;border-radius:0 0 16px 16px!important;
  padding:10px 14px 0!important;color:#fff!important;
}
.pill-row{margin-top:0!important}
.pill{background:transparent!important;color:#fff!important;box-shadow:none!important;font-size:11px!important;font-weight:1000!important;padding:0!important;border-radius:0!important}
.pill .svgico{width:13px!important;height:13px!important;margin-right:3px!important}
.avatar-pa{
  width:78px!important;height:78px!important;border-radius:999px!important;background:#fff!important;color:#2f773b!important;
  margin:18px auto 8px!important;box-shadow:0 5px 13px rgba(0,0,0,.16)!important;border:0!important;
}
.avatar-pa .svgico{width:48px!important;height:48px!important}
.user-name{font-size:13px!important;font-weight:1000!important;color:#fff!important;text-transform:uppercase!important}
.welcome-card{
  left:18px!important;right:18px!important;bottom:-22px!important;min-height:42px!important;border-radius:9px!important;
  box-shadow:0 6px 15px rgba(0,0,0,.16)!important;padding:10px!important;background:#fff!important;
}
.welcome-card b{font-size:12px!important;color:#0f172a!important}
.welcome-card p{display:none!important}
.content{
  min-height:calc(100dvh - 205px)!important;padding:51px 17px 84px!important;background:#fff!important;
}
.stats-row{display:none!important}
.tile-grid{display:grid!important;grid-template-columns:1fr 1fr!important;gap:14px!important;margin-top:0!important}
.tile{
  height:86px!important;border-radius:8px!important;background:#fff!important;border:0!important;
  box-shadow:0 7px 17px rgba(0,0,0,.14)!important;font-size:10px!important;font-weight:1000!important;color:#00112b!important;gap:7px!important;
}
.tile .ico{font-size:0!important;color:#2f773b!important}.tile .svgico{width:30px!important;height:30px!important}
.tile:hover{background:#f7fff8!important;transform:none!important}
.leaves{
  width:108px!important;height:108px!important;margin:22px auto 18px!important;opacity:.38!important;
}
.sync-card{
  position:fixed!important;left:10px!important;right:auto!important;bottom:10px!important;transform:none!important;
  width:330px!important;max-width:calc(100vw - 66px)!important;background:#fff!important;border-radius:8px!important;
  padding:7px 10px!important;box-shadow:0 4px 12px rgba(0,0,0,.16)!important;grid-template-columns:32px 1fr!important;
}
.sync-btn{width:30px!important;height:30px!important;background:#2f773b!important;border:2px solid #0d421e!important;font-size:18px!important}
.sync-title{font-size:10px!important;color:#123d1f!important;font-weight:1000!important}
.sync-sub{font-size:8.5px!important;color:#16823c!important;font-weight:1000!important}
.exit{
  position:fixed!important;right:12px!important;bottom:12px!important;z-index:80!important;color:#ef4444!important;font-size:23px!important;
}
.exit .svgico{width:24px!important;height:24px!important}

/* Páginas internas compactas tipo Tareo */
.page-head{
  height:125px!important;background:#2f773b!important;background-image:none!important;border-radius:0 0 16px 16px!important;
  padding:12px!important;
}
.page-head h1{font-size:13px!important;margin-top:65px!important;color:#fff!important;font-weight:1000!important;text-transform:uppercase!important}
.head-ico{top:30px!important;color:#fff!important}.head-ico .svgico{width:30px!important;height:30px!important}
.back{left:12px!important;top:12px!important;color:#fff!important}.back .svgico{width:20px!important;height:20px!important}
.doc-list,.form-card,.config-card,.profile-card{font-size:11px!important}
.doc-list{padding:14px 9px 76px!important;background:#fff!important}
.doc-row{height:60px!important;border-radius:8px!important;margin:8px 0!important;box-shadow:0 3px 11px rgba(0,0,0,.10)!important;grid-template-columns:42px 1fr 14px!important;padding:0 10px!important}
.doc-ico{width:32px!important;height:32px!important;border-radius:7px!important;color:#2f773b!important;border-color:#d9eadc!important}
.doc-title{font-size:12px!important;color:#00112b!important}.doc-sub{font-size:9.5px!important;color:#6b7280!important}
.chev{font-size:22px!important}
.hero-att{
  height:142px!important;background:#2f773b!important;background-image:none!important;padding:12px 18px 0!important;
  border-radius:0 0 16px 16px!important;
}
.hero-att .cal{font-size:0!important;margin:15px 0 6px!important}.hero-att .svgico{width:28px!important;height:28px!important;color:#fff!important}
.hero-att h1{font-size:14px!important;line-height:1.1!important;text-transform:uppercase!important}.hero-att p{display:none!important}
.att-card{margin:-18px 9px 70px!important;border-radius:9px!important;padding:10px!important;box-shadow:0 4px 13px rgba(0,0,0,.16)!important}
.sys{font-size:9px!important}.big-clock{font-size:30px!important;margin-top:5px!important}.date-line{font-size:10px!important;margin-bottom:9px!important}
.mark-btn{height:42px!important;border-radius:7px!important;font-size:12px!important}
.btn-row{gap:7px!important;margin:8px 0!important}.selected-time{height:38px!important;font-size:11px!important}.map-title{font-size:9px!important;margin-top:12px!important}.signal{font-size:9px!important;padding:5px 9px!important}.map{height:105px!important}.loc-foot{font-size:9px!important}.history-card{font-size:10px!important;margin-top:8px!important}
.config-card,.form-card,.profile-card{
  margin:10px 8px 74px!important;border-radius:9px!important;padding:10px!important;box-shadow:0 5px 14px rgba(0,0,0,.13)!important;
}
.cfg-btn{min-height:38px!important;border-radius:7px!important;font-size:12px!important;margin-bottom:8px!important}
.field label{font-size:9px!important}.field input,.field select{height:32px!important;font-size:11px!important;border-radius:7px!important}
.form-title{font-size:10px!important}.mini-help{font-size:9.5px!important}
.profile-photo{width:70px!important;height:70px!important;font-size:40px!important}.profile-card h2{font-size:15px!important}.profile-card p{font-size:11px!important}
.kv{grid-template-columns:1fr!important;gap:7px!important}.kv div{padding:9px!important;border-radius:8px!important}.kv small{font-size:9px!important}.kv b{font-size:10px!important}
.sign-box{height:160px!important;margin:10px 8px!important}.mini-btn{height:36px!important;font-size:11px!important}.grid2{gap:7px!important;margin:0 8px 8px!important}.grid1{margin:0 8px 8px!important}
.bottom-nav{
  position:fixed!important;left:50%!important;right:auto!important;bottom:0!important;transform:translateX(-50%)!important;
  width:390px!important;max-width:100%!important;height:52px!important;background:#fff!important;border-top:1px solid #e6ece6!important;
  color:#477b4d!important;font-size:9px!important;font-weight:900!important;z-index:70!important;
}
.bottom-nav .svgico{width:18px!important;height:18px!important}.bottom-nav a.active{color:#2f773b!important}
.mass-preview{font-size:9px!important;max-height:130px!important}.mass-ok,.mass-bad{font-size:10px!important;padding:8px!important}
@media(min-width:800px){
  .phone{margin:0 auto!important}
  .sync-card{left:10px!important}
  .exit{right:12px!important}
}
"""

JS = """
function tone(ok=true){try{const C=window.AudioContext||window.webkitAudioContext;const ctx=new C();const o=ctx.createOscillator();const g=ctx.createGain();o.type='sine';o.frequency.value=ok?880:220;g.gain.value=.07;o.connect(g);g.connect(ctx.destination);o.start();setTimeout(()=>{o.stop();ctx.close()},150)}catch(e){}}
function showToast(msg,ok=true){tone(ok);let t=document.createElement('div');t.className='toast';t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.remove(),2400)}
function updateClock(){let el=document.getElementById('liveClock'); if(!el)return; let d=new Date(); let h=d.getHours(); let am=h<12?'a. m.':'p. m.'; h=h%12||12; el.textContent=String(h).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0')+' '+am;}
setInterval(updateClock,1000); document.addEventListener('DOMContentLoaded',()=>{updateClock(); requestLocation(false);});
function setHover(){document.querySelectorAll('.mark-btn').forEach(b=>{b.addEventListener('mouseenter',()=>b.classList.add('active'));b.addEventListener('mouseleave',()=>b.classList.remove('active'));});}
document.addEventListener('DOMContentLoaded',setHover);
function requestLocation(notify=true){let info=document.getElementById('locText'), mapLink=document.getElementById('mapLink'), pin=document.getElementById('pin'); if(!info)return; if(!navigator.geolocation){info.textContent='Tu navegador no permite ubicación.'; return;} navigator.geolocation.getCurrentPosition(pos=>{let lat=pos.coords.latitude.toFixed(6), lng=pos.coords.longitude.toFixed(6), acc=Math.round(pos.coords.accuracy); info.textContent='Ubicación real detectada con precisión aproximada de '+acc+' m.'; if(mapLink)mapLink.href='https://www.google.com/maps?q='+lat+','+lng; if(pin){pin.dataset.lat=lat;pin.dataset.lng=lng;pin.dataset.acc=acc;} if(notify)showToast('Ubicación real actualizada');},()=>{info.textContent='Permite ubicación en el navegador para registrar GPS real.'; if(notify)showToast('No se pudo obtener ubicación',false);},{enableHighAccuracy:true,timeout:9000,maximumAge:0});}
async function marcar(tipo){let btn=document.querySelector('[data-tipo="'+tipo+'"]'); if(btn){btn.classList.add('active'); setTimeout(()=>btn.classList.remove('active'),1200);} let p=document.getElementById('pin')||{}; let payload={tipo:tipo,hora_manual:(document.getElementById('selectedTime')?.dataset.value||''),lat:p.dataset?.lat||'',lng:p.dataset?.lng||'',precision:p.dataset?.acc||''}; function send(){fetch('/api/marcar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).then(r=>r.json()).then(j=>{if(j.ok){showToast(j.msg||'Marcación registrada');}else showToast(j.msg||'Error',false)}).catch(()=>showToast('No se pudo registrar',false));} if(navigator.geolocation){navigator.geolocation.getCurrentPosition(pos=>{payload.lat=pos.coords.latitude;payload.lng=pos.coords.longitude;payload.precision=Math.round(pos.coords.accuracy);send();},()=>send(),{enableHighAccuracy:true,timeout:6000});}else send();}
function openWheel(){let m=document.getElementById('wheelModal'); if(!m)return; m.classList.add('show'); setNowWheel(); syncWheel();}
function closeWheel(){document.getElementById('wheelModal')?.classList.remove('show')}
function syncWheel(){document.querySelectorAll('.wheel').forEach(w=>{let mid=w.scrollTop+80; [...w.children].forEach(c=>c.classList.toggle('sel', Math.abs((c.offsetTop+22)-mid)<24));});}
function setNowWheel(){let d=new Date(); setWheel('wh', d.getHours()%12||12); setWheel('wm', d.getMinutes()); setWheel('wa', d.getHours()<12?'AM':'PM'); syncWheel();}
function setWheel(id,val){let w=document.getElementById(id); if(!w)return; [...w.children].forEach(c=>{if(String(c.dataset.v)===String(val)){w.scrollTop=c.offsetTop-58;}})}
function applyWheel(){let h=pick('wh'), m=pick('wm'), a=pick('wa'); let val=String(h).padStart(2,'0')+':'+String(m).padStart(2,'0')+' '+(a==='AM'?'a. m.':'p. m.'); let e=document.getElementById('selectedTime'); if(e){e.dataset.value=val;e.textContent='🕘 Ajustar hora táctil: '+val;} closeWheel(); showToast('Hora manual seleccionada: '+val);}
function pick(id){let w=document.getElementById(id), mid=w.scrollTop+80, best=w.children[0], bd=999; [...w.children].forEach(c=>{let d=Math.abs((c.offsetTop+22)-mid); if(d<bd){bd=d;best=c;}}); return best.dataset.v;}
document.addEventListener('scroll',e=>{if(e.target.classList&&e.target.classList.contains('wheel')) syncWheel()},true);
function initSignature(){let c=document.getElementById('canvas'); if(!c)return; let ctx=c.getContext('2d'), drawing=false; function resize(){let r=c.getBoundingClientRect(); c.width=r.width*devicePixelRatio; c.height=r.height*devicePixelRatio; ctx.setTransform(devicePixelRatio,0,0,devicePixelRatio,0,0); ctx.lineWidth=2; ctx.lineCap='round'; ctx.strokeStyle='#111';} resize(); window.addEventListener('resize',resize); function p(ev){let r=c.getBoundingClientRect(); let t=ev.touches?ev.touches[0]:ev; return {x:t.clientX-r.left,y:t.clientY-r.top};} c.addEventListener('pointerdown',ev=>{drawing=true;let q=p(ev);ctx.beginPath();ctx.moveTo(q.x,q.y)}); c.addEventListener('pointermove',ev=>{if(!drawing)return;let q=p(ev);ctx.lineTo(q.x,q.y);ctx.stroke()}); c.addEventListener('pointerup',()=>drawing=false); c.addEventListener('pointerleave',()=>drawing=false); window.clearSign=()=>ctx.clearRect(0,0,c.width,c.height); window.saveSign=()=>fetch('/api/firma',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({data:c.toDataURL('image/png')})}).then(r=>r.json()).then(j=>showToast(j.msg||'Firma guardada')); window.downloadSign=()=>{let a=document.createElement('a');a.href=c.toDataURL('image/png');a.download='firma_pa.png';a.click()};}
document.addEventListener('DOMContentLoaded',initSignature);
function fakeOk(msg){showToast(msg||'Opción habilitada');}
function testDB(){document.getElementById('dbOk')?.classList.add('show');showToast('Conexión validada para configuración');}
"""

def shell(body, title="P&A Mobile", page_class=""):
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><meta name="theme-color" content="{GREEN}"><title>{title}</title><style>{CSS}</style></head><body><div class="phone {page_class}">{body}</div><script>{JS}</script></body></html>"""



def app_icon(name):
    """Iconos SVG embebidos para evitar dependencias externas en Render."""
    icons = {
        "support": '<svg viewBox="0 0 24 24"><path d="M12 3a8 8 0 0 0-8 8v4a3 3 0 0 0 3 3h2v-7H6a6 6 0 0 1 12 0h-3v7h3a3 3 0 0 0 3-3v-4a8 8 0 0 0-8-8Zm-1 15h2v2h-2z"/></svg>',
        "settings": '<svg viewBox="0 0 24 24"><path d="M19.4 13.5c.1-.5.1-1 .1-1.5s0-1-.1-1.5l2-1.5-2-3.5-2.4 1a7 7 0 0 0-2.6-1.5L14 2h-4l-.4 2.5A7 7 0 0 0 7 6L4.6 5 2.6 8.5l2 1.5c-.1.5-.1 1-.1 1.5s0 1 .1 1.5l-2 1.5 2 3.5 2.4-1a7 7 0 0 0 2.6 1.5L10 22h4l.4-2.5A7 7 0 0 0 17 18l2.4 1 2-3.5-2-1.5ZM12 15.5A3.5 3.5 0 1 1 12 8a3.5 3.5 0 0 1 0 7.5Z"/></svg>',
        "profile": '<svg viewBox="0 0 24 24"><path d="M12 12a4.5 4.5 0 1 0 0-9 4.5 4.5 0 0 0 0 9Zm0 2c-5 0-8 2.5-8 5.5V21h16v-1.5c0-3-3-5.5-8-5.5Z"/></svg>',
        "attendance": '<svg viewBox="0 0 24 24"><path d="M4 5h3v3H4V5Zm0 6h3v3H4v-3Zm0 6h3v3H4v-3ZM10 6h10v2H10V6Zm0 6h10v2H10v-2Zm0 6h10v2H10v-2Z"/></svg>',
        "docs": '<svg viewBox="0 0 24 24"><path d="M6 2h9l5 5v15H6V2Zm8 1.5V8h4.5L14 3.5ZM8 11h8v2H8v-2Zm0 4h8v2H8v-2Z"/></svg>',
        "sign": '<svg viewBox="0 0 24 24"><path d="M4 17.5V21h3.5L18.1 10.4l-3.5-3.5L4 17.5ZM20.7 7.8c.4-.4.4-1 0-1.4l-2.1-2.1a1 1 0 0 0-1.4 0l-1.6 1.6 3.5 3.5 1.6-1.6Z"/></svg>',
        "home": '<svg viewBox="0 0 24 24"><path d="M3 11 12 3l9 8v10h-6v-6H9v6H3V11Z"/></svg>',
        "logout": '<svg viewBox="0 0 24 24"><path d="M10 3h10v18H10v-2h8V5h-8V3ZM8.6 16.6 7.2 18 2 12.8 7.2 7.6 8.6 9l-2.8 2.8H14v2H5.8l2.8 2.8Z"/></svg>',
        "arrow": '<svg viewBox="0 0 24 24"><path d="M15.5 4.5 8 12l7.5 7.5-1.8 1.8L4.4 12l9.3-9.3 1.8 1.8Z"/></svg>',
        "download": '<svg viewBox="0 0 24 24"><path d="M5 20h14v-2H5v2ZM13 4h-2v8H8l4 4 4-4h-3V4Z"/></svg>',
        "location": '<svg viewBox="0 0 24 24"><path d="M12 2a7 7 0 0 0-7 7c0 5.2 7 13 7 13s7-7.8 7-13a7 7 0 0 0-7-7Zm0 9.5A2.5 2.5 0 1 1 12 6a2.5 2.5 0 0 1 0 5.5Z"/></svg>',
        "time": '<svg viewBox="0 0 24 24"><path d="M12 2a10 10 0 1 0 .1 20.1A10 10 0 0 0 12 2Zm1 11h5v-2h-4V6h-2v7Z"/></svg>',
    }
    return '<span class="svgico">' + icons.get(name, icons["docs"]) + '</span>'

def leaves(): return '<div class="leaves"><i></i><i></i><i></i></div>'

def head(title, icon="docs", back="/home"):
    return f'<div class="page-head"><a class="back" href="{back}">{app_icon("arrow")}</a><div class="head-ico">{app_icon(icon)}</div><h1>{title}</h1></div>'

def bottom(active="home"):
    items=[("home","/home","home","Home"),("asistencia","/asistencia","attendance","Asistencia"),("documentos","/documentos","docs","Documentos"),("perfil","/perfil","profile","Perfil")]
    return '<div class="bottom-nav">' + ''.join([f'<a class="{"active" if k==active else ""}" href="{u}"><div>{app_icon(ic)}</div>{t}</a>' for k,u,ic,t in items]) + '</div>'


# ------------------------- Pages -------------------------
@app.route("/")
def index():
    if session.get("usuario"): return redirect(url_for("home"))
    body=f'''<div class="splash green-grad" onclick="location.href='/login'"><div class="logo-circle pa-logo"><div class="logo-inner">P<span>&</span>A</div></div><div class="brand-title">P<span>&</span>A</div><div class="progress"><i></i></div><div class="splash-foot">{COMPANY}<br>{VERSION}</div></div>'''
    return shell(body,"P&A","green-grad")

@app.route("/login", methods=["GET","POST"])
def login():
    err=""
    if request.method=="POST":
        usuario=dni(request.form.get("usuario","")); password=request.form.get("password","")
        c=conn(); r=c.execute("SELECT * FROM users WHERE usuario=?",(usuario,)).fetchone(); c.close()
        if r and check_password_hash(r["password_hash"], password):
            session["usuario"]=usuario; return redirect(url_for("home"))
        err="Usuario o contraseña incorrectos."
    body=f'''<div class="login-wrap login-pro"><div class="login-green"><div class="mini-top"><button type="button" onclick="fakeOk('Soporte habilitado')">{app_icon('support')} Soporte</button><button type="button" onclick="fakeOk('Configuración disponible luego del ingreso')">{app_icon('settings')} Config.</button></div><div class="login-mark"><div>{app_icon('docs')}</div></div><h1>P&A Mobile</h1><p>Iniciar sesión</p></div><form class="login-card login-float" method="post"><h2>Inicia sesión para continuar</h2>{f'<div class="err">{err}</div>' if err else ''}<div class="label">Usuario</div><input class="input" name="usuario" value="11223344" maxlength="8" inputmode="numeric" required><div class="label">Empresa</div><select class="input"><option>{COMPANY}</option></select><div class="label">Contraseña</div><input class="input" name="password" value="123456" type="password" required><button class="login-btn">{app_icon('logout')} Ingresar</button><div class="link">Olvidaste tu contraseña?</div></form><div class="login-mini-tiles"><a class="mini-login-tile" href="#">{app_icon('attendance')}<span>ASISTENCIA</span></a><a class="mini-login-tile" href="#">{app_icon('docs')}<span>DOCUMENTOS</span></a></div><div class="login-leaf"></div><div class="help">¿Problemas para acceder?<br><b>Contactar a Mesa</b></div></div>'''
    return shell(body,"Login P&A")

@app.route("/home")
@login_required
def home():
    u=current_user() or {}
    c=conn(); hoy=date.today().isoformat()
    marcas=c.execute("SELECT tipo,hora FROM marcas WHERE usuario=? AND fecha=? ORDER BY id",(session['usuario'],hoy)).fetchall(); c.close()
    ult = dict(marcas[-1]) if marcas else None
    body=f'''<div class="top-green"><div class="pill-row"><button class="pill" onclick="fakeOk('Soporte habilitado')">{app_icon('support')} Soporte</button><a class="pill" href="/config">{app_icon('settings')} Config.</a></div><div class="avatar avatar-pa">{app_icon('profile')}</div><div class="user-name">{esc(u.get('nombres','NISEDM01'))}</div><div class="welcome-card"><b>Bienvenido(a)</b><p>Sistema de Asistencia</p></div></div><div class="content"><div class="tile-grid"><a class="tile" href="/asistencia"><div class="ico">{app_icon('attendance')}</div><div>ASISTENCIA</div></a><a class="tile" href="/documentos"><div class="ico">{app_icon('docs')}</div><div>DOCUMENTOS</div></a><a class="tile" href="/firma"><div class="ico">{app_icon('sign')}</div><div>FIRMA</div></a><a class="tile" href="/perfil"><div class="ico">{app_icon('profile')}</div><div>PERFIL</div></a></div>{leaves()}<div class="sync-card"><button class="sync-btn" onclick="fakeOk('Datos sincronizados')">↻</button><div><div class="sync-title">Sincronizar Tablas Maestras</div><div class="sync-sub">Actualizado hasta: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div></div><a class="exit" href="/logout">{app_icon('logout')}</a></div></div>'''
    return shell(body,"Home P&A")

@app.route("/asistencia")
@login_required
def asistencia():
    c=conn(); marcas=c.execute("SELECT * FROM marcas WHERE usuario=? AND fecha=? ORDER BY id",(session['usuario'],date.today().isoformat())).fetchall(); c.close()
    hist=''.join([f'<div class="history-row"><span>{esc(m["tipo"].replace("_"," "))}</span><b>{esc(m["hora_manual"] or m["hora"])}</b></div>' for m in marcas]) or '<div class="history-row"><span>Aún no tienes marcaciones hoy</span><b>--:--</b></div>'
    body=f'''<div class="att-page"><div class="hero-att"><a class="back" href="/home">{app_icon('arrow')}</a><div class="cal">{app_icon('attendance')}</div><h1>Registra tu asistencia fácil</h1><p>Registra tu ingreso, salida y refrigerio de forma rápida y segura desde tu celular.</p></div><div class="att-card"><div class="sys">HORA ACTUAL DEL SISTEMA</div><div class="big-clock" id="liveClock">{h12()}</div><div class="date-line">{today_es()}</div><button class="mark-btn btn-main" data-tipo="INGRESO" onclick="marcar('INGRESO')">↪ Registrar Ingreso</button><div class="btn-row"><button class="mark-btn" data-tipo="SALIDA_REFRIGERIO" onclick="marcar('SALIDA_REFRIGERIO')">Salida a<br>Refrigerio</button><button class="mark-btn" data-tipo="RETORNO_REFRIGERIO" onclick="marcar('RETORNO_REFRIGERIO')">Retorno de<br>Refrigerio</button></div><button class="mark-btn" data-tipo="SALIDA" onclick="marcar('SALIDA')">↩ Registrar Salida</button><div id="selectedTime" class="selected-time" onclick="openWheel()" data-value="">{app_icon('time')} Ajustar hora táctil: Ahora</div><div class="map-title"><span>UBICACIÓN ACTUAL</span><button class="signal" onclick="requestLocation(true)">SEÑAL FUERTE</button></div><div class="map"><div class="road-a"></div><div class="road-b"></div><div class="road-c"></div><div class="bubble">Tu ubicación actual</div><div id="pin" class="pin">{app_icon('location')}</div></div><div class="loc-foot"><span id="locText">Detectando ubicación real...</span><a id="mapLink" target="_blank" href="#">Ver mapa</a></div><div class="history-card"><div class="form-title">Historial del día</div>{hist}</div></div>{wheel_modal()}{bottom('asistencia')}</div>'''
    return shell(body,"Asistencia P&A")

def wheel_modal():
    hours=''.join([f'<div data-v="{i}">{i:02d}</div>' for i in range(1,13)])
    mins=''.join([f'<div data-v="{i}">{i:02d}</div>' for i in range(0,60)])
    ampm='<div data-v="AM">AM</div><div data-v="PM">PM</div>'
    return f'''<div class="wheel-modal" id="wheelModal"><div class="wheel-card"><div class="wheel-head"><button onclick="closeWheel()">Cancelar</button><span>Seleccionar hora</span><button onclick="applyWheel()">Aceptar</button></div><div class="wheels"><div class="wheel" id="wh">{hours}</div><div class="wheel" id="wm">{mins}</div><div class="wheel" id="wa">{ampm}</div></div></div></div>'''

@app.route("/documentos")
@login_required
def documentos():
    docs=[("boletas","Boletas Normales"),("utilidades","Constancias de Utilidades"),("vacaciones","Boletas de Vacaciones"),("cts","Constancias de CTS"),("liquidaciones","Constancias de Liquidaciones"),("gratificaciones","Boletas de Gratificaciones"),("constancia-grati","Constancias de Gratificaciones")]
    rows=''.join([f'<a class="doc-row" href="/documentos/{k}"><div class="doc-ico">{app_icon("download")}</div><div><div class="doc-title">{t}</div><div class="doc-sub">Disponible para descarga · PDF/TXT demo</div></div><div class="chev">›</div></a>' for k,t in docs])
    body=head("Gestión de documentos","docs")+f'<div class="doc-list"><div style="margin-bottom:14px"><span class="pdf-badge">{app_icon("docs")} Documentos laborales habilitados</span></div>{rows}</div>{bottom("documentos")} '
    return shell(body,"Documentos P&A")

@app.route("/documentos/<tipo>")
@login_required
def documento_detalle(tipo):
    nombres={"boletas":"Boleta Normal","utilidades":"Constancia de Utilidades","vacaciones":"Boleta de Vacaciones","cts":"Constancia de CTS","liquidaciones":"Constancia de Liquidación","gratificaciones":"Boleta de Gratificación","constancia-grati":"Constancia de Gratificación"}
    titulo=nombres.get(tipo,"Documento")
    u=current_user() or {}
    body=head(titulo,"docs","/documentos")+f'''<div class="form-card"><div class="form-title">Detalle del documento</div><div class="kv"><div><small>Trabajador</small><b>{esc(u.get('nombres'))}</b></div><div><small>DNI</small><b>{esc(u.get('usuario'))}</b></div><div><small>Periodo</small><b>2026</b></div><div><small>Estado</small><b>DISPONIBLE</b></div></div><br><a class="btn-green" style="display:grid;place-items:center" href="/api/documento/{tipo}">Descargar documento</a><br><button class="cfg-btn" onclick="fakeOk('Vista previa habilitada')">Ver vista previa</button></div>{bottom('documentos')}'''
    return shell(body,titulo)

@app.route("/firma")
@login_required
def firma():
    c=conn(); r=c.execute("SELECT actualizado, almacenamiento FROM firmas WHERE usuario=?",(session['usuario'],)).fetchone(); c.close()
    estado = f'<div class="ok-box show">✓ Firma registrada: {esc(r["actualizado"] or "")}</div>' if r else '<div class="mini-help">Aún no tienes una firma registrada. Dibuja o sube tu firma y presiona Guardar.</div>'
    body=head("Gestión de firma","sign")+f'''<div class="form-card">{estado}<div class="mini-help">La firma se almacenará para documentos, constancias y reportes.</div></div><div class="sign-box"><canvas id="canvas"></canvas></div><div class="grid2"><button class="mini-btn" onclick="document.getElementById('fileSign').click()">⇧ Subir Imagen</button><button class="mini-btn" onclick="downloadSign()">⇩ Descargar Firma</button></div><input id="fileSign" type="file" accept="image/*" style="display:none" onchange="fakeOk('Imagen seleccionada')"><div class="grid1"><button class="mini-btn danger" onclick="clearSign()">🗑 Limpiar Lienzo</button></div><div class="grid1"><button class="mini-btn save" onclick="saveSign()">▣ Guardar Cambios</button></div>'''+bottom('home')
    return shell(body,"Firma P&A")

@app.route("/perfil")
@login_required
def perfil():
    u=current_user() or {}
    body=head("Perfil","profile")+f'''<div class="profile-card"><div class="profile-photo">{app_icon('profile')}</div><h2>{esc((u.get('nombres') or '') + ' ' + (u.get('apellidos') or ''))}</h2><p>{esc(u.get('cargo',''))}</p><div class="kv"><div><small>DNI</small><b>{esc(u.get('usuario'))}</b></div><div><small>Empresa</small><b>{esc(u.get('empresa'))}</b></div><div><small>Área</small><b>{esc(u.get('area'))}</b></div><div><small>Régimen</small><b>{esc(u.get('regimen'))}</b></div><div><small>Ingreso</small><b>{esc(u.get('fecha_ingreso'))}</b></div><div><small>Banco</small><b>{esc(u.get('banco'))}</b></div><div><small>Cuenta</small><b>{esc(u.get('cuenta'))}</b></div><div><small>Celular</small><b>{esc(u.get('celular'))}</b></div></div></div>{bottom('perfil')}'''
    return shell(body,"Perfil P&A")

@app.route("/config")
@login_required
def config():
    body=head("CONFIGURACIÓN","settings")+'''<div class="config-card"><a class="cfg-btn primary" style="display:grid;place-items:center" href="/config/trabajador">👥 Cargar trabajador</a><a class="cfg-btn primary" style="display:grid;place-items:center" href="/config/importar-trabajadores">⇧ Importación masiva trabajadores</a><a class="cfg-btn" style="display:grid;place-items:center" href="/config/plantilla-trabajadores">▤ Descargar plantilla trabajadores</a><button class="cfg-btn disabled" onclick="fakeOk('Usuarios habilitado para etapa admin')">👥 Usuarios</button><a class="cfg-btn" style="display:grid;place-items:center" href="/config/firmas">Almacenamiento firmas</a><a class="cfg-btn" style="display:grid;place-items:center" href="/config/db">Conexión Base de Datos</a><a class="cfg-btn" style="display:grid;place-items:center" href="/home">Volver</a></div>'''
    return shell(body,"Configuración P&A")


@app.route("/config/importar-trabajadores", methods=["GET","POST"])
@login_required
def importar_trabajadores():
    mensaje=''
    preview=[]
    if request.method=='POST':
        texto = request.form.get('csv_text','') or ''
        archivo = request.files.get('archivo')
        if archivo and archivo.filename:
            raw = archivo.read()
            try:
                texto = raw.decode('utf-8-sig')
            except Exception:
                texto = raw.decode('latin-1', errors='ignore')
        if not texto.strip():
            mensaje='<div class="mass-bad">No se recibió información. Pega datos CSV o sube un archivo .csv.</div>'
        else:
            try:
                sample = texto[:2048]
                dialect = csv.Sniffer().sniff(sample, delimiters=';,\t,')
            except Exception:
                dialect = csv.excel
                dialect.delimiter = ';' if texto.count(';') >= texto.count(',') else ','
            rows = list(csv.DictReader(texto.splitlines(), dialect=dialect))
            guardados=0; errores=[]
            c=conn(); cur=c.cursor()
            for idx,row in enumerate(rows, start=2):
                norm={str(k or '').strip().upper().replace('Á','A').replace('É','E').replace('Í','I').replace('Ó','O').replace('Ú','U'): (v or '').strip() for k,v in row.items()}
                d=dni(norm.get('DNI') or norm.get('DOCUMENTO') or norm.get('NRO_DOC') or norm.get('NUMERO DE DOCUMENTO'))
                nombres=norm.get('NOMBRES') or norm.get('NOMBRE') or norm.get('TRABAJADOR') or ''
                apellidos=norm.get('APELLIDOS') or ''
                if len(d)!=8 or not nombres:
                    errores.append(f'Fila {idx}: DNI o nombres inválidos')
                    continue
                vals=(d,generate_password_hash('123456'),nombres,apellidos,COMPANY,norm.get('CARGO','TRABAJADOR') or 'TRABAJADOR',norm.get('AREA','OPERACIONES') or 'OPERACIONES',norm.get('REGIMEN','RÉGIMEN GENERAL') or 'RÉGIMEN GENERAL',norm.get('FECHA_INGRESO',''),norm.get('TIPO_DOC','DNI') or 'DNI',norm.get('NRO_DOC',d) or d,norm.get('BANCO',''),norm.get('CUENTA',''),norm.get('CORREO',''),norm.get('CELULAR',''),norm.get('DIRECCION',''), '')
                cur.execute('''INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,avatar)
                               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                               ON CONFLICT(usuario) DO UPDATE SET nombres=excluded.nombres,apellidos=excluded.apellidos,cargo=excluded.cargo,area=excluded.area,regimen=excluded.regimen,fecha_ingreso=excluded.fecha_ingreso,tipo_doc=excluded.tipo_doc,nro_doc=excluded.nro_doc,banco=excluded.banco,cuenta=excluded.cuenta,correo=excluded.correo,celular=excluded.celular,direccion=excluded.direccion''', vals)
                guardados+=1
                if len(preview)<10: preview.append({'dni':d,'nombres':nombres,'cargo':vals[5],'area':vals[6]})
            c.commit(); c.close()
            extra = ('<br>'.join(errores[:6]) if errores else '')
            mensaje=f'<div class="mass-ok">✓ Importación finalizada: {guardados} trabajador(es) guardado(s). Clave inicial: 123456.</div>' + (f'<div class="mass-bad">{extra}</div>' if extra else '')
    preview_html=''
    if preview:
        preview_html='<div class="mass-preview"><table><tr><th>DNI</th><th>Nombres</th><th>Cargo</th><th>Área</th></tr>' + ''.join([f'<tr><td>{esc(r["dni"])}</td><td>{esc(r["nombres"])}</td><td>{esc(r["cargo"])}</td><td>{esc(r["area"])}</td></tr>' for r in preview]) + '</table></div>'
    ejemplo='DNI;NOMBRES;APELLIDOS;CARGO;AREA;REGIMEN;FECHA_INGRESO;TIPO_DOC;NRO_DOC;BANCO;CUENTA;CORREO;CELULAR;DIRECCION\n11223344;NISEDM01;;TRABAJADOR;OPERACIONES;RÉGIMEN GENERAL;2026-06-22;DNI;11223344;BCP;001-000000000;trabajador@pa.com;999999999;TRUJILLO'
    body=head("IMPORTAR TRABAJADORES","profile","/config")+f'''<form class="form-card" method="post" enctype="multipart/form-data">{mensaje}<div class="form-title">Carga masiva de trabajadores</div><div class="mini-help">Sube un archivo CSV o pega la información usando la plantilla. Se actualizarán DNI existentes y se crearán nuevos usuarios con clave inicial <b>123456</b>.</div><div class="field"><label>Archivo CSV</label><input type="file" name="archivo" accept=".csv,text/csv"></div><div class="field"><label>Pegar CSV</label><textarea name="csv_text" style="width:100%;height:150px;border:1px solid #dde3e9;border-radius:9px;padding:9px;font-size:11px;font-weight:700">{esc(ejemplo)}</textarea></div><button class="btn-green">⇧ Importar trabajadores</button><br><br><a class="cfg-btn" style="display:grid;place-items:center" href="/config/plantilla-trabajadores">Descargar plantilla</a>{preview_html}</form>{bottom('perfil')}'''
    return shell(body,"Importar trabajadores")

@app.route("/config/trabajador", methods=["GET","POST"])
@login_required
def config_trabajador():
    msg=''
    if request.method=='POST':
        data={k:request.form.get(k,'').strip() for k in ['dni','nombres','apellidos','fecha_ingreso','celular','correo','direccion','regimen','tipo_doc','nro_doc','banco','cuenta','cargo','area']}
        d=dni(data['dni'])
        if len(d)!=8: msg='<div class="err">DNI inválido.</div>'
        else:
            c=conn(); cur=c.cursor(); cur.execute('''INSERT INTO users(usuario,password_hash,nombres,apellidos,empresa,cargo,area,regimen,fecha_ingreso,tipo_doc,nro_doc,banco,cuenta,correo,celular,direccion,avatar) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(usuario) DO UPDATE SET nombres=excluded.nombres,apellidos=excluded.apellidos,cargo=excluded.cargo,area=excluded.area,regimen=excluded.regimen,fecha_ingreso=excluded.fecha_ingreso,tipo_doc=excluded.tipo_doc,nro_doc=excluded.nro_doc,banco=excluded.banco,cuenta=excluded.cuenta,correo=excluded.correo,celular=excluded.celular,direccion=excluded.direccion''',(d,generate_password_hash('123456'),data['nombres'],data['apellidos'],COMPANY,data['cargo'] or 'TRABAJADOR',data['area'] or 'OPERACIONES',data['regimen'],data['fecha_ingreso'],data['tipo_doc'],data['nro_doc'],data['banco'],data['cuenta'],data['correo'],data['celular'],data['direccion'],'')); c.commit(); c.close(); msg='<div class="ok-box show">✓ Trabajador guardado. Clave inicial: 123456</div>'
    body=head("CARGAR TRABAJADOR","profile","/config")+f'''<form class="form-card" method="post">{msg}<div class="form-title">Información personal</div><div class="field"><label>DNI *</label><input name="dni" maxlength="8" required></div><div class="field"><label>Nombres *</label><input name="nombres" required></div><div class="field"><label>Apellidos</label><input name="apellidos"></div><div class="two"><div class="field"><label>Cargo</label><input name="cargo"></div><div class="field"><label>Área</label><input name="area"></div></div><div class="field"><label>Celular</label><input name="celular"></div><div class="field"><label>Correo electrónico</label><input name="correo" type="email"></div><div class="field"><label>Dirección</label><input name="direccion"></div><div class="form-title">Información para boletas y documentos</div><div class="two"><div class="field"><label>Régimen laboral *</label><select name="regimen"><option>RÉGIMEN GENERAL</option><option>RÉGIMEN AGRARIO</option><option>CONSTRUCCIÓN CIVIL</option></select></div><div class="field"><label>Fecha de ingreso</label><input name="fecha_ingreso" type="date"></div></div><div class="two"><div class="field"><label>Tipo de documento</label><select name="tipo_doc"><option>DNI</option><option>CE</option><option>PASAPORTE</option></select></div><div class="field"><label>Número de documento</label><input name="nro_doc"></div></div><div class="two"><div class="field"><label>Banco</label><input name="banco"></div><div class="field"><label>Cuenta bancaria</label><input name="cuenta"></div></div><button class="btn-green">▣ Guardar trabajador</button></form>'''
    return shell(body,"Cargar trabajador")

@app.route('/config/plantilla-trabajadores')
@login_required
def plantilla_trabajadores():
    csv='DNI,NOMBRES,APELLIDOS,CARGO,AREA,REGIMEN,FECHA_INGRESO,TIPO_DOC,NRO_DOC,BANCO,CUENTA,CORREO,CELULAR,DIRECCION\n11223344,NISEDM01,,TRABAJADOR,OPERACIONES,REGIMEN GENERAL,2026-06-22,DNI,11223344,BCP,001-000000000,trabajador@pa.com,999999999,TRUJILLO\n'
    return Response(csv, mimetype='text/csv', headers={'Content-Disposition':'attachment; filename=plantilla_trabajadores_pa.csv'})

@app.route('/config/firmas', methods=['GET','POST'])
@login_required
def config_firmas():
    if request.method=='POST':
        modo=request.form.get('almacenamiento','LOCAL')
        c=conn(); c.execute("INSERT INTO firmas(usuario,almacenamiento,actualizado) VALUES(?,?,?) ON CONFLICT(usuario) DO UPDATE SET almacenamiento=excluded.almacenamiento, actualizado=excluded.actualizado",(session['usuario'],modo,datetime.now().isoformat())); c.commit(); c.close()
    body=head("ALMACENAMIENTO DE FIRMAS","sign","/config")+'''<form class="form-card" method="post"><div class="mini-help">Las firmas se almacenan de forma segura para su uso en documentos y reportes.</div><div class="field"><label><input type="radio" name="almacenamiento" value="LOCAL" checked> Almacenamiento local recomendado para Render</label><div class="mini-help">Las firmas se guardarán en la base SQLite del servidor.</div></div><div class="field"><label><input type="radio" name="almacenamiento" value="SQLSERVER"> SQL Server requiere configuración</label><div class="mini-help">Las firmas se guardarán en la base SQL Server cuando se conecte.</div></div><button class="btn-green">▣ Guardar configuración</button></form>'''
    return shell(body,"Firmas")

@app.route('/config/db', methods=['GET','POST'])
@login_required
def config_db():
    saved=''
    if request.method=='POST':
        vals=(request.form.get('servidor',''),request.form.get('base_datos',''),request.form.get('usuario',''),request.form.get('password',''),request.form.get('puerto','1433'),datetime.now().isoformat())
        c=conn(); c.execute("INSERT INTO config_db(motor,servidor,base_datos,usuario,password,puerto,actualizado) VALUES('SQLSERVER',?,?,?,?,?,?)", vals); c.commit(); c.close(); saved='<div class="ok-box show">✓ Configuración guardada para conexión SQL Server futura.</div>'
    body=head("CONEXIÓN BASE DE DATOS","docs","/config")+f'''<form class="form-card" method="post">{saved}<div class="two"><button type="button" class="cfg-btn">SQLite Local</button><button type="button" class="cfg-btn primary">SQL Server</button></div><div class="field"><label>Servidor *</label><input name="servidor" placeholder="Ej. DESKTOP\\SQLEXPRESS"></div><div class="field"><label>Base de datos *</label><input name="base_datos" placeholder="Ej. ASISTENCIA_DB"></div><div class="field"><label>Usuario *</label><input name="usuario" placeholder="Ej. sa"></div><div class="field"><label>Contraseña *</label><input name="password" type="password" placeholder="Ingrese contraseña"></div><div class="field"><label>Puerto</label><input name="puerto" value="1433"></div><button type="button" class="btn-green" onclick="testDB()">🔌 Probar conexión</button><br><br><div id="dbOk" class="ok-box">✓ Conexión lista para validar cuando se habilite el driver SQL Server.</div><br><button class="btn-green">▣ Guardar configuración</button></form>'''
    return shell(body,"Base de datos")

@app.route('/api/marcar', methods=['POST'])
@login_required
def api_marcar():
    data=request.get_json(silent=True) or {}; tipo=data.get('tipo','').strip().upper()
    if tipo not in ['INGRESO','SALIDA_REFRIGERIO','RETORNO_REFRIGERIO','SALIDA']:
        return jsonify(ok=False,msg='Tipo de marcación inválido')
    now=datetime.now(); c=conn(); c.execute('''INSERT INTO marcas(usuario,tipo,fecha,hora,fecha_hora,lat,lng,precision,direccion,hora_manual) VALUES(?,?,?,?,?,?,?,?,?,?)''',(session['usuario'],tipo,date.today().isoformat(),h12(now),now.isoformat(),str(data.get('lat','')),str(data.get('lng','')),str(data.get('precision','')),'GPS Navegador',str(data.get('hora_manual','')))); c.commit(); c.close()
    return jsonify(ok=True,msg=f'{tipo.replace("_"," ").title()} registrado correctamente')

@app.route('/api/firma', methods=['POST'])
@login_required
def api_firma():
    data=request.get_json(silent=True) or {}; img=data.get('data','')
    c=conn(); c.execute("INSERT INTO firmas(usuario,data_url,actualizado) VALUES(?,?,?) ON CONFLICT(usuario) DO UPDATE SET data_url=excluded.data_url, actualizado=excluded.actualizado",(session['usuario'],img,datetime.now().isoformat())); c.commit(); c.close()
    return jsonify(ok=True,msg='Firma guardada correctamente')

@app.route('/api/documento/<tipo>')
@login_required
def api_documento(tipo):
    u=current_user() or {}; title=f"Documento P&A - {tipo}"
    content=f"{title}\nTrabajador: {u.get('nombres','')}\nDNI: {u.get('usuario','')}\nEmpresa: {u.get('empresa','')}\nFecha: {today_es()}\nEstado: Disponible\n"
    return Response(content, mimetype='text/plain; charset=utf-8', headers={'Content-Disposition':f'attachment; filename={tipo}_pa.txt'})

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT','5000')), debug=False)
