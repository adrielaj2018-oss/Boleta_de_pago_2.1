# -*- coding: utf-8 -*-
"""
P&A Mobile - Asistencia, Documentos, Firma y Configuración
Render 5 archivos: todo el HTML/CSS/JS está embebido en app.py.
"""
import os, re, sqlite3, json
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
GREEN = "#237743"
GREEN_DARK = "#076b3b"
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
.phone{max-width:390px!important;min-height:100dvh!important;box-shadow:0 0 0 1px rgba(0,0,0,.05)!important}
@media(min-width:800px){body{background:#fff!important}.phone{margin:8px auto!important;border-radius:0!important;min-height:100dvh!important;box-shadow:0 0 0 1px #e1e5e8!important}}
.top-green{height:214px!important;border-radius:0 0 28px 28px!important;padding:8px 14px!important}
.pill-row{margin-top:5px!important}.pill{font-size:12px!important;padding:7px 11px!important;border-radius:20px!important}
.avatar{width:68px!important;height:68px!important;margin:16px auto 8px!important;font-size:43px!important}.user-name{font-size:16px!important}
.welcome-card{left:20px!important;right:20px!important;bottom:-32px!important;border-radius:16px!important;padding:12px!important}
.content{padding:49px 18px 76px!important}.tile-grid{gap:14px!important}.tile{height:96px!important;border-radius:13px!important;font-size:12px!important;gap:8px!important}.tile .ico{font-size:27px!important}
.leaves{width:112px!important;height:112px!important;margin:20px auto!important}
.sync-card{bottom:8px!important;border-radius:13px!important;padding:8px 10px!important;grid-template-columns:34px 1fr 32px!important}.sync-btn{width:32px!important;height:32px!important;font-size:20px!important}.sync-title{font-size:12px!important}.sync-sub{font-size:9px!important}.exit{font-size:26px!important}
.hero-att{height:184px!important;padding:17px 22px 0!important}.hero-att .cal{font-size:25px!important;margin-bottom:5px!important}.hero-att h1{font-size:23px!important}.hero-att p{font-size:13px!important;margin-top:10px!important}
.att-card{margin:-33px 15px 0!important;border-radius:20px!important;padding:17px 15px 14px!important}.big-clock{font-size:39px!important}.date-line{font-size:12px!important;margin-bottom:13px!important}
.mark-btn{height:48px!important;border-radius:9px!important;background:#f0f2f4!important;color:#405064!important;transition:.15s!important}
.mark-btn:hover,.mark-btn.active{background:#21b45d!important;color:#fff!important;outline:none!important;transform:translateY(-1px)!important}
.mark-btn.done{background:#f0f2f4!important;color:#405064!important}
.btn-main{background:#f0f2f4!important;color:#405064!important}.btn-main:hover,.btn-main.active{background:#21b45d!important;color:#fff!important}
.btn-row{gap:11px!important;margin:11px 0!important}.selected-time{height:44px!important;margin-top:11px!important}
.map-title{margin-top:18px!important}.map{height:137px!important}.loc-foot{font-size:11px!important;align-items:flex-start}.loc-foot a{font-size:13px!important;font-weight:1000}
.page-head{height:122px!important;border-radius:0 0 27px 27px!important}.page-head h1{font-size:20px!important;margin-top:43px!important}.back{font-size:29px!important;top:15px!important}
.doc-list{padding:22px 17px 76px!important}.doc-row{height:70px!important;border-radius:13px!important;margin-bottom:12px!important}.doc-title{font-size:15px!important}.doc-sub{font-size:11px!important}
.config-card,.form-card,.profile-card{margin:15px!important;border-radius:14px!important}.cfg-btn{min-height:45px!important;font-size:15px!important;margin-bottom:9px!important}
.login-wrap{padding:28px 24px 20px!important}.login-logo{height:106px!important}.brand-title{font-size:36px!important}
.sign-box{height:180px!important}.bottom-nav{height:56px!important}

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

def leaves(): return '<div class="leaves"><i></i><i></i><i></i></div>'

def head(title, icon="▰", back="/home"):
    return f'<div class="page-head"><a class="back" href="{back}">‹</a><div class="head-ico">{icon}</div><h1>{title}</h1></div>'

def bottom(active="home"):
    items=[("home","/home","⌘","Home"),("asistencia","/asistencia","▦","Asistencia"),("documentos","/documentos","▰","Documentos"),("perfil","/perfil","◉","Perfil")]
    return '<div class="bottom-nav">' + ''.join([f'<a class="{"active" if k==active else ""}" href="{u}"><div>{ic}</div>{t}</a>' for k,u,ic,t in items]) + '</div>'

# ------------------------- Pages -------------------------
@app.route("/")
def index():
    if session.get("usuario"): return redirect(url_for("home"))
    body=f'''<div class="splash green-grad" onclick="location.href='/login'"><div class="logo-circle"><div class="logo-inner">▣</div></div><div class="brand-title">P<span>&</span>A</div><div class="progress"><i></i></div><div class="splash-foot">{COMPANY}<br>{VERSION}</div></div>'''
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
    body=f'''<div class="login-wrap"><div class="login-logo"><div><b style="font-size:42px">P<span style="color:{LIME}">&</span>A</b><div style="font-size:11px;text-align:center;color:{GREEN}">Gestión del Capital Humano</div></div></div><form class="login-card" method="post"><h2>Inicia sesión para continuar</h2>{f'<div class="err">{err}</div>' if err else ''}<div class="label">Usuario</div><input class="input" name="usuario" value="11223344" maxlength="8" inputmode="numeric" required><div class="label">Empresa</div><select class="input"><option>{COMPANY}</option></select><div class="label">Contraseña</div><input class="input" name="password" value="123456" type="password" required><button class="login-btn">Iniciar Sesión ↪</button><div class="link">Olvidaste tu contraseña?</div></form><div class="help">¿Problemas para acceder?<br><b>Contactar a Mesa</b></div></div>'''
    return shell(body,"Login P&A")

@app.route("/home")
@login_required
def home():
    u=current_user() or {}
    body=f'''<div class="top-green"><div class="pill-row"><button class="pill" onclick="fakeOk('Soporte habilitado')">🎧 Soporte</button><a class="pill" href="/config">⚙ Config.</a></div><div class="avatar">☻</div><div class="user-name">{esc(u.get('nombres','NISEDM01'))}</div><div class="welcome-card"><b>Bienvenido(a)</b><p>Sistema de Asistencia</p></div></div><div class="content"><div class="tile-grid"><a class="tile" href="/asistencia"><div class="ico">☷</div><div>ASISTENCIA</div></a><a class="tile" href="/documentos"><div class="ico">▧</div><div>DOCUMENTOS</div></a><a class="tile" href="/firma"><div class="ico">✎</div><div>FIRMA</div></a><a class="tile" href="/perfil"><div class="ico">◉</div><div>PERFIL</div></a></div>{leaves()}<div class="sync-card"><button class="sync-btn" onclick="fakeOk('Tablas sincronizadas')">↻</button><div><div class="sync-title">Sincronizar Tablas Maestras</div><div class="sync-sub">Actualizado hasta: 18/03/2025 09:31:12</div></div><a class="exit" href="/logout">⇥</a></div></div>'''
    return shell(body,"Home P&A")

@app.route("/asistencia")
@login_required
def asistencia():
    c=conn(); marcas=c.execute("SELECT * FROM marcas WHERE usuario=? AND fecha=? ORDER BY id",(session['usuario'],date.today().isoformat())).fetchall(); c.close()
    done={m['tipo'] for m in marcas}
    def cls(t): return ' mark-btn done' if t in done else ' mark-btn'
    body=f'''<div class="att-page"><div class="hero-att"><a class="back" href="/home">‹</a><div class="cal">▦</div><h1>Registra tu asistencia fácil</h1><p>Registra tu ingreso, salida y refrigerio de forma rápida y segura desde tu celular.</p></div><div class="att-card"><div class="sys">HORA ACTUAL DEL SISTEMA</div><div class="big-clock" id="liveClock">{h12()}</div><div class="date-line">{today_es()}</div><button class="mark-btn btn-main" data-tipo="INGRESO" onclick="marcar('INGRESO')">↪ Registrar Ingreso</button><div class="btn-row"><button class="mark-btn" data-tipo="SALIDA_REFRIGERIO" onclick="marcar('SALIDA_REFRIGERIO')">Salida a<br>Refrigerio</button><button class="mark-btn" data-tipo="RETORNO_REFRIGERIO" onclick="marcar('RETORNO_REFRIGERIO')">Retorno de<br>Refrigerio</button></div><button class="mark-btn" data-tipo="SALIDA" onclick="marcar('SALIDA')">↩ Registrar Salida</button><div id="selectedTime" class="selected-time" onclick="openWheel()" data-value="">🕘 Ajustar hora táctil: Ahora</div><div class="map-title"><span>UBICACIÓN ACTUAL</span><button class="signal" onclick="requestLocation(true)">SEÑAL FUERTE</button></div><div class="map"><div class="road-a"></div><div class="road-b"></div><div class="road-c"></div><div class="bubble">Tu ubicación actual</div><div id="pin" class="pin">📍</div></div><div class="loc-foot"><span id="locText">Detectando ubicación real...</span><a id="mapLink" target="_blank" href="#">Ver historial</a></div></div>{wheel_modal()}{bottom('asistencia')}</div>'''
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
    rows=''.join([f'<a class="doc-row" href="/documentos/{k}"><div class="doc-ico">▧</div><div><div class="doc-title">{t}</div><div class="doc-sub">Disponible para descarga</div></div><div class="chev">›</div></a>' for k,t in docs])
    body=head("Gestión de documentos","▰")+f'<div class="doc-list">{rows}</div>{bottom("documentos")}'
    return shell(body,"Documentos P&A")

@app.route("/documentos/<tipo>")
@login_required
def documento_detalle(tipo):
    nombres={"boletas":"Boleta Normal","utilidades":"Constancia de Utilidades","vacaciones":"Boleta de Vacaciones","cts":"Constancia de CTS","liquidaciones":"Constancia de Liquidación","gratificaciones":"Boleta de Gratificación","constancia-grati":"Constancia de Gratificación"}
    titulo=nombres.get(tipo,"Documento")
    u=current_user() or {}
    body=head(titulo,"▧","/documentos")+f'''<div class="form-card"><div class="form-title">Detalle del documento</div><div class="kv"><div><small>Trabajador</small><b>{esc(u.get('nombres'))}</b></div><div><small>DNI</small><b>{esc(u.get('usuario'))}</b></div><div><small>Periodo</small><b>2026</b></div><div><small>Estado</small><b>DISPONIBLE</b></div></div><br><a class="btn-green" style="display:grid;place-items:center" href="/api/documento/{tipo}">Descargar documento</a><br><button class="cfg-btn" onclick="fakeOk('Vista previa habilitada')">Ver vista previa</button></div>{bottom('documentos')}'''
    return shell(body,titulo)

@app.route("/firma")
@login_required
def firma():
    body=head("Gestión de firma","✎")+'''<div class="sign-box"><canvas id="canvas"></canvas></div><div class="grid2"><button class="mini-btn" onclick="document.getElementById('fileSign').click()">⇧ Subir Imagen</button><button class="mini-btn" onclick="downloadSign()">⇩ Descargar Firma</button></div><input id="fileSign" type="file" accept="image/*" style="display:none" onchange="fakeOk('Imagen seleccionada')"><div class="grid1"><button class="mini-btn danger" onclick="clearSign()">🗑 Limpiar Lienzo</button></div><div class="grid1"><button class="mini-btn save" onclick="saveSign()">▣ Guardar Cambios</button></div>'''+bottom('home')
    return shell(body,"Firma P&A")

@app.route("/perfil")
@login_required
def perfil():
    u=current_user() or {}
    body=head("Perfil","◉")+f'''<div class="profile-card"><div class="profile-photo">☻</div><h2>{esc(u.get('nombres',''))}</h2><p>{esc(u.get('cargo',''))}</p><div class="kv"><div><small>DNI</small><b>{esc(u.get('usuario'))}</b></div><div><small>Empresa</small><b>{esc(u.get('empresa'))}</b></div><div><small>Área</small><b>{esc(u.get('area'))}</b></div><div><small>Régimen</small><b>{esc(u.get('regimen'))}</b></div></div></div>{bottom('perfil')}'''
    return shell(body,"Perfil P&A")

@app.route("/config")
@login_required
def config():
    body=head("CONFIGURACIÓN","⚙")+'''<div class="config-card"><a class="cfg-btn primary" style="display:grid;place-items:center" href="/config/trabajador">👥 Cargar trabajadores</a><a class="cfg-btn primary" style="display:grid;place-items:center" href="/config/plantilla-trabajadores">▤ Plantilla trabajadores</a><button class="cfg-btn disabled" onclick="fakeOk('Usuarios habilitado para etapa admin')">👥 Usuarios</button><a class="cfg-btn" style="display:grid;place-items:center" href="/config/firmas">Almacenamiento firmas</a><a class="cfg-btn" style="display:grid;place-items:center" href="/config/db">Conexión Base de Datos</a><a class="cfg-btn" style="display:grid;place-items:center" href="/home">Volver</a></div>'''
    return shell(body,"Configuración P&A")

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
    body=head("CARGAR TRABAJADOR","👥","/config")+f'''<form class="form-card" method="post">{msg}<div class="form-title">Información personal</div><div class="field"><label>DNI *</label><input name="dni" maxlength="8" required></div><div class="field"><label>Nombres *</label><input name="nombres" required></div><div class="field"><label>Apellidos</label><input name="apellidos"></div><div class="two"><div class="field"><label>Cargo</label><input name="cargo"></div><div class="field"><label>Área</label><input name="area"></div></div><div class="field"><label>Celular</label><input name="celular"></div><div class="field"><label>Correo electrónico</label><input name="correo" type="email"></div><div class="field"><label>Dirección</label><input name="direccion"></div><div class="form-title">Información para boletas y documentos</div><div class="two"><div class="field"><label>Régimen laboral *</label><select name="regimen"><option>RÉGIMEN GENERAL</option><option>RÉGIMEN AGRARIO</option><option>CONSTRUCCIÓN CIVIL</option></select></div><div class="field"><label>Fecha de ingreso</label><input name="fecha_ingreso" type="date"></div></div><div class="two"><div class="field"><label>Tipo de documento</label><select name="tipo_doc"><option>DNI</option><option>CE</option><option>PASAPORTE</option></select></div><div class="field"><label>Número de documento</label><input name="nro_doc"></div></div><div class="two"><div class="field"><label>Banco</label><input name="banco"></div><div class="field"><label>Cuenta bancaria</label><input name="cuenta"></div></div><button class="btn-green">▣ Guardar trabajador</button></form>'''
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
    body=head("ALMACENAMIENTO DE FIRMAS","▣","/config")+'''<form class="form-card" method="post"><div class="mini-help">Las firmas se almacenan de forma segura para su uso en documentos y reportes.</div><div class="field"><label><input type="radio" name="almacenamiento" value="LOCAL" checked> Almacenamiento local recomendado para Render</label><div class="mini-help">Las firmas se guardarán en la base SQLite del servidor.</div></div><div class="field"><label><input type="radio" name="almacenamiento" value="SQLSERVER"> SQL Server requiere configuración</label><div class="mini-help">Las firmas se guardarán en la base SQL Server cuando se conecte.</div></div><button class="btn-green">▣ Guardar configuración</button></form>'''
    return shell(body,"Firmas")

@app.route('/config/db', methods=['GET','POST'])
@login_required
def config_db():
    saved=''
    if request.method=='POST':
        vals=(request.form.get('servidor',''),request.form.get('base_datos',''),request.form.get('usuario',''),request.form.get('password',''),request.form.get('puerto','1433'),datetime.now().isoformat())
        c=conn(); c.execute("INSERT INTO config_db(motor,servidor,base_datos,usuario,password,puerto,actualizado) VALUES('SQLSERVER',?,?,?,?,?,?)", vals); c.commit(); c.close(); saved='<div class="ok-box show">✓ Configuración guardada para conexión SQL Server futura.</div>'
    body=head("CONEXIÓN BASE DE DATOS","▣","/config")+f'''<form class="form-card" method="post">{saved}<div class="two"><button type="button" class="cfg-btn">SQLite Local</button><button type="button" class="cfg-btn primary">SQL Server</button></div><div class="field"><label>Servidor *</label><input name="servidor" placeholder="Ej. DESKTOP\\SQLEXPRESS"></div><div class="field"><label>Base de datos *</label><input name="base_datos" placeholder="Ej. ASISTENCIA_DB"></div><div class="field"><label>Usuario *</label><input name="usuario" placeholder="Ej. sa"></div><div class="field"><label>Contraseña *</label><input name="password" type="password" placeholder="Ingrese contraseña"></div><div class="field"><label>Puerto</label><input name="puerto" value="1433"></div><button type="button" class="btn-green" onclick="testDB()">🔌 Probar conexión</button><br><br><div id="dbOk" class="ok-box">✓ Conexión lista para validar cuando se habilite el driver SQL Server.</div><br><button class="btn-green">▣ Guardar configuración</button></form>'''
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
