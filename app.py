# -*- coding: utf-8 -*-
"""
P&A Mobile - Asistencia, Documentos y Firma Digital
Versión Render 5 archivos: app.py + Procfile + requirements.txt + runtime.txt + README.md
Todo el HTML/CSS/JS está embebido para evitar errores TemplateNotFound.
"""
import os
import re
import sqlite3
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
        nombres TEXT, empresa TEXT, cargo TEXT, area TEXT, avatar TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS marcas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT, tipo TEXT, fecha TEXT, hora TEXT, fecha_hora TEXT,
        lat TEXT, lng TEXT, precision TEXT, direccion TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS firmas(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE, data_url TEXT, actualizado TEXT
    )""")
    cur.execute("SELECT id FROM users WHERE usuario=?", ("11223344",))
    if not cur.fetchone():
        cur.execute("""INSERT INTO users(usuario,password_hash,nombres,empresa,cargo,area,avatar)
                     VALUES(?,?,?,?,?,?,?)""",
                    ("11223344", generate_password_hash("123456"), "NISEDM01", COMPANY,
                     "Trabajador", "Operaciones", ""))
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
    dt = dt or datetime.now()
    h = dt.hour
    ampm = "a. m." if h < 12 else "p. m."
    hh = h % 12 or 12
    return f"{hh:02d}:{dt.minute:02d} {ampm}"

def esc(s):
    return str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# ------------------------- CSS/JS Shell -------------------------
CSS = f"""
:root{{--green:{GREEN};--green-dark:{GREEN_DARK};--lime:{LIME};--ink:#101528;--muted:#8a94a6;--soft:#f5f7f8;--line:#e6eaee;--shadow:0 14px 35px rgba(0,0,0,.16)}}
*{{box-sizing:border-box;-webkit-tap-highlight-color:transparent}}
html,body{{margin:0;width:100%;min-height:100%;font-family:Inter,Segoe UI,Arial,sans-serif;background:#eef2f1;color:var(--ink)}}
body{{display:flex;justify-content:center}}
a{{text-decoration:none;color:inherit}}
.phone{{width:100%;max-width:430px;min-height:100dvh;background:#fff;position:relative;overflow:hidden;box-shadow:0 0 0 1px rgba(0,0,0,.04)}}
.green-bg{{background:radial-gradient(circle at 55% 28%, #2d884c 0%, var(--green) 38%, var(--green-dark) 100%);color:white}}
.status{{height:26px;display:flex;align-items:center;justify-content:space-between;padding:7px 14px 0;font-weight:900;font-size:13px;letter-spacing:.2px}}
.status .icons{{display:flex;gap:5px;align-items:center}}.bar{{width:8px;height:12px;background:#0c1426;border-radius:1px}}.wifi{{width:15px;height:10px;border-top:3px solid #0c1426;border-radius:50%}}
.splash{{height:100dvh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:26px;position:relative}}
.logo-circle{{width:178px;height:178px;border-radius:50%;background:#fff;border:9px solid var(--lime);display:grid;place-items:center;box-shadow:0 15px 35px rgba(0,0,0,.22);position:relative;overflow:hidden}}
.logo-circle:before{{content:"";position:absolute;inset:110px -10px -8px;background:linear-gradient(135deg,#69a62d,#c7db22 45%,#2a8f3d 46%,#1f7335);transform:skewY(-18deg)}}
.logo-inner{{position:relative;z-index:2;color:#1f783c;font-size:72px;font-weight:1000;line-height:1}}
.brand-title{{font-size:48px;font-weight:1000;margin-top:25px;letter-spacing:.5px;text-shadow:0 3px 12px rgba(0,0,0,.18)}}.brand-title span{{color:var(--lime)}}
.progress{{position:absolute;bottom:112px;width:160px;height:8px;background:#07592f;border-radius:999px;overflow:hidden}}.progress i{{display:block;width:56%;height:100%;background:var(--lime);border-radius:999px}}
.splash-foot{{position:absolute;bottom:35px;text-align:center;font-size:15px;line-height:1.45;color:#fff}}
.top-green{{height:296px;border-radius:0 0 32px 32px;padding:10px 18px;background:radial-gradient(circle at 60% 15%,#2f8c4f 0%,var(--green) 36%,var(--green-dark) 100%);color:white;position:relative}}
.pill-row{{display:flex;justify-content:space-between;align-items:center;margin-top:8px}}.pill{{background:white;color:var(--green);border-radius:26px;padding:9px 13px;font-weight:900;font-size:14px;box-shadow:0 8px 18px rgba(0,0,0,.15)}}
.avatar{{width:86px;height:86px;border-radius:50%;background:#fff;color:var(--green);display:grid;place-items:center;margin:18px auto 9px;font-size:62px;font-weight:900}}.user-name{{text-align:center;font-size:19px;font-weight:1000}}
.welcome-card{{background:white;color:var(--ink);border-radius:20px;padding:18px 14px;text-align:center;box-shadow:var(--shadow);position:absolute;left:24px;right:24px;bottom:-44px}}.welcome-card b{{font-size:18px}}.welcome-card p{{margin:7px 0 0;color:#535b6b;font-size:16px}}
.content{{padding:66px 20px 92px;min-height:calc(100dvh - 296px);position:relative;background:#fff}}
.tile-grid{{display:grid;grid-template-columns:1fr 1fr;gap:22px;margin-top:14px}}.tile{{height:146px;border-radius:18px;background:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;box-shadow:0 13px 30px rgba(0,0,0,.14);font-weight:1000;font-size:16px;border:1px solid #f0f1f2}}.tile .ico{{font-size:44px;color:var(--green)}}
.leaves{{width:170px;height:170px;margin:38px auto 36px;position:relative;opacity:.55}}.leaves i{{position:absolute;width:86px;height:145px;border-radius:70% 25% 70% 25%;filter:drop-shadow(0 4px 3px rgba(0,0,0,.12))}}.leaves i:nth-child(1){{background:#ffc9b3;left:20px;transform:rotate(26deg)}}.leaves i:nth-child(2){{background:#bdd2a4;left:70px;transform:rotate(26deg)}}.leaves i:nth-child(3){{background:#fff0a4;left:98px;top:42px;transform:rotate(26deg)}}
.sync-card{{position:absolute;left:14px;right:14px;bottom:13px;background:#fff;border-radius:18px;padding:14px 17px;box-shadow:var(--shadow);display:grid;grid-template-columns:45px 1fr 45px;gap:10px;align-items:center}}.sync-btn{{width:43px;height:43px;border-radius:50%;background:var(--green);display:grid;place-items:center;color:white;font-size:28px}}.exit{{font-size:38px;color:#ff3030;text-align:center}}.sync-title{{font-weight:1000}}.sync-sub{{color:#008a3f;font-weight:900;font-size:12px;margin-top:3px}}
.hero-att{{height:281px;padding:28px 26px 0;text-align:center;background:radial-gradient(circle at 50% 0%,#159a8e 0%,#0b827b 44%,#006b3b 100%);color:white}}.hero-att .cal{{font-size:38px;color:var(--lime);margin-bottom:12px}}.hero-att h1{{margin:0;font-size:28px;line-height:1.05;font-weight:1000}}.hero-att p{{font-size:16px;line-height:1.45;font-weight:700;margin:17px 6px 0}}
.att-page{{min-height:100dvh;background:#fff;overflow:hidden}}.att-card{{margin:-55px 22px 0;background:white;border-radius:23px;padding:22px 18px 18px;box-shadow:var(--shadow);position:relative;z-index:2}}.sys{{font-size:12px;font-weight:1000;color:#009b4a;text-align:center}}.big-clock{{font-size:50px;font-weight:1000;text-align:center;letter-spacing:-2px;margin-top:10px;color:#070b22}}.date-line{{text-align:center;color:#979daf;font-size:14px;margin:6px 0 18px}}
.btn-main{{height:57px;border:0;border-radius:10px;background:#21b45d;color:white;font-size:18px;font-weight:1000;width:100%;box-shadow:0 6px 11px rgba(23,149,76,.2)}}.btn-row{{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin:14px 0}}.btn-soft{{height:56px;border:0;border-radius:10px;background:#f1f3f5;color:#46505f;font-weight:1000;font-size:16px}}.btn-out{{height:56px;border:0;border-radius:10px;background:#f1f3f5;color:#46505f;font-weight:1000;font-size:18px;width:100%}}
.map-title{{display:flex;align-items:center;justify-content:space-between;margin-top:25px;color:#838da0;font-size:13px;font-weight:1000}}.signal{{background:#d8f8df;color:#0c923e;border-radius:18px;padding:8px 13px;font-size:12px;font-weight:1000}}.map{{height:150px;border-radius:14px;border:1px solid #d6dde6;margin-top:10px;position:relative;overflow:hidden;background:#f4f6f9}}.road-a,.road-b{{position:absolute;background:#fbd577}}.road-a{{width:420px;height:15px;left:-35px;top:54px;transform:rotate(38deg)}}.road-b{{width:320px;height:10px;left:-35px;top:100px;background:#cfd5dd;transform:rotate(-2deg)}}.road-c{{position:absolute;width:14px;height:220px;background:#cfd5dd;left:200px;top:-20px;transform:rotate(28deg)}}.map:before,.map:after{{content:"";position:absolute;left:0;right:0;height:2px;background:#dbe1e8}}.map:before{{top:48px}}.map:after{{top:102px}}.pin{{position:absolute;left:50%;top:53%;transform:translate(-50%,-50%);width:35px;height:35px;border-radius:50% 50% 50% 0;background:#1775ff;transform:translate(-50%,-50%) rotate(-45deg)}}.pin:after{{content:"";position:absolute;width:13px;height:13px;border-radius:50%;background:white;left:11px;top:11px}}.bubble{{position:absolute;left:50%;top:38%;transform:translateX(-50%);background:white;border-radius:7px;padding:8px 12px;font-size:12px;box-shadow:0 3px 10px rgba(0,0,0,.16);white-space:nowrap}}.loc-note{{font-size:12px;color:#616b7b;margin:12px 4px 5px;font-weight:700}}.hist{{display:block;text-align:right;color:#006dff;font-weight:900;font-size:14px}}
.bottom-nav{{position:absolute;left:0;right:0;bottom:0;height:68px;background:white;border-top:1px solid #e7e9ed;display:grid;grid-template-columns:repeat(4,1fr);align-items:center;box-shadow:0 -6px 18px rgba(0,0,0,.04)}}.bottom-nav a{{text-align:center;color:#7d8594;font-size:11px;font-weight:800}}.bottom-nav a.active{{color:#00946a}}.bottom-nav span{{display:block;font-size:23px;margin-bottom:2px}}
.login-wrap{{min-height:100dvh;background:#fff;display:flex;flex-direction:column;align-items:center;padding:34px 22px}}.login-logo{{width:150px;height:118px;margin:12px auto 12px;display:grid;place-items:center;color:var(--green);font-size:58px;font-weight:1000}}.login-card{{width:100%;background:#fff;border:1px solid #e8e8e8;border-radius:8px;box-shadow:0 5px 12px rgba(0,0,0,.18);padding:18px 15px}}.login-card h2{{text-align:center;font-size:13px;font-weight:500;color:#565d66;margin:0 0 16px}}.label{{font-size:11px;color:#5b6470;margin:9px 0 5px}}.input{{width:100%;height:41px;border:1px solid #e1e5ea;border-radius:7px;padding:0 11px;font-weight:700;background:#fff}}.login-btn{{width:100%;height:45px;border:0;border-radius:7px;background:#159890;color:white;font-weight:900;margin-top:18px}}.link{{text-align:center;color:#4668ff;font-size:11px;font-weight:800;margin-top:14px}}.help{{margin-top:30px;text-align:center;font-size:10px;color:#9199a5}}.help b{{color:#08978f}}
.simple-page{{min-height:100dvh;background:#fff;padding-bottom:76px;position:relative}}.page-head{{height:130px;background:var(--green);color:white;text-align:center;padding:24px 18px;border-radius:0 0 30px 30px}}.page-head h1{{margin:8px 0 0;font-size:24px}}.back{{position:absolute;left:16px;top:38px;color:white;font-size:26px}}.list{{padding:18px}}.doc-row{{height:70px;border:1px solid #eceff3;border-radius:13px;margin-bottom:12px;display:grid;grid-template-columns:48px 1fr 20px;align-items:center;padding:0 14px;box-shadow:0 3px 8px rgba(0,0,0,.04)}}.doc-icon{{width:34px;height:34px;border-radius:10px;border:1px solid #8ddbd5;color:#009b94;display:grid;place-items:center}}.doc-title{{font-weight:900}}.doc-sub{{font-size:12px;color:#9aa2af;margin-top:3px}}.chev{{color:#8c94a0;font-size:24px}}
.sign-box{{margin:18px;background:#fff;border:2px dashed #dbe3ef;border-radius:14px;height:210px;position:relative}}#canvas{{width:100%;height:100%;touch-action:none}}.mini-btn{{border:0;border-radius:10px;background:#f1f4f7;height:44px;font-weight:900;color:#39424e}}.danger{{background:#feecec;color:#e93030}}.save{{background:#22b45e;color:#fff}}.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:0 18px 12px}}.grid1{{margin:0 18px 12px}}.grid1 button{{width:100%}}
.profile-card{{margin:18px;border-radius:20px;background:white;box-shadow:var(--shadow);padding:22px;text-align:center}}.profile-photo{{width:92px;height:92px;border-radius:50%;background:#edf4ef;color:var(--green);display:grid;place-items:center;font-size:60px;margin:0 auto 14px}}.kv{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:18px;text-align:left}}.kv div{{background:#f6f8f8;border-radius:13px;padding:13px}}.kv small{{display:block;color:#8791a0;font-weight:800}}.kv b{{display:block;margin-top:5px}}
.toast{{position:fixed;left:50%;top:16px;transform:translateX(-50%);background:#10251a;color:white;border-radius:999px;padding:11px 18px;font-weight:900;z-index:99;box-shadow:var(--shadow)}}
.wheel-modal{{position:fixed;inset:auto 0 0;z-index:50;background:rgba(0,0,0,.25);display:none;align-items:end;justify-content:center}}.wheel-modal.show{{display:flex}}.wheel-card{{width:100%;max-width:430px;background:white;border-radius:24px 24px 0 0;padding:14px 18px 22px;box-shadow:0 -12px 35px rgba(0,0,0,.22)}}.wheel-head{{display:flex;justify-content:space-between;align-items:center;font-weight:1000;margin-bottom:8px}}.wheel-head button{{border:0;background:white;color:#0aa25b;font-weight:1000;font-size:16px}}.wheels{{height:170px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;position:relative}}.wheels:before{{content:"";position:absolute;left:0;right:0;top:61px;height:48px;background:rgba(0,0,0,.04);border-radius:12px;pointer-events:none}}.wheel{{height:170px;overflow-y:auto;scroll-snap-type:y mandatory;text-align:center;padding:61px 0;scrollbar-width:none}}.wheel::-webkit-scrollbar{{display:none}}.wheel div{{height:48px;line-height:48px;font-size:24px;color:#b6bdc7;scroll-snap-align:center;font-weight:500}}.wheel div.sel{{color:#303642;font-size:28px}}
.err{{background:#fee2e2;color:#991b1b;border-radius:8px;padding:10px 12px;font-size:12px;margin-bottom:10px;font-weight:800}}
@media(min-width:800px){{.phone{{margin:18px 0;min-height:900px;border-radius:22px}}}}
"""

JS = """
function showToast(msg){let t=document.createElement('div');t.className='toast';t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.remove(),2200)}
function go(p){location.href=p}
function updateClock(){let el=document.getElementById('liveClock'); if(!el)return; let d=new Date(); let h=d.getHours(); let am=h<12?'a. m.':'p. m.'; h=h%12||12; el.textContent=String(h).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0')+' '+am;}
setInterval(updateClock,1000); document.addEventListener('DOMContentLoaded',updateClock);
async function marcar(tipo){
  let payload={tipo:tipo,hora:(document.getElementById('selectedTime')?.dataset.value||'')};
  function send(){fetch('/api/marcar',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).then(r=>r.json()).then(j=>{showToast(j.msg||'Registrado'); setTimeout(()=>location.reload(),800)}).catch(()=>showToast('No se pudo registrar'));}
  if(navigator.geolocation){navigator.geolocation.getCurrentPosition(pos=>{payload.lat=pos.coords.latitude;payload.lng=pos.coords.longitude;payload.precision=Math.round(pos.coords.accuracy);send();},()=>send(),{enableHighAccuracy:true,timeout:6000});}else send();
}
function openWheel(){let m=document.getElementById('wheelModal'); if(!m)return; m.classList.add('show'); syncWheel();}
function closeWheel(){document.getElementById('wheelModal')?.classList.remove('show')}
function syncWheel(){document.querySelectorAll('.wheel').forEach(w=>{let mid=w.scrollTop+85; [...w.children].forEach(c=>c.classList.toggle('sel', Math.abs((c.offsetTop+24)-mid)<25));});}
function setNowWheel(){let d=new Date(); setWheel('wh', d.getHours()%12||12); setWheel('wm', d.getMinutes()); setWheel('wa', d.getHours()<12?'AM':'PM'); syncWheel();}
function setWheel(id,val){let w=document.getElementById(id); if(!w)return; [...w.children].forEach(c=>{if(String(c.dataset.v)===String(val)){w.scrollTop=c.offsetTop-61;}})}
function applyWheel(){let h=pick('wh'), m=pick('wm'), a=pick('wa'); let val=String(h).padStart(2,'0')+':'+String(m).padStart(2,'0')+' '+(a==='AM'?'a. m.':'p. m.'); let e=document.getElementById('selectedTime'); if(e){e.dataset.value=val;e.textContent=val;} closeWheel();}
function pick(id){let w=document.getElementById(id), mid=w.scrollTop+85, best=w.children[0], bd=999; [...w.children].forEach(c=>{let d=Math.abs((c.offsetTop+24)-mid); if(d<bd){bd=d;best=c;}}); return best.dataset.v;}
document.addEventListener('scroll',e=>{if(e.target.classList&&e.target.classList.contains('wheel')) syncWheel()},true);
function initSignature(){let c=document.getElementById('canvas'); if(!c)return; let ctx=c.getContext('2d'), drawing=false; function resize(){let r=c.getBoundingClientRect(); c.width=r.width*devicePixelRatio; c.height=r.height*devicePixelRatio; ctx.scale(devicePixelRatio,devicePixelRatio); ctx.lineWidth=2; ctx.lineCap='round'; ctx.strokeStyle='#111';} resize(); window.addEventListener('resize',resize); function p(ev){let r=c.getBoundingClientRect(); let t=ev.touches?ev.touches[0]:ev; return {x:t.clientX-r.left,y:t.clientY-r.top};} c.addEventListener('pointerdown',ev=>{drawing=true;let q=p(ev);ctx.beginPath();ctx.moveTo(q.x,q.y)}); c.addEventListener('pointermove',ev=>{if(!drawing)return;let q=p(ev);ctx.lineTo(q.x,q.y);ctx.stroke()}); c.addEventListener('pointerup',()=>drawing=false); c.addEventListener('pointerleave',()=>drawing=false); window.clearSign=()=>ctx.clearRect(0,0,c.width,c.height); window.saveSign=()=>fetch('/api/firma',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({data:c.toDataURL('image/png')})}).then(r=>r.json()).then(j=>showToast(j.msg||'Firma guardada')); window.downloadSign=()=>{let a=document.createElement('a');a.href=c.toDataURL('image/png');a.download='firma_pa.png';a.click()};}
document.addEventListener('DOMContentLoaded',initSignature);
"""

def shell(body, title="P&A Mobile", page_class=""):
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><meta name="theme-color" content="{GREEN}"><title>{title}</title><style>{CSS}</style></head><body><div class="phone {page_class}">{body}</div><script>{JS}</script></body></html>"""

# ------------------------- Pages -------------------------
@app.route("/")
def index():
    if session.get("usuario"):
        return redirect(url_for("home"))
    body = f"""
    <div class="splash green-bg" onclick="location.href='/login'">
      <div class="logo-circle"><div class="logo-inner">▣</div></div>
      <div class="brand-title">P<span>&</span>A</div>
      <div class="progress"><i></i></div>
      <div class="splash-foot">{COMPANY}<br>{VERSION}</div>
    </div>"""
    return shell(body, "P&A", "green-bg")

@app.route("/login", methods=["GET","POST"])
def login():
    err = ""
    if request.method == "POST":
        usuario = re.sub(r"\D", "", request.form.get("usuario", ""))[-8:]
        password = request.form.get("password", "")
        c = conn(); r = c.execute("SELECT * FROM users WHERE usuario=?", (usuario,)).fetchone(); c.close()
        if r and check_password_hash(r["password_hash"], password):
            session["usuario"] = usuario
            return redirect(url_for("home"))
        err = "Usuario o contraseña incorrectos."
    body = f"""
    <div class="login-wrap">
      <div class="login-logo"><div><b style="font-size:46px">P<span style="color:{LIME}">&</span>A</b><div style="font-size:11px;text-align:center;color:{GREEN}">Gestión del Capital Humano</div></div></div>
      <form class="login-card" method="post">
        <h2>Inicia sesión para continuar</h2>
        {f'<div class="err">{err}</div>' if err else ''}
        <div class="label">Usuario</div><input class="input" name="usuario" value="11223344" maxlength="8" inputmode="numeric" required>
        <div class="label">Empresa</div><select class="input"><option>{COMPANY}</option></select>
        <div class="label">Contraseña</div><input class="input" name="password" value="123456" type="password" required>
        <button class="login-btn">Iniciar Sesión ↪</button>
        <div class="link">Olvidaste tu contraseña?</div>
      </form>
      <div class="help">¿Problemas para acceder?<br><b>Contactar a Mesa</b></div>
    </div>"""
    return shell(body, "Login P&A")

@app.route("/home")
@login_required
def home():
    u = current_user() or {}
    body = f"""
    <div class="top-green">
      <div class="pill-row"><div class="pill">🎧 Soporte</div><div class="pill">⚙ Config.</div></div>
      <div class="avatar">☻</div><div class="user-name">{esc(u.get('nombres','NISEDM01'))}</div>
      <div class="welcome-card"><b>Bienvenido(a)</b><p>Sistema de Asistencia</p></div>
    </div>
    <div class="content">
      <div class="tile-grid">
        <a class="tile" href="/asistencia"><div class="ico">☷</div><div>ASISTENCIA</div></a>
        <a class="tile" href="/documentos"><div class="ico">▧</div><div>DOCUMENTOS</div></a>
        <a class="tile" href="/firma"><div class="ico">✎</div><div>FIRMA</div></a>
        <a class="tile" href="/perfil"><div class="ico">◉</div><div>PERFIL</div></a>
      </div>
      <div class="leaves"><i></i><i></i><i></i></div>
      <div class="sync-card"><div class="sync-btn">↻</div><div><div class="sync-title">Sincronizar Tablas Maestras</div><div class="sync-sub">Actualizado hasta: 18/03/2025 09:31:12</div></div><a class="exit" href="/logout">⇥</a></div>
    </div>"""
    return shell(body, "P&A Home")

@app.route("/asistencia")
@login_required
def asistencia():
    c = conn(); rows = c.execute("SELECT * FROM marcas WHERE usuario=? AND fecha=? ORDER BY id DESC LIMIT 4", (session["usuario"], date.today().isoformat())).fetchall(); c.close()
    hist = "".join([f"<div class='doc-row' style='height:50px'><div class='doc-icon'>✓</div><div><div class='doc-title'>{esc(r['tipo'])}</div><div class='doc-sub'>{esc(r['hora'])}</div></div><div></div></div>" for r in rows])
    body = f"""
    <div class="att-page">
      <div class="status"><b>9:41</b><div class="icons"><i class="bar"></i><i class="bar"></i><i class="bar"></i><i class="wifi"></i></div></div>
      <div class="hero-att"><div class="cal">▦</div><h1>Registra tu asistencia fácil</h1><p>Registra tu ingreso, salida y refrigerio de forma rápida y segura desde tu celular.</p></div>
      <div class="att-card">
        <div class="sys">HORA ACTUAL DEL SISTEMA</div>
        <div id="liveClock" class="big-clock">{h12()}</div>
        <div class="date-line">{today_es()}</div>
        <button class="btn-main" onclick="marcar('INGRESO')">↪ &nbsp; Registrar Ingreso</button>
        <div class="btn-row"><button class="btn-soft" onclick="marcar('SALIDA REFRIGERIO')">Salida a<br>Refrigerio</button><button class="btn-soft" onclick="marcar('RETORNO REFRIGERIO')">Retorno de<br>Refrigerio</button></div>
        <button class="btn-out" onclick="marcar('SALIDA')">↩ &nbsp; Registrar Salida</button>
        <button class="btn-soft" style="width:100%;margin-top:12px;height:42px" onclick="openWheel()">🕘 Ajustar hora táctil: <span id="selectedTime" data-value="">Ahora</span></button>
        <div class="map-title"><span>UBICACIÓN ACTUAL</span><span class="signal">SEÑAL FUERTE</span></div>
        <div class="map"><i class="road-a"></i><i class="road-b"></i><i class="road-c"></i><div class="bubble">Tu ubicación actual</div><div class="pin"></div></div>
        <div class="loc-note">Ubicación detectada con precisión aproximada de 97 m.</div><a class="hist" href="#historial">Ver historial</a>
      </div>
      <div id="historial" class="list" style="padding-top:12px">{hist}</div>
      {bottom_nav('asistencia')}
      {wheel_modal()}
    </div>"""
    return shell(body, "Asistencia P&A")

def wheel_modal():
    hours = "".join([f"<div data-v='{i}'>{i:02d}</div>" for i in range(1,13)])
    mins = "".join([f"<div data-v='{i}'>{i:02d}</div>" for i in range(0,60)])
    ampm = "<div data-v='AM'>AM</div><div data-v='PM'>PM</div>"
    return f"""<div class="wheel-modal" id="wheelModal"><div class="wheel-card"><div class="wheel-head"><button onclick="closeWheel()">Cancelar</button><b>Seleccionar hora</b><button onclick="applyWheel()">OK</button></div><div class="wheels"><div class="wheel" id="wh" onscroll="syncWheel()">{hours}</div><div class="wheel" id="wm" onscroll="syncWheel()">{mins}</div><div class="wheel" id="wa" onscroll="syncWheel()">{ampm}</div></div><button class="btn-main" style="height:44px;margin-top:10px" onclick="setNowWheel()">Usar hora actual</button></div></div>"""

def bottom_nav(active):
    items = [("home","Home","⌘","/home"),("asistencia","Asistencia","▦","/asistencia"),("documentos","Documentos","▰","/documentos"),("perfil","Perfil","☻","/perfil")]
    return "<div class='bottom-nav'>" + "".join([f"<a class='{ 'active' if k==active else ''}' href='{href}'><span>{ic}</span>{txt}</a>" for k,txt,ic,href in items]) + "</div>"

@app.route("/documentos")
@login_required
def documentos():
    docs = ["Boletas Normales", "Constancias de Utilidades", "Boletas de Vacaciones", "Constancias de CTS", "Constancias de Liquidaciones", "Boletas de Gratificaciones", "Constancias de Gratificaciones"]
    rows = "".join([f"<div class='doc-row'><div class='doc-icon'>▧</div><div><div class='doc-title'>{d}</div><div class='doc-sub'>Disponible para descarga</div></div><div class='chev'>›</div></div>" for d in docs])
    body = f"""<div class="simple-page"><a class="back" href="/home">‹</a><div class="page-head"><div style="font-size:34px;color:{LIME}">▰</div><h1>Gestión de documentos</h1></div><div class="list">{rows}</div>{bottom_nav('documentos')}</div>"""
    return shell(body, "Documentos P&A")

@app.route("/firma")
@login_required
def firma():
    body = f"""<div class="simple-page"><a class="back" href="/home">‹</a><div class="page-head"><div style="font-size:34px;color:{LIME}">✎</div><h1>Gestión de Firma</h1></div><div style="padding:18px 18px 0;color:#8b94a2;font-weight:900;font-size:13px">FIRMA DIGITAL ACTUAL</div><div class="sign-box"><canvas id="canvas"></canvas></div><div class="grid2"><button class="mini-btn" onclick="downloadSign()">↥ Descargar Firma</button><button class="mini-btn">↧ Subir Imagen</button></div><div class="grid1"><button class="mini-btn danger" onclick="clearSign()">🗑 Limpiar Lienzo</button></div><div class="grid1"><button class="mini-btn save" onclick="saveSign()">▣ Guardar Cambios</button></div>{bottom_nav('perfil')}</div>"""
    return shell(body, "Firma P&A")

@app.route("/perfil")
@login_required
def perfil():
    u = current_user() or {}
    body = f"""<div class="simple-page"><a class="back" href="/home">‹</a><div class="page-head"><div style="font-size:34px;color:{LIME}">☻</div><h1>Perfil</h1></div><div class="profile-card"><div class="profile-photo">☻</div><h2>{esc(u.get('nombres'))}</h2><p style="color:#7f8795;font-weight:800">DNI {esc(u.get('usuario'))}</p><div class="kv"><div><small>Empresa</small><b>{COMPANY}</b></div><div><small>Área</small><b>{esc(u.get('area'))}</b></div><div><small>Cargo</small><b>{esc(u.get('cargo'))}</b></div><div><small>Estado</small><b>ACTIVO</b></div></div><a href="/logout" class="btn-soft" style="display:grid;place-items:center;margin-top:18px">Cerrar sesión</a></div>{bottom_nav('perfil')}</div>"""
    return shell(body, "Perfil P&A")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ------------------------- API -------------------------
@app.post("/api/marcar")
@login_required
def api_marcar():
    data = request.get_json(silent=True) or {}
    tipo = str(data.get("tipo") or "MARCA").upper()
    now = datetime.now()
    hora = data.get("hora") or h12(now)
    c = conn(); c.execute("""INSERT INTO marcas(usuario,tipo,fecha,hora,fecha_hora,lat,lng,precision,direccion)
                         VALUES(?,?,?,?,?,?,?,?,?)""",
                      (session["usuario"], tipo, date.today().isoformat(), hora, now.isoformat(sep=" ", timespec="seconds"),
                       str(data.get("lat", "")), str(data.get("lng", "")), str(data.get("precision", "")), "Ubicación actual"))
    c.commit(); c.close()
    return jsonify(ok=True, msg=f"{tipo.title()} registrado")

@app.post("/api/firma")
@login_required
def api_firma():
    data = request.get_json(silent=True) or {}
    c = conn(); c.execute("""INSERT INTO firmas(usuario,data_url,actualizado) VALUES(?,?,?)
                         ON CONFLICT(usuario) DO UPDATE SET data_url=excluded.data_url, actualizado=excluded.actualizado""",
                      (session["usuario"], data.get("data", ""), datetime.now().isoformat(sep=" ", timespec="seconds")))
    c.commit(); c.close()
    return jsonify(ok=True, msg="Firma guardada")

@app.route("/manifest.json")
def manifest():
    return jsonify({"name":"P&A Mobile","short_name":"P&A","start_url":"/","display":"standalone","background_color":"#ffffff","theme_color":GREEN,"icons":[]})

@app.route("/sw.js")
def sw():
    return Response("self.addEventListener('fetch',()=>{});", mimetype="application/javascript")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
