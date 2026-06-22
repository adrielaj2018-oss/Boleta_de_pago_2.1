# -*- coding: utf-8 -*-
"""
Nisira GCH Mobile - App web/PWA para Render
Módulos: Login, Home, Asistencia con ubicación, Firma digital, Documentos y Perfil.
Usuario demo: 11223344 / 123456
"""
import os, sqlite3, base64, re, json
from datetime import datetime, date
from functools import wraps
from flask import Flask, request, redirect, url_for, session, flash, jsonify, send_file, render_template_string, Response
from werkzeug.security import generate_password_hash, check_password_hash
from io import BytesIO

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR = os.getenv('PERSIST_DIR', '/data' if os.path.isdir('/data') else BASE_DIR)
os.makedirs(PERSIST_DIR, exist_ok=True)
DB_PATH = os.path.join(PERSIST_DIR, 'nisira_gch_mobile.db')
UPLOAD_DIR = os.path.join(PERSIST_DIR, 'documentos')
SIGN_DIR = os.path.join(PERSIST_DIR, 'firmas')
os.makedirs(UPLOAD_DIR, exist_ok=True); os.makedirs(SIGN_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'cambiar-en-render')
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024

TEAL = '#0f9693'; TEAL_DARK = '#063434'; GREEN = '#24b35b'; LIME = '#d8ff4f'; YELLOW = '#ffc51b'

def conn():
    c = sqlite3.connect(DB_PATH); c.row_factory = sqlite3.Row; return c

def now(): return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
def today(): return date.today().isoformat()
def clean_dni(v): return re.sub(r'\D', '', str(v or ''))[-8:]
def rowdict(r): return dict(r) if r else None

def db_exec(sql, params=(), one=False, all=False, commit=False):
    c = conn(); cur = c.cursor(); cur.execute(sql, params); data = None
    if one: data = cur.fetchone()
    if all: data = cur.fetchall()
    if commit: c.commit()
    cur.close(); c.close(); return data

def init_db():
    c = conn(); cur = c.cursor()
    cur.executescript('''
    CREATE TABLE IF NOT EXISTS empresas(id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT UNIQUE, ruc TEXT);
    CREATE TABLE IF NOT EXISTS usuarios(id INTEGER PRIMARY KEY AUTOINCREMENT, dni TEXT UNIQUE, nombres TEXT, empresa TEXT, password_hash TEXT, rol TEXT DEFAULT 'trabajador', cargo TEXT, area TEXT, foto TEXT, firma_path TEXT, creado_en TEXT);
    CREATE TABLE IF NOT EXISTS asistencias(id INTEGER PRIMARY KEY AUTOINCREMENT, dni TEXT, nombres TEXT, empresa TEXT, tipo TEXT, fecha TEXT, hora TEXT, fecha_hora TEXT, latitud TEXT, longitud TEXT, direccion TEXT, precision_gps TEXT, observacion TEXT);
    CREATE TABLE IF NOT EXISTS documentos(id INTEGER PRIMARY KEY AUTOINCREMENT, dni TEXT, categoria TEXT, titulo TEXT, periodo TEXT, filename TEXT, content_type TEXT, size_bytes INTEGER, creado_en TEXT);
    ''')
    cur.execute("INSERT OR IGNORE INTO empresas(nombre,ruc) VALUES(?,?)", ('NISIRA SYSTEMS S.A.C', '20600000000'))
    cur.execute("SELECT id FROM usuarios WHERE dni=?", ('11223344',))
    if not cur.fetchone():
        cur.execute("INSERT INTO usuarios(dni,nombres,empresa,password_hash,rol,cargo,area,creado_en) VALUES(?,?,?,?,?,?,?,?)",
                    ('11223344','OMAR AZABACHE','NISIRA SYSTEMS S.A.C', generate_password_hash('123456'),'admin','Analista RRHH','Gestión Humana', now()))
    demo_docs = [
        ('11223344','Boletas Normales','Boleta Febrero 2026','FEBRERO 2026','boleta_febrero_2026.pdf'),
        ('11223344','Empresa','Contrato Anexo','2026','contrato_anexo.pdf'),
        ('11223344','Empresa','Política de SST','2026','politica_sst.pdf'),
        ('11223344','Boletas Normales','Boleta Enero 2026','ENERO 2026','boleta_enero_2026.pdf'),
        ('11223344','Constancias de Utilidades','Constancia de Utilidades','2026','constancia_utilidades.pdf'),
        ('11223344','Constancias de CTS','Constancia de CTS','2026','constancia_cts.pdf'),
        ('11223344','Boletas de Gratificaciones','Boleta de Gratificación','2026','gratificacion_2026.pdf'),
    ]
    for d in demo_docs:
        cur.execute("SELECT id FROM documentos WHERE dni=? AND titulo=?", (d[0], d[2]))
        if not cur.fetchone():
            path = os.path.join(UPLOAD_DIR, d[4])
            if not os.path.exists(path):
                with open(path, 'wb') as f: f.write(b'%PDF-1.4\n% Demo PDF Nisira GCH\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Count 0>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF')
            cur.execute("INSERT INTO documentos(dni,categoria,titulo,periodo,filename,content_type,size_bytes,creado_en) VALUES(?,?,?,?,?,?,?,?)",
                        (d[0],d[1],d[2],d[3],d[4],'application/pdf',os.path.getsize(path),now()))
    c.commit(); c.close()

def login_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if not session.get('dni'): return redirect(url_for('login'))
        return fn(*a, **kw)
    return wrapper

def current_user():
    return rowdict(db_exec('SELECT * FROM usuarios WHERE dni=?', (session.get('dni'),), one=True))

BASE = r'''
<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover"><meta name="theme-color" content="#0f9693">
<link rel="manifest" href="/manifest.json"><title>{{title or 'Nisira GCH'}}</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css" rel="stylesheet">
<style>
:root{--teal:#0f9693;--dark:#042b2b;--green:#25b65c;--lime:#d9ff57;--yellow:#ffc51b;--muted:#6b7280;--line:#e8eef1;--bg:#f7f9fb}*{box-sizing:border-box}html,body{margin:0;font-family:Inter,system-ui,Segoe UI,Arial,sans-serif;background:linear-gradient(180deg,#0f9693 0%,#042b2b 100%);color:#111827}a{text-decoration:none;color:inherit}.stage{min-height:100vh;display:grid;place-items:center;padding:18px}.phone{width:min(390px,100%);min-height:780px;background:#fff;border-radius:38px;box-shadow:0 24px 60px rgba(0,0,0,.42);overflow:hidden;position:relative;border:10px solid #050505}.phone:before{content:'';position:absolute;top:0;left:50%;transform:translateX(-50%);width:120px;height:30px;background:#050505;border-radius:0 0 17px 17px;z-index:20}.status{height:32px;display:flex;align-items:center;justify-content:space-between;padding:8px 20px 0;font-size:13px;font-weight:800;position:relative;z-index:21}.status .icons{letter-spacing:2px}.content{min-height:695px;background:#fff;position:relative}.logo{display:grid;place-items:center;margin:34px auto 16px}.mark{width:112px;height:84px;position:relative}.mark .crown{position:absolute;left:19px;right:19px;bottom:19px;height:26px;border-radius:0 0 28px 28px;background:var(--teal)}.mark .head{position:absolute;width:22px;height:22px;border-radius:50%;top:2px;background:#fff}.mark .h1{left:12px;background:#25a9a6}.mark .h2{left:45px;width:26px;height:26px;background:var(--yellow)}.mark .h3{right:12px;background:#25a9a6}.brand{text-align:center;color:var(--teal);font-weight:900;font-style:italic;font-size:22px;line-height:1}.brand small{display:block;font-size:8px;color:var(--teal);font-style:normal}.card{background:#fff;border:1px solid var(--line);border-radius:13px;box-shadow:0 8px 22px rgba(15,35,45,.13);padding:20px}.login-card{margin:0 18px}.login-card h4{text-align:center;font-size:12px;font-weight:600;color:#667085;margin:0 0 18px}.label{font-size:11px;color:#667085;font-weight:700;margin:10px 0 6px}.input,.select{width:100%;height:42px;border:1px solid #e5e7eb;border-radius:8px;background:#fff;padding:0 12px;font-weight:700;color:#374151}.btn{height:45px;border:0;border-radius:8px;background:var(--teal);color:#fff;font-weight:900;width:100%;display:flex;align-items:center;justify-content:center;gap:8px;cursor:pointer}.btn.green{background:var(--green)}.btn.ghost{background:#f4f7f8;color:#374151}.btn.danger{background:#fff1f2;color:#ef4444}.link{color:#5577ff;font-size:11px;font-weight:800;text-align:center;margin-top:14px}.help{position:absolute;bottom:36px;left:0;right:0;text-align:center;color:#8b949e;font-size:10px}.help b{color:var(--teal)}.hero{background:linear-gradient(180deg,#0f9693,#0b7c79);color:white;padding:44px 22px 24px;border-radius:0 0 28px 28px}.hero .icon{color:var(--lime);font-size:28px;text-align:center;margin-bottom:8px}.hero h1{font-size:27px;line-height:1.08;text-align:center;margin:0;font-weight:1000;letter-spacing:-.5px}.hero p{text-align:center;font-size:14px;line-height:1.45;font-weight:800;margin:17px auto 0;max-width:290px}.mock{margin:-8px 20px 0;background:white;border-radius:28px;border:8px solid #060606;box-shadow:0 18px 36px rgba(0,0,0,.28);padding:22px 18px;min-height:430px}.clock{text-align:center}.clock .sys{font-size:9px;color:#0f9693;font-weight:900}.clock .time{font-size:48px;color:#02071b;font-weight:1000;letter-spacing:-2px}.clock .date{font-size:12px;color:#98a2b3}.grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}.map{height:132px;background:linear-gradient(45deg,#e8eef1 25%,#f8fafc 25% 50%,#e8eef1 50% 75%,#f8fafc 75%);border:1px solid #e5e7eb;border-radius:12px;position:relative;overflow:hidden}.pin{position:absolute;left:50%;top:48%;transform:translate(-50%,-50%);font-size:34px;color:#2e90fa}.badge{display:inline-block;background:#dcfae6;color:#17a34a;border-radius:999px;padding:6px 10px;font-size:10px;font-weight:900}.nav{height:66px;border-top:1px solid var(--line);background:#fff;display:grid;grid-template-columns:repeat(4,1fr);position:absolute;bottom:0;left:0;right:0}.nav a{display:grid;place-items:center;font-size:11px;color:#6b7280;font-weight:700}.nav i{font-size:20px}.nav .active{color:var(--teal)}.app-body{padding-bottom:78px;background:#fff;min-height:695px}.top-home{background:linear-gradient(135deg,#0f9693,#23aa9f);height:150px;color:white;padding:38px 22px 18px;border-radius:0 0 24px 24px;position:relative;overflow:hidden}.top-home:after{content:'$';position:absolute;right:-6px;top:-14px;font-size:100px;font-weight:1000;opacity:.15}.welcome{font-size:12px;opacity:.9}.welcome b{display:block;font-size:22px;line-height:1.1}.quick{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:-28px 18px 16px;position:relative}.qcard{background:white;border:1px solid var(--line);border-radius:13px;box-shadow:0 10px 22px rgba(15,35,45,.1);padding:13px 10px;min-height:82px}.qcard i{font-size:24px;border:1px solid #dbeafe;border-radius:8px;color:#2563eb;padding:4px}.qcard:nth-child(2) i{color:#f59e0b;border-color:#fed7aa}.qcard:nth-child(3) i{color:#2563eb}.qcard b{display:block;margin-top:11px;font-size:12px}.qcard small{color:#6b7280;font-size:10px}.section-title{display:flex;justify-content:space-between;align-items:center;margin:18px 18px 9px;color:#98a2b3;font-size:12px;font-weight:900;letter-spacing:.5px}.doc-row{margin:10px 18px;border:1px solid var(--line);border-radius:12px;box-shadow:0 6px 15px rgba(15,35,45,.06);padding:12px;display:grid;grid-template-columns:45px 1fr 20px;align-items:center}.doc-ico{width:35px;height:35px;border-radius:9px;background:#ecfdf3;color:#65b32e;display:grid;place-items:center;font-weight:900;font-size:11px;border:1px solid #bbf7d0}.doc-row:nth-child(odd) .doc-ico{background:#fff7ed;color:#f97316;border-color:#fed7aa}.doc-row b{font-size:13px}.doc-row small{display:block;color:#98a2b3;font-size:10px}.page-title{height:76px;display:flex;align-items:center;gap:10px;padding:35px 18px 12px;font-weight:1000}.page-title i{font-size:24px}.doc-menu{padding:8px 20px}.menu-row{display:grid;grid-template-columns:36px 1fr 20px;align-items:center;padding:14px 0;font-size:15px}.menu-row .mi{width:25px;height:25px;border:1px solid #a5e3e4;border-radius:7px;color:var(--teal);display:grid;place-items:center}.firma-box{margin:28px 18px 16px}.canvas-wrap{border:2px dashed #d9e4ed;border-radius:12px;height:170px;position:relative;background:#fcfcfd}.canvas-wrap canvas{width:100%;height:100%;touch-action:none}.draw-label{position:absolute;right:12px;top:10px;border:1px solid #dce7f0;background:white;border-radius:8px;font-size:11px;font-weight:800;padding:7px}.form-pad{padding:0 18px}.profile{padding:22px 18px}.avatar{width:82px;height:82px;border-radius:50%;background:#e6fffb;color:var(--teal);display:grid;place-items:center;font-size:40px;margin:0 auto 10px}.center{text-align:center}.dark-section{background:linear-gradient(180deg,#0b5657,#020707);color:#fff;text-align:center;padding:40px 24px;min-height:695px}.dark-section .bigicon{font-size:36px;color:var(--lime);margin-bottom:18px}.dark-section h2{font-size:28px;line-height:1.1;margin:0 0 14px;font-weight:1000}.dark-section p{font-size:15px;line-height:1.5;font-weight:800}.splash{min-height:695px;background:linear-gradient(rgba(15,150,147,.78),rgba(15,150,147,.78)),linear-gradient(135deg,#d9e8e8,#8ba0a2);display:grid;place-items:center;color:#fff}.splash .brand{color:white;font-size:24px}.splash .brand small{color:white}.toast{position:fixed;top:18px;left:50%;transform:translateX(-50%);z-index:1000;background:#111827;color:white;border-radius:10px;padding:10px 14px;font-size:12px;font-weight:800}.muted{color:#98a2b3}.row{display:flex;gap:10px}.row>*{flex:1}@media(max-width:500px){.stage{padding:0}.phone{width:100%;min-height:100vh;border-radius:0;border:0}.content,.app-body,.dark-section,.splash{min-height:calc(100vh - 32px)}.phone:before{border-radius:0 0 16px 16px}}
</style></head><body><div class="stage"><div class="phone"><div class="status"><span>9:41</span><span class="icons">▮▮ WiFi ▰</span></div><div class="content">
{% with msgs=get_flashed_messages(with_categories=true) %}{% if msgs %}{% for c,m in msgs %}<div class="toast">{{m}}</div>{% endfor %}{% endif %}{% endwith %}
{{body|safe}}</div></div></div><script>setTimeout(()=>document.querySelectorAll('.toast').forEach(e=>e.remove()),3200);if('serviceWorker'in navigator){navigator.serviceWorker.register('/sw.js').catch(()=>{})}</script>{{script|safe}}</body></html>
'''

def render(body, title='Nisira GCH', script=''):
    return render_template_string(BASE, body=body, title=title, script=script)

@app.route('/')
def index(): return redirect(url_for('home' if session.get('dni') else 'login'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        dni = clean_dni(request.form.get('dni')); password = request.form.get('password',''); empresa = request.form.get('empresa','')
        u = rowdict(db_exec('SELECT * FROM usuarios WHERE dni=?', (dni,), one=True))
        if u and check_password_hash(u['password_hash'], password):
            session['dni']=u['dni']; session['nombres']=u['nombres']; session['empresa']=empresa or u['empresa']; return redirect(url_for('home'))
        flash('Usuario o contraseña incorrectos', 'danger')
    empresas = db_exec('SELECT nombre FROM empresas ORDER BY nombre', all=True)
    opts = ''.join([f"<option>{e['nombre']}</option>" for e in empresas])
    body = f'''
    <div class="logo"><div><div class="mark"><span class="head h1"></span><span class="head h2"></span><span class="head h3"></span><div class="crown"></div></div><div class="brand">Nisira GCH<small>Gestión del Capital Humano</small></div></div></div>
    <form class="card login-card" method="post"><h4>Inicia sesión para continuar</h4><div class="label">Usuario</div><input class="input" name="dni" value="11223344" inputmode="numeric" maxlength="8" required><div class="label">Empresa</div><select class="select" name="empresa">{opts}</select><div class="label">Contraseña</div><input class="input" name="password" type="password" value="123456" required><button class="btn" style="margin-top:18px">Iniciar Sesión <i class="bi bi-box-arrow-in-right"></i></button><div class="link">Olvidaste tu contraseña?</div></form><div class="help">¿Problemas para acceder?<br><b>Contacta a Nisira</b></div>'''
    return render(body, 'Login')

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/splash')
def splash():
    body = '<div class="splash"><div><div class="mark" style="margin:auto"><span class="head h1"></span><span class="head h2"></span><span class="head h3"></span><div class="crown" style="background:#fff"></div></div><div class="brand">Nisira GCH<small>Gestión del Capital Humano</small></div></div></div>'
    return render(body, 'Splash')

@app.route('/home')
@login_required
def home():
    u=current_user(); docs=db_exec('SELECT * FROM documentos WHERE dni=? ORDER BY creado_en DESC LIMIT 4',(u['dni'],),all=True)
    rows=''.join([f"<a class='doc-row' href='/documento/{d['id']}'><div class='doc-ico'>{d['categoria'][:2].upper()}</div><div><b>{d['titulo']}</b><small>Subido hace días · {round((d['size_bytes'] or 0)/1024)} KB</small></div><i class='bi bi-chevron-right muted'></i></a>" for d in docs])
    body=f'''<div class="app-body"><div class="top-home"><div class="welcome">Hola,<b>{u['nombres'].split()[0]}</b></div></div><div class="quick"><a class="qcard" href="/asistencia"><i class="bi bi-calendar-check"></i><b>Asistencia</b><small>Registrar mi asistencia</small></a><a class="qcard" href="/documentos"><i class="bi bi-person-badge"></i><b>Personales</b><small>Ver archivos</small></a><a class="qcard" href="/firma"><i class="bi bi-pen"></i><b>Firma</b><small>Gestionar</small></a></div><div class="section-title"><span>ÚLTIMOS DOCUMENTOS</span><a href="/documentos" style="color:#2563eb">Ver todo</a></div>{rows}{nav('home')}</div>'''
    return render(body,'Home')

def nav(active):
    items=[('home','Home','bi-grid-fill','/home'),('asistencia','Asistencia','bi-calendar3','/asistencia'),('documentos','Documentos','bi-folder-fill','/documentos'),('perfil','Perfil','bi-person-circle','/perfil')]
    return '<div class="nav">'+''.join([f"<a class='{ 'active' if active==k else ''}' href='{url}'><span><i class='bi {ic}'></i>{txt}</span></a>" for k,txt,ic,url in items])+'</div>'

@app.route('/asistencia')
@login_required
def asistencia():
    body='''<div class="app-body"><div class="hero"><div class="icon"><i class="bi bi-calendar3"></i></div><h1>Registra tu asistencia fácil</h1><p>Registra tu ingreso, salida y refrigerio de forma rápida y segura desde tu celular.</p></div><div class="mock"><div class="clock"><div class="sys">HORA ACTUAL DEL SISTEMA</div><div class="time" id="reloj">06:00 PM</div><div class="date" id="fechaTxt"></div></div><form id="frmMarcacion"><input type="hidden" name="latitud" id="lat"><input type="hidden" name="longitud" id="lng"><input type="hidden" name="direccion" id="dir"><button class="btn green" name="tipo" value="INGRESO"><i class="bi bi-box-arrow-in-right"></i> Registrar Ingreso</button><div class="grid2" style="margin-top:12px"><button class="btn ghost" name="tipo" value="SALIDA_REFRIGERIO">Salida a Refrigerio</button><button class="btn ghost" name="tipo" value="RETORNO_REFRIGERIO">Retorno de Refrigerio</button></div><button class="btn ghost" style="margin-top:12px" name="tipo" value="SALIDA"><i class="bi bi-box-arrow-left"></i> Registrar Salida</button></form><div style="display:flex;justify-content:space-between;margin-top:22px"><b style="font-size:11px;color:#98a2b3">UBICACIÓN ACTUAL</b><span class="badge">SEÑAL FUERTE</span></div><div class="map"><div class="pin"><i class="bi bi-geo-alt-fill"></i></div></div><small class="muted" id="gpsTxt">Obteniendo ubicación...</small></div>'''+nav('asistencia')+'</div>'
    script='''<script>function tick(){const d=new Date();document.getElementById('reloj').textContent=d.toLocaleTimeString('es-PE',{hour:'2-digit',minute:'2-digit'});document.getElementById('fechaTxt').textContent=d.toLocaleDateString('es-PE',{weekday:'long',day:'2-digit',month:'long',year:'numeric'});}tick();setInterval(tick,1000);if(navigator.geolocation){navigator.geolocation.getCurrentPosition(p=>{lat.value=p.coords.latitude;lng.value=p.coords.longitude;gpsTxt.textContent='Tu ubicación actual fue detectada con precisión aproximada de '+Math.round(p.coords.accuracy)+' m.'},()=>gpsTxt.textContent='No se pudo obtener ubicación. Activa GPS.')}});frmMarcacion.addEventListener('submit',async e=>{e.preventDefault();const fd=new FormData(e.submitter.form);fd.set('tipo',e.submitter.value);const r=await fetch('/api/marcar',{method:'POST',body:fd});const j=await r.json();alert(j.msg);});</script>'''
    return render(body,'Asistencia',script)

@app.post('/api/marcar')
@login_required
def api_marcar():
    u=current_user(); tipo=request.form.get('tipo','INGRESO')
    db_exec('INSERT INTO asistencias(dni,nombres,empresa,tipo,fecha,hora,fecha_hora,latitud,longitud,direccion,precision_gps) VALUES(?,?,?,?,?,?,?,?,?,?,?)',
            (u['dni'],u['nombres'],u['empresa'],tipo,today(),datetime.now().strftime('%H:%M:%S'),now(),request.form.get('latitud'),request.form.get('longitud'),request.form.get('direccion'),''),commit=True)
    return jsonify(ok=True,msg=f'Marcación {tipo.replace("_"," ")} registrada correctamente.')

@app.route('/firma', methods=['GET','POST'])
@login_required
def firma():
    u=current_user()
    if request.method=='POST':
        data=request.form.get('firma_data','')
        if data.startswith('data:image'):
            raw=base64.b64decode(data.split(',',1)[1]); fn=f"firma_{u['dni']}.png"; path=os.path.join(SIGN_DIR,fn); open(path,'wb').write(raw)
            db_exec('UPDATE usuarios SET firma_path=? WHERE dni=?',(fn,u['dni']),commit=True); flash('Firma guardada correctamente','ok'); return redirect(url_for('firma'))
    img=''
    if u.get('firma_path') and os.path.exists(os.path.join(SIGN_DIR,u['firma_path'])):
        b64=base64.b64encode(open(os.path.join(SIGN_DIR,u['firma_path']),'rb').read()).decode(); img=f"<img src='data:image/png;base64,{b64}' style='max-width:100%;max-height:150px'>"
    body=f'''<div class="app-body"><div class="page-title"><a href="/home"><i class="bi bi-chevron-left"></i></a><span>GESTIÓN DE FIRMA</span></div><form method="post" class="firma-box"><div class="label">FIRMA DIGITAL ACTUAL</div><div class="canvas-wrap"><div class="draw-label"><i class="bi bi-pencil"></i> Dibujar</div><canvas id="canvas"></canvas><div id="oldFirma" style="position:absolute;inset:10px;display:grid;place-items:center;pointer-events:none">{img}</div></div><input type="hidden" name="firma_data" id="firma_data"><div class="row" style="margin-top:18px"><button type="button" class="btn ghost" onclick="document.getElementById('fileFirma').click()"><i class="bi bi-upload"></i> Subir Imagen</button><a class="btn ghost" href="/firma/descargar"><i class="bi bi-download"></i> Descargar Firma</a></div><input id="fileFirma" type="file" accept="image/*" style="display:none"><button type="button" class="btn danger" style="margin-top:16px" onclick="clearCanvas()"><i class="bi bi-trash"></i> Limpiar Lienzo</button><button class="btn green" style="margin-top:16px"><i class="bi bi-save-fill"></i> Guardar Cambios</button></form>{nav('perfil')}</div>'''
    script='''<script>const c=document.getElementById('canvas'),ctx=c.getContext('2d');function resize(){c.width=c.offsetWidth*2;c.height=c.offsetHeight*2;ctx.lineWidth=3;ctx.lineCap='round';ctx.strokeStyle='#111'}resize();let draw=false,last=null;function pos(e){let r=c.getBoundingClientRect(),p=e.touches?e.touches[0]:e;return{x:(p.clientX-r.left)*2,y:(p.clientY-r.top)*2}};c.addEventListener('pointerdown',e=>{draw=true;last=pos(e);oldFirma.style.display='none'});c.addEventListener('pointermove',e=>{if(!draw)return;let p=pos(e);ctx.beginPath();ctx.moveTo(last.x,last.y);ctx.lineTo(p.x,p.y);ctx.stroke();last=p});window.addEventListener('pointerup',()=>draw=false);function clearCanvas(){ctx.clearRect(0,0,c.width,c.height);oldFirma.style.display='none'}document.querySelector('form').addEventListener('submit',()=>firma_data.value=c.toDataURL('image/png'));fileFirma.onchange=e=>{let f=e.target.files[0];if(!f)return;let img=new Image();img.onload=()=>{clearCanvas();ctx.drawImage(img,15,15,c.width-30,c.height-30);firma_data.value=c.toDataURL('image/png')};img.src=URL.createObjectURL(f)};</script>'''
    return render(body,'Firma',script)

@app.route('/firma/descargar')
@login_required
def firma_descargar():
    u=current_user();
    if not u.get('firma_path'): flash('Aún no tienes firma guardada','bad'); return redirect(url_for('firma'))
    return send_file(os.path.join(SIGN_DIR,u['firma_path']), as_attachment=True, download_name=u['firma_path'])

@app.route('/documentos')
@login_required
def documentos():
    cats=['Boletas Normales','Constancias de Utilidades','Boletas de Vacaciones','Constancias de CTS','Constancias de Liquidaciones','Boletas de Gratificaciones','Constancias de Gratificaciones']
    rows=''.join([f"<a class='menu-row' href='/documentos/{c}'><span class='mi'><i class='bi bi-file-earmark-text'></i></span><span>{c}</span><i class='bi bi-chevron-right muted'></i></a>" for c in cats])
    body=f'''<div class="app-body"><div class="top-home" style="height:105px"><b style="font-size:20px">Gestión de documentos</b><small style="display:block;margin-top:7px">Accede a tus boletas, constancias y documentos laborales.</small></div><div class="doc-menu">{rows}</div>{nav('documentos')}</div>'''
    return render(body,'Documentos')

@app.route('/documentos/<categoria>')
@login_required
def documentos_cat(categoria):
    u=current_user(); docs=db_exec('SELECT * FROM documentos WHERE dni=? AND categoria=? ORDER BY creado_en DESC',(u['dni'],categoria),all=True)
    rows=''.join([f"<a class='doc-row' href='/documento/{d['id']}'><div class='doc-ico'>{d['categoria'][:2].upper()}</div><div><b>{d['titulo']}</b><small>{d['periodo']} · {round((d['size_bytes'] or 0)/1024)} KB</small></div><i class='bi bi-chevron-right muted'></i></a>" for d in docs]) or '<p class="center muted">No hay documentos en esta categoría.</p>'
    return render(f'<div class="app-body"><div class="page-title"><a href="/documentos"><i class="bi bi-chevron-left"></i></a><span>{categoria}</span></div>{rows}{nav("documentos")}</div>',categoria)

@app.route('/documento/<int:doc_id>')
@login_required
def documento(doc_id):
    d=rowdict(db_exec('SELECT * FROM documentos WHERE id=? AND dni=?',(doc_id,session['dni']),one=True))
    if not d: flash('Documento no encontrado','bad'); return redirect(url_for('documentos'))
    return send_file(os.path.join(UPLOAD_DIR,d['filename']), as_attachment=True, download_name=d['filename'], mimetype=d['content_type'] or 'application/pdf')

@app.route('/perfil')
@login_required
def perfil():
    u=current_user(); total=db_exec('SELECT COUNT(*) c FROM asistencias WHERE dni=?',(u['dni'],),one=True)['c']
    body=f'''<div class="app-body"><div class="page-title"><span>Perfil</span></div><div class="profile center"><div class="avatar"><i class="bi bi-person-fill"></i></div><h2>{u['nombres']}</h2><p class="muted">DNI {u['dni']} · {u['empresa']}</p><div class="card" style="text-align:left"><p><b>Área:</b> {u['area']}</p><p><b>Cargo:</b> {u['cargo']}</p><p><b>Marcaciones:</b> {total}</p></div><a class="btn danger" style="margin-top:18px" href="/logout">Cerrar sesión</a></div>{nav('perfil')}</div>'''
    return render(body,'Perfil')

@app.route('/promo/<tipo>')
def promo(tipo):
    if tipo=='documentos': icon='bi-file-earmark-text-fill'; title='Gestión de documentos'; txt='Accede a tus boletas, constancias y documentos laborales de forma rápida, organizada y segura.'
    elif tipo=='todo': icon='bi-grid-fill'; title='Todo en un solo lugar'; txt='Accede rápidamente a tus documentos y funciones más importantes desde el inicio.'
    else: icon='bi-calendar3'; title='Registra tu asistencia fácil'; txt='Registra tu ingreso, salida y refrigerio de forma rápida y segura desde tu celular.'
    return render(f'<div class="dark-section"><div class="bigicon"><i class="bi {icon}"></i></div><h2>{title}</h2><p>{txt}</p></div>', title)

@app.route('/manifest.json')
def manifest():
    return jsonify(name='Nisira GCH',short_name='NisiraGCH',start_url='/login',display='standalone',background_color='#ffffff',theme_color='#0f9693',icons=[])

@app.route('/sw.js')
def sw():
    return Response("self.addEventListener('install',e=>self.skipWaiting());self.addEventListener('fetch',e=>{});", mimetype='text/javascript')

if __name__ == '__main__':
    init_db(); app.run(host='0.0.0.0', port=int(os.getenv('PORT','5000')), debug=True)
else:
    init_db()
