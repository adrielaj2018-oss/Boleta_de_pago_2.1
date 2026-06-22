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
from jinja2 import DictLoader

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

app = Flask(__name__, static_folder=None)
app.secret_key = os.getenv("SECRET_KEY", "cambiar-clave-en-render")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

# ========================= PLANTILLAS Y ESTÁTICOS EMBEBIDOS =========================
TEMPLATES = {'base.html': '<!doctype html>\n<html lang="es">\n<head>\n  <meta charset="utf-8">\n  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n  <meta name="theme-color" content="#0f9693">\n  <title>{% block title %}Nisira GCH{% endblock %}</title>\n  <link rel="manifest" href="{{ url_for(\'manifest\') }}">\n  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css">\n  <link rel="stylesheet" href="{{ url_for(\'static\', filename=\'css/app.css\') }}">\n</head>\n<body>\n  <div class="stage">\n    <div class="phone">\n      <div class="status"><span>9:41</span><span class="icons">▮▮▮ ᯤ ▰</span></div>\n      <div class="screen">\n        {% with messages = get_flashed_messages(with_categories=true) %}\n          {% if messages %}\n            <div class="flashbox">\n              {% for cat,msg in messages %}<div class="flash {{ cat }}">{{ msg }}</div>{% endfor %}\n            </div>\n          {% endif %}\n        {% endwith %}\n        {% block content %}{% endblock %}\n      </div>\n    </div>\n  </div>\n  <script src="{{ url_for(\'static\', filename=\'js/app.js\') }}"></script>\n  {% block scripts %}{% endblock %}\n</body>\n</html>\n', 'partials.html': '{% macro logo() %}\n<div class="logo">\n  <div class="mark"><span class="head h1"></span><span class="head h2"></span><span class="head h3"></span><span class="crown"></span></div>\n  <div class="brand">Nisira GCH<small>Gestión del Capital Humano</small></div>\n</div>\n{% endmacro %}\n\n{% macro nav(active) %}\n<nav class="nav">\n  <a class="{{ \'active\' if active==\'home\' else \'\' }}" href="{{ url_for(\'home\') }}"><i class="bi bi-grid-fill"></i><span>Home</span></a>\n  <a class="{{ \'active\' if active==\'asistencia\' else \'\' }}" href="{{ url_for(\'asistencia\') }}"><i class="bi bi-calendar3"></i><span>Asistencia</span></a>\n  <a class="{{ \'active\' if active==\'documentos\' else \'\' }}" href="{{ url_for(\'documentos\') }}"><i class="bi bi-folder-fill"></i><span>Documentos</span></a>\n  <a class="{{ \'active\' if active==\'perfil\' else \'\' }}" href="{{ url_for(\'perfil\') }}"><i class="bi bi-person-circle"></i><span>Perfil</span></a>\n</nav>\n{% endmacro %}\n', 'login.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import logo %}\n{% block title %}Login - Nisira GCH{% endblock %}\n{% block content %}\n<div class="login-screen">\n  {{ logo() }}\n  <form class="login-card card" method="post">\n    <h4>Inicia sesión para continuar</h4>\n    <label>Usuario</label>\n    <input class="input" name="dni" value="11223344" maxlength="8" inputmode="numeric" required>\n    <label>Empresa</label>\n    <select class="input" name="empresa" required>\n      {% for e in empresas %}<option>{{ e.nombre }}</option>{% endfor %}\n    </select>\n    <label>Contraseña</label>\n    <div class="passbox"><input id="pass" class="input" name="password" type="password" value="123456" required><button type="button" onclick="togglePass()"><i class="bi bi-eye-slash"></i></button></div>\n    <button class="btn teal" style="margin-top:18px">Iniciar Sesión <i class="bi bi-box-arrow-in-right"></i></button>\n    <div class="recover">Olvidaste tu contraseña?</div>\n  </form>\n  <div class="help">¿Problemas para acceder?<br><b>Contactar a Nisira</b></div>\n</div>\n{% endblock %}\n', 'splash.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import logo %}\n{% block content %}\n<div class="splash">\n  <div class="splash-overlay"></div>\n  <div class="splash-logo">{{ logo() }}</div>\n</div>\n<script>setTimeout(()=>location.href="{{ url_for(\'home\') }}", 1200)</script>\n{% endblock %}\n', 'home.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import nav %}\n{% block content %}\n<div class="app-body">\n  <section class="home-hero">\n    <div class="hello">Hola, <b>{{ current_user.nombres.split()[0] }}</b></div>\n    <p>Gestiona tu información laboral desde un solo lugar.</p>\n  </section>\n  <section class="quick-grid">\n    <a class="feature big" href="{{ url_for(\'asistencia\') }}"><i class="bi bi-calendar-check"></i><b>Asistencia</b><span>Registrar mi asistencia</span></a>\n    <a class="feature" href="{{ url_for(\'documentos\') }}"><i class="bi bi-building"></i><b>Empresa</b><span>Ver documentación</span></a>\n    <a class="feature" href="{{ url_for(\'documentos\') }}"><i class="bi bi-person-badge"></i><b>Personales</b><span>Ver Archivos</span></a>\n    <a class="feature" href="{{ url_for(\'documentos\') }}"><i class="bi bi-clock-history"></i><b>Vacaciones</b><span>Solicitar</span></a>\n  </section>\n  <div class="section-title"><span>ÚLTIMOS DOCUMENTOS</span><a href="{{ url_for(\'documentos\') }}">Ver todo</a></div>\n  <section class="list">\n    {% for d in docs %}\n    <a class="doc-row" href="{{ url_for(\'descargar_documento\', doc_id=d.id) }}">\n      <span class="doc-badge">{{ d.categoria[:2].upper() }}</span>\n      <span><b>{{ d.titulo }}</b><small>Subido el {{ d.creado_en[:10] }} · {{ (d.size_bytes/1024)|round|int }} KB</small></span>\n      <i class="bi bi-chevron-right"></i>\n    </a>\n    {% endfor %}\n  </section>\n  {{ nav(active) }}\n</div>\n{% endblock %}\n', 'asistencia.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import nav %}\n{% block content %}\n<div class="app-body">\n  <section class="assist-head">\n    <div class="promo-icon"><i class="bi bi-calendar3"></i></div>\n    <h1>Registra tu asistencia fácil</h1>\n    <p>Registra tu ingreso, salida y refrigerio de forma rápida y segura desde tu celular.</p>\n  </section>\n  <section class="assist-card tilted">\n    <div class="clock"><span>HORA ACTUAL DEL SISTEMA</span><b id="liveTime">06:00 PM</b><small id="liveDate">Martes, 10 de Marzo de 2026</small></div>\n    <form id="frmMarcacion">\n      <input type="hidden" name="latitud" id="lat"><input type="hidden" name="longitud" id="lng"><input type="hidden" name="direccion" id="dir"><input type="hidden" name="precision_gps" id="acc">\n      <button class="btn green" name="tipo" value="INGRESO"><i class="bi bi-box-arrow-in-right"></i> Registrar Ingreso</button>\n      <div class="two"><button class="btn soft" name="tipo" value="SALIDA_REFRIGERIO">Salida a Refrigerio</button><button class="btn soft" name="tipo" value="RETORNO_REFRIGERIO">Retorno de Refrigerio</button></div>\n      <button class="btn soft" name="tipo" value="SALIDA"><i class="bi bi-box-arrow-left"></i> Registrar Salida</button>\n    </form>\n    <div class="gps-line"><b>UBICACIÓN ACTUAL</b><span>SEÑAL FUERTE</span></div>\n    <div class="map"><div class="roads"></div><i class="bi bi-geo-alt-fill pin"></i><em>Tu ubicación actual</em></div>\n    <p class="address" id="gpsText">Obteniendo ubicación actual...</p>\n    <a class="history" href="#historial">Ver historial</a>\n  </section>\n  <section class="historial" id="historial">\n    <b>REPORTE DEL DÍA</b>\n    {% for a in ultimas %}<div class="mini-row"><span>{{ a.tipo.replace(\'_\',\' \') }}</span><small>{{ a.fecha }} {{ a.hora }}</small></div>{% else %}<small>Aún no hay marcaciones.</small>{% endfor %}\n  </section>\n  {{ nav(active) }}\n</div>\n{% endblock %}\n{% block scripts %}<script>initAsistencia();</script>{% endblock %}\n', 'firma.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import nav %}\n{% block content %}\n<div class="app-body">\n  <div class="page-title"><a href="{{ url_for(\'home\') }}"><i class="bi bi-chevron-left"></i></a><b>GESTIÓN DE FIRMA</b></div>\n  <form method="post" class="signature-form">\n    <label>FIRMA DIGITAL ACTUAL</label>\n    <div class="canvas-wrap">\n      <button class="draw-pill" type="button"><i class="bi bi-pencil-fill"></i> Dibujar</button>\n      <canvas id="signatureCanvas"></canvas>\n      {% if firma_b64 %}<img id="oldSignature" src="data:image/png;base64,{{ firma_b64 }}">{% endif %}\n    </div>\n    <input type="hidden" name="firma_data" id="firmaData">\n    <input type="file" id="uploadSignature" accept="image/*" hidden>\n    <div class="two"><button type="button" class="btn soft" onclick="uploadSignature.click()"><i class="bi bi-upload"></i> Subir Imagen</button><a class="btn soft" href="{{ url_for(\'descargar_firma\') }}"><i class="bi bi-download"></i> Descargar Firma</a></div>\n    <button type="button" class="btn danger" onclick="clearSignature()"><i class="bi bi-trash"></i> Limpiar Lienzo</button>\n    <button class="btn green"><i class="bi bi-folder-check"></i> Guardar Cambios</button>\n  </form>\n  {{ nav(active) }}\n</div>\n{% endblock %}\n{% block scripts %}<script>initFirma();</script>{% endblock %}\n', 'documentos.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import nav %}\n{% block content %}\n<div class="app-body docs-screen">\n  <div class="docs-card">\n    <div class="docs-head"><b>Centro de Pago</b><i class="bi bi-cash-coin"></i></div>\n    {% for c,icon in cats %}\n    <a class="menu-row" href="{{ url_for(\'documentos_categoria\', categoria=c) }}"><span><i class="bi {{ icon }}"></i></span><b>{{ c }}</b><i class="bi bi-chevron-right"></i></a>\n    {% endfor %}\n  </div>\n  <section class="dark-caption"><i class="bi bi-file-earmark-text-fill"></i><h1>Gestión de documentos</h1><p>Accede a tus boletas, constancias y documentos laborales de forma rápida, organizada y segura.</p></section>\n  {{ nav(active) }}\n</div>\n{% endblock %}\n', 'documentos_categoria.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import nav %}\n{% block content %}\n<div class="app-body">\n  <div class="page-title"><a href="{{ url_for(\'documentos\') }}"><i class="bi bi-chevron-left"></i></a><b>{{ categoria }}</b></div>\n  <section class="list padded">\n  {% for d in docs %}\n    <a class="doc-row" href="{{ url_for(\'descargar_documento\', doc_id=d.id) }}"><span class="doc-badge">{{ d.categoria[:2].upper() }}</span><span><b>{{ d.titulo }}</b><small>{{ d.periodo }} · {{ (d.size_bytes/1024)|round|int }} KB</small></span><i class="bi bi-chevron-right"></i></a>\n  {% else %}<p class="empty">No hay documentos en esta categoría.</p>{% endfor %}\n  </section>\n  {{ nav(active) }}\n</div>\n{% endblock %}\n', 'perfil.html': '{% extends \'base.html\' %}\n{% from \'partials.html\' import nav %}\n{% block content %}\n<div class="app-body">\n  <div class="page-title"><b>Perfil</b></div>\n  <section class="profile">\n    <div class="avatar"><i class="bi bi-person-fill"></i></div>\n    <h2>{{ current_user.nombres }}</h2>\n    <p>DNI {{ current_user.dni }} · {{ current_user.empresa }}</p>\n    <div class="card info"><p><b>Área:</b> {{ current_user.area }}</p><p><b>Cargo:</b> {{ current_user.cargo }}</p><p><b>Marcaciones:</b> {{ total }}</p></div>\n    <a class="btn soft" href="{{ url_for(\'firma\') }}"><i class="bi bi-pencil-fill"></i> Gestionar firma digital</a>\n    <a class="btn danger" href="{{ url_for(\'logout\') }}">Cerrar sesión</a>\n  </section>\n  {{ nav(active) }}\n</div>\n{% endblock %}\n', 'promo.html': '{% extends \'base.html\' %}\n{% block content %}\n<section class="promo-page"><i class="bi {{ icon }}"></i><h1>{{ title }}</h1><p>{{ text }}</p></section>\n{% endblock %}\n'}
CSS_APP = ':root{--teal:#0f9693;--teal2:#0b7f7d;--dark:#042b2b;--green:#25b65c;--lime:#d9ff57;--yellow:#ffc51b;--muted:#8a94a6;--line:#e6ecf0;--soft:#f5f8fb}*{box-sizing:border-box}html,body{margin:0;font-family:Inter,system-ui,Segoe UI,Arial,sans-serif;background:linear-gradient(180deg,#109693 0%,#001a1a 100%);color:#111827}a{text-decoration:none;color:inherit}.stage{min-height:100vh;display:grid;place-items:center;padding:10px}.phone{width:min(392px,100%);height:min(850px,calc(100vh - 20px));min-height:760px;background:#fff;border:9px solid #060606;border-radius:42px;box-shadow:0 25px 65px rgba(0,0,0,.45);overflow:hidden;position:relative}.phone:before{content:"";position:absolute;top:0;left:50%;transform:translateX(-50%);width:118px;height:28px;background:#060606;border-radius:0 0 16px 16px;z-index:50}.status{height:34px;padding:8px 18px 0;display:flex;align-items:center;justify-content:space-between;font-size:13px;font-weight:900;background:#fff;position:relative;z-index:51}.icons{font-size:11px;letter-spacing:1px}.screen{height:calc(100% - 34px);overflow:auto;background:#fff;position:relative}.flashbox{position:absolute;top:8px;left:16px;right:16px;z-index:60}.flash{padding:9px 11px;border-radius:10px;margin-bottom:6px;font-size:12px;font-weight:800;background:#ecfdf3;color:#067647;border:1px solid #abefc6}.flash.danger{background:#fef3f2;color:#b42318;border-color:#fecdca}.card{background:#fff;border:1px solid var(--line);border-radius:13px;box-shadow:0 8px 22px rgba(15,35,45,.14);padding:20px}.btn{height:43px;border:0;border-radius:9px;width:100%;display:flex;align-items:center;justify-content:center;gap:8px;font-weight:900;cursor:pointer;font-size:13px}.btn.teal{background:var(--teal);color:#fff}.btn.green{background:var(--green);color:#fff}.btn.soft{background:#f4f7f8;color:#4b5563}.btn.danger{background:#fff0f0;color:#ef4444}.two{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:12px 0}.input{width:100%;height:41px;border:1px solid #e5e7eb;border-radius:8px;padding:0 12px;background:#fff;color:#374151;font-weight:700}.login-screen{min-height:100%;position:relative;padding-top:34px;background:#fff}.logo{display:grid;place-items:center;margin:0 auto 16px}.mark{width:112px;height:78px;position:relative}.mark .head{position:absolute;width:22px;height:22px;border-radius:50%;top:2px}.mark .h1{left:12px;background:#26a6a2}.mark .h2{left:44px;width:27px;height:27px;background:var(--yellow)}.mark .h3{right:12px;background:#26a6a2}.mark .crown{position:absolute;left:18px;right:18px;bottom:17px;height:27px;background:var(--teal);border-radius:0 0 30px 30px}.brand{color:var(--teal);font-weight:1000;font-style:italic;font-size:22px;text-align:center;line-height:.9}.brand small{display:block;font-size:8px;font-style:normal;margin-top:4px}.login-card{margin:0 17px;padding:19px 18px}.login-card h4{text-align:center;color:#667085;font-size:12px;margin:0 0 18px}.login-card label{font-size:10px;color:#667085;font-weight:800;display:block;margin:10px 0 6px}.passbox{position:relative}.passbox button{position:absolute;right:4px;top:4px;width:35px;height:33px;border:0;background:white;color:#a0a7b1;border-radius:8px}.recover{color:#5577ff;text-align:center;font-size:10px;font-weight:800;margin-top:14px}.help{position:absolute;bottom:34px;left:0;right:0;text-align:center;font-size:9px;color:#8b949e}.help b{color:var(--teal)}.splash{height:100%;background:linear-gradient(rgba(15,150,147,.86),rgba(15,150,147,.86)),linear-gradient(135deg,#dde7e7,#536166);display:grid;place-items:center}.splash-logo .logo{transform:scale(1.16)}.splash-logo .brand,.splash-logo .brand small{color:white}.splash-logo .mark .h1,.splash-logo .mark .h3{background:#fff}.splash-logo .mark .crown{background:#fff}.app-body{min-height:100%;padding-bottom:72px;background:#fff;position:relative}.home-hero{height:150px;background:linear-gradient(135deg,var(--teal),#20a89f);color:#fff;padding:40px 22px 18px;border-radius:0 0 25px 25px;position:relative;overflow:hidden}.home-hero:after{content:"$";position:absolute;right:-2px;top:-20px;font-size:100px;font-weight:1000;opacity:.15}.hello{font-size:13px}.hello b{display:block;font-size:24px;line-height:1.05}.home-hero p{font-size:12px;opacity:.9;font-weight:700}.quick-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:-30px 18px 16px;position:relative}.feature{height:94px;background:white;border:1px solid var(--line);border-radius:12px;box-shadow:0 9px 22px rgba(15,35,45,.12);padding:12px;display:flex;flex-direction:column;gap:6px;justify-content:center}.feature.big{grid-column:1/4;background:linear-gradient(135deg,var(--teal),#41beb5);color:#fff;height:92px}.feature i{width:28px;height:28px;border:1px solid #dbeafe;border-radius:8px;display:grid;place-items:center;color:#2563eb}.feature.big i{color:#fff;border-color:rgba(255,255,255,.4)}.feature b{font-size:13px}.feature span{font-size:10px;color:#667085;font-weight:700}.feature.big span{color:#e0fbf8}.section-title{display:flex;justify-content:space-between;align-items:center;margin:12px 20px 8px;font-size:11px;color:#98a2b3;font-weight:900}.section-title a{color:#155eef}.list{padding:0 16px}.list.padded{padding-top:14px}.doc-row{min-height:64px;background:white;border:1px solid var(--line);border-radius:12px;margin:10px 0;padding:10px 12px;display:grid;grid-template-columns:42px 1fr 20px;gap:10px;align-items:center;box-shadow:0 4px 12px rgba(0,0,0,.05)}.doc-badge{width:32px;height:32px;border-radius:9px;background:#edfbd8;color:#76a815;display:grid;place-items:center;font-size:11px;font-weight:900}.doc-row b{display:block;font-size:12px}.doc-row small{display:block;color:#98a2b3;font-size:10px;margin-top:3px}.doc-row i{color:#98a2b3}.nav{height:66px;border-top:1px solid var(--line);background:white;display:grid;grid-template-columns:repeat(4,1fr);position:absolute;bottom:0;left:0;right:0}.nav a{display:grid;place-items:center;text-align:center;color:#737b86;font-size:10px;font-weight:800;padding:5px 0}.nav i{font-size:20px;display:block}.nav .active{color:var(--teal)}.assist-head{background:linear-gradient(180deg,var(--teal),#0a7977);color:white;padding:40px 22px 64px;text-align:center}.promo-icon{color:var(--lime);font-size:28px;margin-bottom:6px}.assist-head h1{font-size:27px;line-height:1.06;margin:0;font-weight:1000;letter-spacing:-.6px}.assist-head p{font-size:13px;line-height:1.45;font-weight:800;margin:16px auto 0;max-width:300px}.assist-card{margin:-48px 22px 16px;background:white;border-radius:26px;border:7px solid #060606;box-shadow:0 20px 42px rgba(0,0,0,.28);padding:20px 18px;position:relative;z-index:2}.tilted{transform:rotate(-4deg);transform-origin:center top}.assist-card>*{transform:rotate(4deg)}.clock{text-align:center;margin-bottom:16px}.clock span{display:block;font-size:9px;color:var(--teal);font-weight:1000}.clock b{display:block;font-size:44px;font-weight:1000;color:#030722;letter-spacing:-2px}.clock small{display:block;font-size:11px;color:#98a2b3}.assist-card .btn{margin-bottom:10px}.gps-line{display:flex;align-items:center;justify-content:space-between;margin-top:18px}.gps-line b{font-size:10px;color:#98a2b3}.gps-line span{font-size:9px;background:#ddfbe7;color:#17a34a;border-radius:999px;padding:5px 8px;font-weight:1000}.map{height:128px;border:1px solid #d7dde3;border-radius:12px;margin-top:8px;position:relative;overflow:hidden;background:#f5f7fb}.roads{position:absolute;inset:0;background:linear-gradient(35deg,transparent 48%,#fad77c 49% 52%,transparent 53%),linear-gradient(115deg,transparent 47%,#cbd5e1 48% 51%,transparent 52%),linear-gradient(0deg,transparent 48%,#cbd5e1 49% 51%,transparent 52%);background-size:100% 100%,100% 100%,70px 70px}.pin{position:absolute;left:50%;top:58%;transform:translate(-50%,-50%);font-size:32px;color:#2e90fa}.map em{position:absolute;left:42%;top:35%;background:white;border:1px solid #ddd;border-radius:6px;padding:5px;font-size:8px;font-style:normal}.address{font-size:9px;color:#667085;font-weight:700}.history{display:block;text-align:right;color:#155eef;font-size:10px;font-weight:900}.historial{padding:8px 24px 0}.historial b{font-size:10px;color:#98a2b3}.mini-row{display:flex;justify-content:space-between;border-bottom:1px solid var(--line);padding:7px 0;font-size:10px}.page-title{height:52px;display:flex;align-items:center;gap:14px;padding:0 20px;font-size:12px;font-weight:900}.page-title a{font-size:22px}.signature-form{padding:18px}.signature-form>label{display:block;color:#98a2b3;font-weight:900;font-size:11px;margin-bottom:10px}.canvas-wrap{height:174px;border:2px dashed #e1e8ef;border-radius:13px;position:relative;background:#fff}.canvas-wrap canvas{width:100%;height:100%;display:block;touch-action:none}.draw-pill{position:absolute;right:10px;top:10px;z-index:3;border:1px solid #e3e8ef;background:white;border-radius:9px;padding:8px 12px;color:#344054;font-weight:900}.canvas-wrap img{position:absolute;inset:12px;max-width:calc(100% - 24px);max-height:calc(100% - 24px);margin:auto;pointer-events:none}.docs-screen{background:linear-gradient(180deg,var(--teal) 0%,#001717 100%);overflow:hidden}.docs-card{background:white;border-radius:0 0 32px 32px;margin-top:-20px;padding:28px 18px 20px;transform:rotate(-4deg);transform-origin:center top;box-shadow:0 16px 36px rgba(0,0,0,.25)}.docs-card>*{transform:rotate(4deg)}.docs-head{height:58px;background:linear-gradient(135deg,var(--teal),#2cb2a8);color:white;border-radius:0 0 18px 18px;margin:-28px 20px 12px;display:flex;align-items:center;justify-content:space-between;padding:0 18px}.docs-head i{font-size:48px;opacity:.45}.menu-row{height:54px;display:grid;grid-template-columns:34px 1fr 20px;gap:10px;align-items:center}.menu-row span{width:25px;height:25px;border:1px solid #96d9dd;border-radius:7px;display:grid;place-items:center;color:var(--teal)}.menu-row b{font-size:14px;font-weight:700}.dark-caption{text-align:center;color:white;padding:38px 28px 0}.dark-caption i{font-size:32px;color:var(--lime)}.dark-caption h1{font-size:28px;line-height:1.1;margin:12px 0;font-weight:1000}.dark-caption p{font-size:13px;line-height:1.45;font-weight:800}.profile{text-align:center;padding:22px 22px}.avatar{width:92px;height:92px;border-radius:50%;background:linear-gradient(135deg,var(--teal),var(--green));color:white;display:grid;place-items:center;font-size:46px;margin:10px auto 12px}.profile h2{font-size:20px;margin:0}.profile p{font-size:12px;color:#667085}.info{text-align:left;margin:18px 0}.info p{font-size:13px;color:#344054}.empty{text-align:center;color:#98a2b3;font-size:13px;margin-top:40px}.promo-page{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(180deg,var(--teal),#001717);color:white;text-align:center;padding:40px}.promo-page i{font-size:44px;color:var(--lime)}.promo-page h1{font-size:30px;line-height:1.05;font-weight:1000}.promo-page p{font-size:14px;line-height:1.5;font-weight:800}@media(max-width:430px){.stage{padding:0}.phone{border-radius:0;border:0;width:100%;height:100vh;min-height:100vh}.phone:before{display:none}.tilted,.assist-card>*,.docs-card,.docs-card>*{transform:none}.assist-card{margin:-42px 16px 16px}.docs-card{border-radius:0 0 28px 28px}.status{display:none}.screen{height:100%}}\n'
JS_APP = "function togglePass(){const p=document.getElementById('pass'); if(!p)return; p.type=p.type==='password'?'text':'password';}\nif('serviceWorker' in navigator){navigator.serviceWorker.register('/sw.js').catch(()=>{});}\nfunction initAsistencia(){\n  const liveTime=document.getElementById('liveTime'), liveDate=document.getElementById('liveDate');\n  function tick(){const d=new Date(); if(liveTime)liveTime.textContent=d.toLocaleTimeString('es-PE',{hour:'2-digit',minute:'2-digit'}); if(liveDate)liveDate.textContent=d.toLocaleDateString('es-PE',{weekday:'long',day:'2-digit',month:'long',year:'numeric'});}\n  tick(); setInterval(tick,1000);\n  const gpsText=document.getElementById('gpsText');\n  if(navigator.geolocation){navigator.geolocation.getCurrentPosition(p=>{document.getElementById('lat').value=p.coords.latitude;document.getElementById('lng').value=p.coords.longitude;document.getElementById('acc').value=Math.round(p.coords.accuracy); if(gpsText)gpsText.textContent='Ubicación detectada con precisión aproximada de '+Math.round(p.coords.accuracy)+' m.';},()=>{if(gpsText)gpsText.textContent='No se pudo obtener ubicación. Activa GPS o permite ubicación en el navegador.'},{enableHighAccuracy:true,timeout:10000});}\n  const frm=document.getElementById('frmMarcacion'); if(frm){frm.addEventListener('submit',async e=>{e.preventDefault();const fd=new FormData(frm);fd.set('tipo',e.submitter.value);const r=await fetch('/api/marcar',{method:'POST',body:fd});const j=await r.json();alert(j.msg);location.reload();});}\n}\nlet sigCtx,sigCanvas,sigDrawing=false,sigLast=null;\nfunction initFirma(){sigCanvas=document.getElementById('signatureCanvas'); if(!sigCanvas)return; sigCtx=sigCanvas.getContext('2d'); resizeSig(); window.addEventListener('resize',resizeSig); sigCanvas.addEventListener('pointerdown',e=>{sigDrawing=true;sigLast=pos(e);const old=document.getElementById('oldSignature'); if(old)old.style.display='none';}); sigCanvas.addEventListener('pointermove',e=>{if(!sigDrawing)return;const p=pos(e);sigCtx.beginPath();sigCtx.moveTo(sigLast.x,sigLast.y);sigCtx.lineTo(p.x,p.y);sigCtx.stroke();sigLast=p;}); window.addEventListener('pointerup',()=>sigDrawing=false); const up=document.getElementById('uploadSignature'); if(up)up.onchange=e=>{const file=e.target.files[0]; if(!file)return; const img=new Image(); img.onload=()=>{clearSignature();sigCtx.drawImage(img,20,20,sigCanvas.width-40,sigCanvas.height-40);}; img.src=URL.createObjectURL(file);}; const form=document.querySelector('.signature-form'); if(form)form.addEventListener('submit',()=>{document.getElementById('firmaData').value=sigCanvas.toDataURL('image/png');});}\nfunction resizeSig(){const old=sigCanvas.toDataURL(); sigCanvas.width=sigCanvas.offsetWidth*2;sigCanvas.height=sigCanvas.offsetHeight*2;sigCtx.lineWidth=4;sigCtx.lineCap='round';sigCtx.strokeStyle='#111';}\nfunction pos(e){const r=sigCanvas.getBoundingClientRect(); return {x:(e.clientX-r.left)*2,y:(e.clientY-r.top)*2};}\nfunction clearSignature(){if(!sigCanvas)return;sigCtx.clearRect(0,0,sigCanvas.width,sigCanvas.height);const old=document.getElementById('oldSignature'); if(old)old.style.display='none';}\n"
MANIFEST_JSON = '{"name":"Nisira GCH","short_name":"NisiraGCH","start_url":"/login","display":"standalone","background_color":"#ffffff","theme_color":"#0f9693","icons":[]}\n'
SW_JS = "self.addEventListener('install',e=>self.skipWaiting());\nself.addEventListener('activate',e=>self.clients.claim());\nself.addEventListener('fetch',e=>{});\n"
app.jinja_loader = DictLoader(TEMPLATES)

def serve_embedded_static(filename):
    if filename == "css/app.css":
        return Response(CSS_APP, mimetype="text/css")
    if filename == "js/app.js":
        return Response(JS_APP, mimetype="text/javascript")
    return Response("Not found", status=404)
app.add_url_rule("/static/<path:filename>", "static", serve_embedded_static)


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
    return Response(MANIFEST_JSON, mimetype="application/manifest+json")


@app.route("/sw.js")
def sw():
    return Response(SW_JS, mimetype="text/javascript")


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
else:
    init_db()
