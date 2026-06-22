# Nisira GCH Mobile

Aplicativo móvil tipo PWA preparado para Render.

## Módulos
- Login Nisira GCH
- Home
- Asistencia con GPS
- Firma digital
- Documentos laborales
- Perfil

## Usuario demo
- Usuario: `11223344`
- Contraseña: `123456`

## Render
Build command:
```bash
pip install -r requirements.txt
```
Start command:
```bash
gunicorn app:app
```

Variables recomendadas:
- `SECRET_KEY`
- `DATABASE_URL` si usarás PostgreSQL de Render
