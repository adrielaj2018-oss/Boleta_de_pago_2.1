# P&A Mobile

Aplicativo móvil estilo P&A para Render.

Incluye en un solo `app.py`:

- Splash screen P&A.
- Login.
- Home dashboard.
- Asistencia con ubicación y reloj táctil tipo wheel picker.
- Documentos.
- Firma digital con canvas.
- Perfil.

## Usuario demo

Usuario: `11223344`  
Contraseña: `123456`

## Render

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn app:app
```
