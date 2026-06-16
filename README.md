# DG Approval Generator — Web App

Generador automático de plantillas DG Approval para MSC.
Desarrollado por Juan Deo.

## ¿Cómo desplegar en Streamlit Cloud? (GRATIS)

### Paso 1 — Crea una cuenta en GitHub
Ve a github.com y crea una cuenta si no tienes.

### Paso 2 — Crea un repositorio nuevo
- Click en "New repository"
- Nombre: `dg-approval-generator`
- Visibilidad: Public (necesario para Streamlit Cloud gratis)
- Click "Create repository"

### Paso 3 — Sube los archivos
Sube estos dos archivos al repositorio:
- `app.py`
- `requirements.txt`

### Paso 4 — Despliega en Streamlit Cloud
- Ve a share.streamlit.io
- Inicia sesión con tu cuenta de GitHub
- Click "New app"
- Selecciona tu repositorio `dg-approval-generator`
- Branch: `main`
- Main file path: `app.py`
- Click "Deploy"

### Paso 5 — Listo
En 2-3 minutos tendrás una URL publica como:
https://tu-usuario-dg-approval-generator.streamlit.app

Puedes compartir esa URL con cualquier persona.

## ¿Cómo usarla?

1. Abre la URL de la app
2. En el panel izquierdo edita el nombre del barco si cambio
3. Sube el archivo DCR (.xls o .xlsx)
4. Revisa el preview de bookings
5. Click "GENERATE APPROVALS"
6. Descarga el Excel con las plantillas

## Reglas aplicadas
- `*` en MP = N (no marine pollutant)
- `P` en MP = P (marine pollutant)
- Gross Weight vacio = estimado (Net Weight x 1.05)
- Port of Loading siempre fijo (configurable en sidebar)
- Sin placeholders en puntos 1 y 7
- Battery remark automatico para UN3556, UN3171, UN3480, UN3481
- Containers consolidados y deduplicados por booking
- Multiples UNs por booking generan sub-bloques separados
