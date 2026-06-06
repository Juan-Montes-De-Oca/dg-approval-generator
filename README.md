# DG Approval Generator — Web App

Generador automático de plantillas DG Approval para MSC.
Desarrollado por Juan.


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
