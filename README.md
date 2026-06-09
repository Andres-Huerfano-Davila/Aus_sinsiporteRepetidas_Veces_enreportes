# Generador Salida_Ausentismos

Aplicación en Streamlit Cloud para generar el archivo `Salida_Ausentismos.xlsx`.

## Insumos

La app solicita 4 archivos:

1. `Reporte_Aus_Acumulado.xlsx`
2. `Ausnom_total_SF.csv`
3. `REP_AUS_TS.xlsx`
4. `MD activos actual` en formato `.txt`, `.csv` o `.xlsx`

## Salida

Genera un Excel con estas hojas:

- `Matriz`: registros que siguen pendientes por revisar o gestionar.
- `Gestionados`: registros que en SF/Hello aparecen aprobados (`APPROVED`) con fechas iguales o solapadas.
- `Conteo_Notificaciones`: conteo por empleado y rango de ausencia.
- `Log`: resumen del proceso.

## Regla principal

La base nace del `Reporte_Aus_Acumulado.xlsx`.

- Si en SF existe una ausencia `APPROVED` con mismo ID y fecha igual o solapada, pasa a `Gestionados`.
- Si en SF existe `REJECTED`, queda en `Matriz` y trae `Workflow Steps Comments`.
- Si en SF existe pendiente, queda en `Matriz` como pendiente.
- Si no existe en SF, queda en `Matriz` como `Sin registro en Hello/SF`.

El MD es foto actual de activos. No evalúa la fecha histórica de la ausencia.

## Despliegue en Streamlit Cloud

1. Crea un repositorio en GitHub.
2. Sube `app.py`, `requirements.txt` y este `README.md`.
3. Entra a Streamlit Cloud.
4. Selecciona el repositorio.
5. Main file path: `app.py`.
6. Deploy.
