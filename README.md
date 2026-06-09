# Generador Salida_Ausentismos V5

App de Streamlit Cloud para generar `Salida_Ausentismos.xlsx` sin gráficas.

## Corrección V5

La columna `id` de la salida se toma de `Nº pers.` / `N pers.` / SAP.

En los reportes SAP:

- `Nº pers.` = SAP / llave principal.
- `Número de personal` = nombre del empleado.

## Insumos requeridos

1. `Reporte_Aus_Acumulado.xlsx`
2. `Ausnom_total_SF.csv`
3. `REP_AUS_TS.xlsx`
4. `MD activos actual` en `.txt`, `.csv` o `.xlsx`

## Salida

- `Matriz`
- `Gestionados`
- `Conteo_Notificaciones`
- `Log`

## Runtime

Incluye `runtime.txt` con `python-3.12`.
