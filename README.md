# 🦜 Cantidad de veces repetidas Ausencias sin soporte y rechazadas por docs. V7

App de Streamlit Cloud para generar `Salida_Ausentismos.xlsx`.

## Cambio V7

Se agrega el insumo obligatorio `IT2001 actual a hoy.xlsx`.

El IT2001 actual define la separación:

- `Matriz`: ausencias del acumulado que siguen existiendo en IT2001 actual.
- `Gestionados`: ausencias del acumulado que ya no existen en IT2001 actual.

SF/Hello queda como soporte para estados y comentarios, pero no define solo si está gestionado.

## Llave principal

```text
SAP / Nº pers. + inicio + fin de la ausencia
```

## Conteo de veces

`veces_en_reporte` se calcula como el mayor valor entre:

1. conteo de apariciones del registro en el acumulado,
2. conteo de meses distintos de reporte en el acumulado,
3. meses entre fecha de modificación y último reporte detectado.

## Insumos requeridos

1. `Reporte_Aus_Acumulado.xlsx`
2. `IT2001 actual a hoy.xlsx`
3. `Ausnom_total_SF.csv`
4. `REP_AUS_TS.xlsx`
5. `MD activos actual` en `.txt`, `.csv` o `.xlsx`

## Salida

- `Matriz`
- `Gestionados`
- `Conteo_Notificaciones`
- `Log`

## Runtime

Incluye `runtime.txt` con `python-3.12`.
