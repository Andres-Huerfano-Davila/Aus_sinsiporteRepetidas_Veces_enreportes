# Generador Salida_Ausentismos

App de Streamlit Cloud para generar `Salida_Ausentismos.xlsx` sin gráficas.

## Archivos del repositorio

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`

## Configuración

El archivo `.streamlit/config.toml` permite cargas hasta 600 MB:

```toml
[server]
maxUploadSize = 600
```

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