
# -*- coding: utf-8 -*-
import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Salida Ausentismos",
    page_icon="📋",
    layout="wide",
)

st.markdown("""
<style>
.main-title {
    font-size: 34px;
    font-weight: 800;
    color: #ef6c00;
    margin-bottom: 0;
}
.sub-title {
    font-size: 15px;
    color: #555;
    margin-top: 0;
}
.info-box {
    background-color: #fff7ed;
    border-left: 5px solid #ef6c00;
    padding: 12px 16px;
    border-radius: 8px;
    margin-bottom: 16px;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">Generador Salida_Ausentismos</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-title">Genera el Excel final con Matriz y Gestionados. Sin gráficas.</p>',
    unsafe_allow_html=True
)

st.markdown("""
<div class="info-box">
<b>Regla:</b> la base nace del Reporte_Aus_Acumulado. 
Si en SF/Hello existe un registro APPROVED con mismo ID y fecha igual o solapada, pasa a Gestionados.
Los demás registros quedan en Matriz.
</div>
""", unsafe_allow_html=True)


# ============================================================
# UTILIDADES
# ============================================================

def norm_text(value):
    value = "" if value is None else str(value)
    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def clean_id(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    match = re.search(r"\d+", text)
    return match.group(0) if match else text


def parse_date(value):
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.to_datetime(value).normalize()

    text = str(value).strip()
    if text.lower() in {"", "nan", "nat", "none"}:
        return pd.NaT

    text = re.sub(r"\s+00:00:00$", "", text)

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return pd.Timestamp(datetime.strptime(text, fmt)).normalize()
        except Exception:
            pass

    return pd.to_datetime(text, errors="coerce", dayfirst=True).normalize()


def fmt_date(value):
    if pd.isna(value):
        return ""
    value = pd.to_datetime(value)
    return f"{value.day:02d}/{value.month:02d}/{value.year}"


def month_key(value):
    if pd.isna(value):
        return ""
    value = pd.to_datetime(value)
    return f"{value.year:04d}-{value.month:02d}"


def ranges_overlap(a_ini, a_fin, b_ini, b_fin):
    if pd.isna(a_ini) or pd.isna(a_fin) or pd.isna(b_ini) or pd.isna(b_fin):
        return False
    return (a_ini <= b_fin) and (b_ini <= a_fin)


def unique_join(values):
    out = []
    seen = set()
    for value in values:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if text.lower() in {"", "nan", "none", "nat"}:
            continue
        if text not in seen:
            seen.add(text)
            out.append(text)
    return " | ".join(out)


def find_col(df, candidates, required=True, label=""):
    norm_cols = {norm_text(c): c for c in df.columns}

    for cand in candidates:
        nc = norm_text(cand)
        if nc in norm_cols:
            return norm_cols[nc]

    for cand in candidates:
        nc = norm_text(cand)
        if not nc:
            continue
        for real_norm, real_col in norm_cols.items():
            if nc in real_norm or real_norm in nc:
                return real_col

    if required:
        raise ValueError(
            f"No encontré la columna requerida: {label or candidates}. "
            f"Columnas disponibles: {list(df.columns)}"
        )
    return None


def make_key(df, id_col="id", ini_col="inicio", fin_col="fin"):
    ini = pd.to_datetime(df[ini_col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    fin = pd.to_datetime(df[fin_col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    return df[id_col].astype(str).fillna("") + "|" + ini + "|" + fin


# ============================================================
# LECTURA DE ARCHIVOS
# ============================================================

@st.cache_data(show_spinner=False)
def read_excel_cached(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, engine="openpyxl")


def detect_csv_header_line(raw_bytes, keywords):
    text = raw_bytes.decode("utf-8-sig", errors="ignore")
    lines = text.splitlines()
    norm_keywords = [norm_text(k) for k in keywords]
    for idx, line in enumerate(lines[:100]):
        nline = norm_text(line)
        if all(k in nline for k in norm_keywords):
            return idx
    return 0


@st.cache_data(show_spinner=False)
def read_csv_cached(raw_bytes, skiprows):
    last_error = None
    for encoding in ["utf-8-sig", "latin-1"]:
        for sep in [None, ";", ",", "\t", "|"]:
            try:
                df = pd.read_csv(
                    io.BytesIO(raw_bytes),
                    sep=sep,
                    engine="python",
                    skiprows=skiprows,
                    encoding=encoding,
                    dtype=str,
                    on_bad_lines="skip"
                )
                if df.shape[1] > 1:
                    df.columns = [str(c).strip() for c in df.columns]
                    return df
            except Exception as e:
                last_error = e
    raise ValueError(f"No pude leer el CSV. Último error: {last_error}")


@st.cache_data(show_spinner=False)
def read_md_txt_cached(raw_bytes):
    text = raw_bytes.decode("utf-8-sig", errors="ignore")

    # Si parece tabla SAP con |
    if "|" in text:
        lines = text.splitlines()
        rows = []
        for line in lines:
            if "|" in line and not set(line.strip()).issubset({"-", "|"}):
                parts = [p.strip() for p in line.strip().strip("|").split("|")]
                if len(parts) >= 5:
                    rows.append(parts)

        header_idx = None
        for i, row in enumerate(rows):
            nrow = " ".join(norm_text(x) for x in row)
            if "status ocupacion" in nrow and ("n pers" in nrow or "numero de personal" in nrow):
                header_idx = i
                break

        if header_idx is not None:
            header = rows[header_idx]
            width = len(header)
            data = [r for r in rows[header_idx + 1:] if len(r) == width]
            df = pd.DataFrame(data, columns=header)
            df.columns = [str(c).strip() for c in df.columns]
            return df

    # Fallback como CSV
    skip = detect_csv_header_line(raw_bytes, ["status", "personal"])
    return read_csv_cached(raw_bytes, skip)


def read_uploaded(uploaded, kind):
    raw = uploaded.getvalue()
    name = uploaded.name.lower()

    if kind == "excel":
        return read_excel_cached(raw)

    if kind == "sf":
        skip = detect_csv_header_line(raw, ["ID personal", "startDate", "endDate"])
        return read_csv_cached(raw, skip)

    if kind == "md":
        if name.endswith((".xlsx", ".xlsm")):
            return read_excel_cached(raw)
        return read_md_txt_cached(raw)

    return read_csv_cached(raw, 0)


# ============================================================
# NORMALIZADORES
# ============================================================

def normalize_acumulado(df):
    c_id = find_col(df, ["Número de personal", "Numero de personal", "Nº pers.", "No pers", "id"], label="ID acumulado")
    c_nombre = find_col(df, ["Nombre empl./cand.", "Nombre", "Nombre empleado", "Nombre completo"], required=False)
    c_tipo = find_col(df, [
        "Txt.cl.pres./ab.3", "Txt.cl.pres", "Clase absent./pres.2",
        "Clase de absentismo o presenci", "Clase absentismo", "Tipo",
        "Descripción", "Descripc.enfermedad"
    ], required=False)
    c_ini = find_col(df, ["Inicio de validez", "Inicio", "Fecha Inicio", "startDate"], label="Inicio acumulado")
    c_fin = find_col(df, ["Fin de validez", "Final", "Fin", "Fecha Final", "endDate"], label="Fin acumulado")
    c_ceco = find_col(df, ["Centro de coste5", "Centro de coste", "Ce.coste", "Ceco", "Tipo Ceco"], required=False)
    c_desc = find_col(df, ["Descripción6", "Descripción tienda", "Centro de coste descripción"], required=False)
    c_zona = find_col(df, ["Zona", "Zone"], required=False)
    c_reg = find_col(df, ["Region", "Región", "Texto división pers."], required=False)
    c_resp = find_col(df, ["Nombre de personal del superio", "Superior", "Responsable", "Jefe"], required=False)

    out = pd.DataFrame()
    out["id"] = df[c_id].apply(clean_id)
    out["Nombre"] = df[c_nombre].astype(str).str.strip() if c_nombre else ""
    out["Tipo"] = df[c_tipo].astype(str).str.strip() if c_tipo else ""
    out["inicio"] = df[c_ini].apply(parse_date)
    out["fin"] = df[c_fin].apply(parse_date)
    out["mes_ausencia"] = out["inicio"].apply(month_key)
    out["Ceco"] = df[c_ceco].astype(str).str.strip() if c_ceco else ""
    out["Descripción de tienda"] = df[c_desc].astype(str).str.strip() if c_desc else ""
    out["zona_acumulado"] = df[c_zona].astype(str).str.strip() if c_zona else ""
    out["regional_acumulado"] = df[c_reg].astype(str).str.strip() if c_reg else ""
    out["Responsable"] = df[c_resp].astype(str).str.strip() if c_resp else ""

    out = out[(out["id"] != "") & out["inicio"].notna() & out["fin"].notna()].copy()
    out["key"] = make_key(out, "id", "inicio", "fin")
    return out


def normalize_sf(df):
    c_id = find_col(df, ["ID personal", "id", "Número de personal", "Codigo Empleado"], label="ID SF")
    c_nombre = find_col(df, ["Nombre completo", "Nombre", "employeeName"], required=False)
    c_ini = find_col(df, ["startDate", "Fecha de inicio de ausentismo", "Fecha de inicio", "Inicio"], label="Inicio SF")
    c_fin = find_col(df, ["endDate", "Fecha fin", "Fin", "Fecha final"], label="Fin SF")
    c_estado = find_col(df, ["approvalStatus", "Approval Status", "Status", "Current Step Status"], required=False)
    c_status = find_col(df, ["Status", "Current Step Status", "approvalStatus"], required=False)
    c_tipo = find_col(df, ["externalName (Label)", "Tipo", "Ausencia", "Descripción General (Picklist Label)"], required=False)
    c_comment = find_col(df, ["Workflow Steps Comments", "Comentarios", "Comments", "comentario"], required=False)

    out = pd.DataFrame()
    out["id"] = df[c_id].apply(clean_id)
    out["nombre_sf"] = df[c_nombre].astype(str).str.strip() if c_nombre else ""
    out["inicio_sf"] = df[c_ini].apply(parse_date)
    out["fin_sf"] = df[c_fin].apply(parse_date)
    out["approvalStatus"] = df[c_estado].astype(str).str.strip().str.upper() if c_estado else ""
    out["Status_SF"] = df[c_status].astype(str).str.strip().str.upper() if c_status else out["approvalStatus"]
    out["tipo_sf"] = df[c_tipo].astype(str).str.strip() if c_tipo else ""
    out["comentario_sf"] = df[c_comment].astype(str).str.strip() if c_comment else ""

    out = out[(out["id"] != "") & out["inicio_sf"].notna() & out["fin_sf"].notna()].copy()
    return out


def normalize_ts(df):
    c_id = find_col(df, ["Codigo_Empleado", "Código Empleado", "id", "Número de personal"], label="ID TS")
    c_ini = find_col(df, ["Fecha_Inicio", "Fecha Inicio", "Inicio", "startDate"], label="Inicio TS")
    c_fin = find_col(df, ["Fecha_Final", "Fecha Final", "Fin", "endDate"], label="Fin TS")
    c_obs = find_col(df, ["Observaciones", "Observación", "Comentario", "Notas"], required=False)
    c_tipo = find_col(df, ["Tipo_Ausentismo", "Tipo Ausentismo", "Motivo_Ausentismo", "Motivo"], required=False)

    out = pd.DataFrame()
    out["id"] = df[c_id].apply(clean_id)
    out["inicio_ts"] = df[c_ini].apply(parse_date)
    out["fin_ts"] = df[c_fin].apply(parse_date)
    out["comentario_ts"] = df[c_obs].astype(str).str.strip() if c_obs else ""
    out["tipo_ts"] = df[c_tipo].astype(str).str.strip() if c_tipo else ""

    out = out[(out["id"] != "") & out["inicio_ts"].notna() & out["fin_ts"].notna()].copy()
    return out


def normalize_md(df):
    c_id = find_col(df, ["Nº pers.", "N° pers.", "No pers", "Nro pers", "Numero personal", "Número de personal"], label="ID MD")
    c_nombre = find_col(df, ["Número de personal", "Nombre", "Nombre empleado"], required=False)
    c_status = find_col(df, ["Status ocupación", "Estado", "Status"], required=False)
    c_div = find_col(df, ["División de personal", "Division de personal"], required=False)
    c_area = find_col(df, ["Área de nómina", "Area de nomina"], required=False)
    c_subdiv = find_col(df, ["Subdivisión de personal", "Subdivision de personal", "Zona"], required=False)
    c_region = find_col(df, ["Región (Estado federal", "Region", "Región"], required=False)
    c_ceco = find_col(df, ["Ce.coste", "Ceco", "Centro de coste"], required=False)
    c_desc_ceco = find_col(df, ["Centro de coste"], required=False)
    c_resp = find_col(df, ["Encargado para registro de tie", "Administrador para datos maest", "Responsable"], required=False)
    c_cargo = find_col(df, ["Función", "Funcion", "Posición", "Cargo"], required=False)

    out = pd.DataFrame()
    out["id"] = df[c_id].apply(clean_id)
    out["Nombre_MD"] = df[c_nombre].astype(str).str.strip() if c_nombre else ""
    out["activo"] = df[c_status].astype(str).str.strip() if c_status else "Activo"
    out["División de personal"] = df[c_div].astype(str).str.strip() if c_div else ""
    out["Área de nómina"] = df[c_area].astype(str).str.strip() if c_area else ""
    out["zona"] = df[c_subdiv].astype(str).str.strip() if c_subdiv else ""
    out["regional"] = df[c_region].astype(str).str.strip() if c_region else ""
    out["Ceco_MD"] = df[c_ceco].astype(str).str.strip() if c_ceco else ""
    out["Descripción de tienda_MD"] = df[c_desc_ceco].astype(str).str.strip() if c_desc_ceco else ""
    out["Responsable_MD"] = df[c_resp].astype(str).str.strip() if c_resp else ""
    out["Cargo_MD"] = df[c_cargo].astype(str).str.strip() if c_cargo else ""

    out = out[out["id"] != ""].copy()
    out["_prio"] = out["activo"].str.upper().str.contains("ACTIVO", na=False).astype(int)
    out = out.sort_values(["id", "_prio"], ascending=[True, False])
    out = out.drop_duplicates("id", keep="first").drop(columns=["_prio"], errors="ignore")
    return out


# ============================================================
# PROCESAMIENTO
# ============================================================

def find_ts_comment(row, ts_by_id):
    emp = row["id"]
    if emp not in ts_by_id:
        return ""
    sub = ts_by_id[emp]
    mask = sub.apply(lambda r: ranges_overlap(row["inicio"], row["fin"], r["inicio_ts"], r["fin_ts"]), axis=1)
    return unique_join(sub.loc[mask, "comentario_ts"].tolist())


def classify_sf(row, sf_by_id):
    emp = row["id"]
    result = {
        "Estado_final": "Sin registro en Hello",
        "approvalStatus_SF": "",
        "Status_SF": "",
        "comentario_sf": "",
        "tipo_sf": "",
        "inicio_sf": "",
        "fin_sf": "",
        "Acción sugerida": "Validar creación o gestión en Hello",
    }

    if emp not in sf_by_id:
        return result

    sub = sf_by_id[emp]
    mask = sub.apply(lambda r: ranges_overlap(row["inicio"], row["fin"], r["inicio_sf"], r["fin_sf"]), axis=1)
    matches = sub.loc[mask].copy()

    if matches.empty:
        return result

    norm_status = matches["approvalStatus"].astype(str).map(norm_text)

    approved = matches[norm_status.str.contains("approved|aprobado", regex=True, na=False)]
    rejected = matches[norm_status.str.contains("rejected|rechazado", regex=True, na=False)]

    if not approved.empty:
        chosen = approved.iloc[0]
        result.update({
            "Estado_final": "Gestionado en Hello",
            "approvalStatus_SF": chosen.get("approvalStatus", ""),
            "Status_SF": chosen.get("Status_SF", ""),
            "comentario_sf": unique_join(approved["comentario_sf"].tolist()),
            "tipo_sf": unique_join(approved["tipo_sf"].tolist()),
            "inicio_sf": fmt_date(chosen.get("inicio_sf", "")),
            "fin_sf": fmt_date(chosen.get("fin_sf", "")),
            "Acción sugerida": "No gestionar, ya está aprobado en Hello",
        })
        return result

    if not rejected.empty:
        chosen = rejected.iloc[0]
        result.update({
            "Estado_final": "Rechazado en Hello",
            "approvalStatus_SF": chosen.get("approvalStatus", ""),
            "Status_SF": chosen.get("Status_SF", ""),
            "comentario_sf": unique_join(rejected["comentario_sf"].tolist()),
            "tipo_sf": unique_join(rejected["tipo_sf"].tolist()),
            "inicio_sf": fmt_date(chosen.get("inicio_sf", "")),
            "fin_sf": fmt_date(chosen.get("fin_sf", "")),
            "Acción sugerida": "Revisar comentario de rechazo y validar si sigue pendiente",
        })
        return result

    chosen = matches.iloc[0]
    result.update({
        "Estado_final": "Pendiente en Hello",
        "approvalStatus_SF": unique_join(matches["approvalStatus"].tolist()),
        "Status_SF": unique_join(matches["Status_SF"].tolist()),
        "comentario_sf": unique_join(matches["comentario_sf"].tolist()),
        "tipo_sf": unique_join(matches["tipo_sf"].tolist()),
        "inicio_sf": fmt_date(chosen.get("inicio_sf", "")),
        "fin_sf": fmt_date(chosen.get("fin_sf", "")),
        "Acción sugerida": "Hacer seguimiento en Hello",
    })
    return result


def build_output(acum_raw, sf_raw, ts_raw, md_raw):
    logs = []

    acum = normalize_acumulado(acum_raw)
    sf = normalize_sf(sf_raw)
    ts = normalize_ts(ts_raw)
    md = normalize_md(md_raw)

    logs.append({"Paso": "Acumulado", "Detalle": f"{len(acum):,} registros válidos"})
    logs.append({"Paso": "SF/Hello", "Detalle": f"{len(sf):,} registros válidos"})
    logs.append({"Paso": "Timesoft", "Detalle": f"{len(ts):,} registros válidos"})
    logs.append({"Paso": "MD activos", "Detalle": f"{len(md):,} empleados únicos"})

    counts = acum.groupby("key").size().rename("veces_en_reporte").reset_index()
    acum = acum.merge(counts, on="key", how="left")
    acum["veces_en_reporte"] = acum["veces_en_reporte"].fillna(1).astype(int)

    conteo = (
        acum.groupby(["id", "Nombre", "inicio", "fin", "mes_ausencia"], dropna=False)
        .size()
        .rename("veces_en_reporte")
        .reset_index()
    )
    conteo["inicio_txt"] = conteo["inicio"].apply(fmt_date)
    conteo["fin_txt"] = conteo["fin"].apply(fmt_date)
    conteo = conteo[["id", "Nombre", "inicio_txt", "fin_txt", "mes_ausencia", "veces_en_reporte"]]

    acum = acum.merge(md, on="id", how="left")
    acum["activo"] = acum["activo"].fillna("No está en MD activos actual")

    for col, fallback in [
        ("Nombre", "Nombre_MD"),
        ("Ceco", "Ceco_MD"),
        ("Descripción de tienda", "Descripción de tienda_MD"),
        ("Responsable", "Responsable_MD"),
    ]:
        if fallback in acum.columns:
            acum[col] = acum[col].mask(acum[col].astype(str).str.lower().isin(["", "nan", "none"]), acum[fallback].fillna(""))

    ts_by_id = {k: v for k, v in ts.groupby("id")}
    acum["comentario_ts"] = acum.apply(lambda r: find_ts_comment(r, ts_by_id), axis=1)

    sf_by_id = {k: v for k, v in sf.groupby("id")}
    sf_results = acum.apply(lambda r: pd.Series(classify_sf(r, sf_by_id)), axis=1)
    acum = pd.concat([acum.reset_index(drop=True), sf_results.reset_index(drop=True)], axis=1)

    def obs(r):
        if r["Estado_final"] == "Gestionado en Hello":
            return "Ausentismo aprobado en Hello/SF"
        if r["Estado_final"] == "Rechazado en Hello":
            return "Rechazado en Hello/SF con comentario" if str(r.get("comentario_sf", "")).strip() else "Rechazado en Hello/SF sin comentario"
        if r["Estado_final"] == "Pendiente en Hello":
            return "Pendiente de gestión en Hello/SF"
        return "Sin registro en Hello/SF"

    acum["Observación"] = acum.apply(obs, axis=1)
    acum["inicio_txt"] = acum["inicio"].apply(fmt_date)
    acum["fin_txt"] = acum["fin"].apply(fmt_date)

    final_cols = [
        "id", "Nombre", "Tipo", "inicio_txt", "fin_txt", "mes_ausencia",
        "veces_en_reporte", "activo", "zona", "regional",
        "comentario_ts", "comentario_sf", "Observación", "Estado_final",
        "approvalStatus_SF", "Status_SF", "tipo_sf", "inicio_sf", "fin_sf",
        "Ceco", "Descripción de tienda", "Responsable", "Cargo_MD",
        "División de personal", "Área de nómina", "Acción sugerida"
    ]

    for col in final_cols:
        if col not in acum.columns:
            acum[col] = ""

    salida = acum[final_cols].copy()
    gestionados = salida[salida["Estado_final"] == "Gestionado en Hello"].copy()
    matriz = salida[salida["Estado_final"] != "Gestionado en Hello"].copy()

    logs.append({"Paso": "Matriz", "Detalle": f"{len(matriz):,} registros"})
    logs.append({"Paso": "Gestionados", "Detalle": f"{len(gestionados):,} registros"})

    return matriz, gestionados, conteo, pd.DataFrame(logs)


def write_excel(matriz, gestionados, conteo, log_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheets = {
            "Matriz": matriz,
            "Gestionados": gestionados,
            "Conteo_Notificaciones": conteo,
            "Log": log_df,
        }

        for sheet, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#F28C28",
            "font_color": "white",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})

        for sheet, df in sheets.items():
            ws = writer.sheets[sheet]
            ws.freeze_panes(1, 0)
            if len(df.columns) > 0:
                ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
            for i, col in enumerate(df.columns):
                ws.write(0, i, col, header_fmt)
                width = min(max(len(str(col)) + 4, 12), 42)
                if col in {"comentario_ts", "comentario_sf", "Observación", "Acción sugerida"}:
                    ws.set_column(i, i, 48, wrap_fmt)
                else:
                    ws.set_column(i, i, width)

    output.seek(0)
    return output.getvalue()


# ============================================================
# UI
# ============================================================

st.subheader("1. Carga de archivos")

c1, c2 = st.columns(2)
with c1:
    file_acum = st.file_uploader("Reporte_Aus_Acumulado.xlsx", type=["xlsx", "xlsm"])
    file_ts = st.file_uploader("REP_AUS_TS.xlsx", type=["xlsx", "xlsm"])
with c2:
    file_sf = st.file_uploader("Ausnom_total_SF.csv", type=["csv", "txt"])
    file_md = st.file_uploader("MD activos actual (.txt / .csv / .xlsx)", type=["txt", "csv", "xlsx", "xlsm"])

st.divider()
st.subheader("2. Generar salida")

if not all([file_acum, file_sf, file_ts, file_md]):
    st.info("Carga los 4 archivos para habilitar la generación.")
else:
    if st.button("Generar Salida_Ausentismos.xlsx", type="primary", use_container_width=True):
        try:
            with st.spinner("Procesando archivos..."):
                acum_raw = read_uploaded(file_acum, "excel")
                sf_raw = read_uploaded(file_sf, "sf")
                ts_raw = read_uploaded(file_ts, "excel")
                md_raw = read_uploaded(file_md, "md")

                matriz, gestionados, conteo, log_df = build_output(acum_raw, sf_raw, ts_raw, md_raw)
                excel_bytes = write_excel(matriz, gestionados, conteo, log_df)

            st.success("Archivo generado correctamente.")

            a, b, c = st.columns(3)
            a.metric("Matriz", f"{len(matriz):,}")
            b.metric("Gestionados", f"{len(gestionados):,}")
            c.metric("Conteo", f"{len(conteo):,}")

            st.download_button(
                "Descargar Salida_Ausentismos.xlsx",
                data=excel_bytes,
                file_name="Salida_Ausentismos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            with st.expander("Vista previa Matriz"):
                st.dataframe(matriz.head(100), use_container_width=True)

            with st.expander("Log"):
                st.dataframe(log_df, use_container_width=True)

        except Exception as e:
            st.error("No se pudo generar el archivo.")
            st.exception(e)

st.caption("Creado por Andrés Huérfano Dávila - Nómina JMC")
