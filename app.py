
# -*- coding: utf-8 -*-
import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Cantidad de veces repetidas",
    page_icon="🦜",
    layout="wide",
)

st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<p class="main-title">🦜 Cantidad de veces repetidas Ausencias sin soporte y rechazadas por docs.</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p class="sub-title">Genera el Excel final con Matriz y Gestionados para seguimiento de ausencias sin soporte y rechazos por documentación.</p>',
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="info-box">
<b>Objetivo:</b> identificar ausencias sin soporte y rechazadas por documentos que se repiten en los reportes.
<br><br>
<b>Llave principal:</b> SAP / Nº pers. + inicio + fin de la ausencia.
<br>
<b>Validador de gestión:</b> IT2001 actual. Si la ausencia ya no existe en IT2001 actual, pasa a Gestionados.
</div>
""",
    unsafe_allow_html=True,
)

with st.expander("Criterios usados por la herramienta", expanded=False):
    st.markdown(
        """
- La base principal es el **Reporte_Aus_Acumulado.xlsx**.
- El **SAP / Nº pers.** es la llave principal del proceso.
- El **IT2001 actual** define si el ausentismo sigue vivo o ya fue gestionado/eliminado/corregido.
- Si el ausentismo aparece en IT2001 actual con el mismo SAP y fechas iguales o solapadas, queda en **Matriz**.
- Si el ausentismo ya no aparece en IT2001 actual, pasa a **Gestionados**.
- SF/Hello se usa para traer estado y comentarios, pero no define por sí solo si está gestionado.
- Timesoft se cruza por SAP y fechas solapadas para traer observaciones.
- El MD se usa como foto de **activos a hoy**, no como validación histórica.
- `veces_en_reporte` se calcula considerando el conteo del acumulado y los meses entre la fecha de modificación y el último reporte detectado.
        """
    )


# ============================================================
# UTILIDADES
# ============================================================

MONTHS_ES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}


def norm_text(value):
    value = "" if value is None else str(value)
    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def dedupe_columns(df):
    seen = {}
    new_cols = []
    for c in df.columns:
        base = str(c).strip()
        if base in seen:
            seen[base] += 1
            new_cols.append(f"{base}.{seen[base]}")
        else:
            seen[base] = 0
            new_cols.append(base)
    df.columns = new_cols
    return df


def clean_id_series(s):
    return (
        s.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.extract(r"(\d+)", expand=False)
        .fillna(s.astype(str).str.strip())
    )


def parse_date_series(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True).dt.normalize()


def fmt_date(value):
    if pd.isna(value):
        return ""
    value = pd.to_datetime(value)
    return f"{value.day:02d}/{value.month:02d}/{value.year}"


def month_key_series(s):
    return pd.to_datetime(s, errors="coerce").dt.strftime("%Y-%m").fillna("")


def make_key(df, id_col="id", ini_col="inicio", fin_col="fin"):
    ini = pd.to_datetime(df[ini_col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    fin = pd.to_datetime(df[fin_col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    return df[id_col].astype(str).fillna("") + "|" + ini + "|" + fin


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


def is_blank_series(s):
    return s.astype(str).str.strip().str.lower().isin(["", "nan", "none", "nat"])


def parse_report_period(value):
    """
    Detecta mes/año desde Source.Name, por ejemplo:
    01. Informe Ausentismos Enero 2026.xlsx -> 2026-01
    """
    text = norm_text(value)

    year_match = re.search(r"(20\d{2})", text)
    if not year_match:
        return pd.NaT
    year = int(year_match.group(1))

    for month_name, month_num in MONTHS_ES.items():
        if month_name in text:
            return pd.Period(year=year, month=month_num, freq="M")

    # Soporta nombres tipo 2026-05 o 05-2026
    m = re.search(r"(20\d{2})\D(0?[1-9]|1[0-2])", text)
    if m:
        return pd.Period(year=int(m.group(1)), month=int(m.group(2)), freq="M")

    m = re.search(r"(0?[1-9]|1[0-2])\D(20\d{2})", text)
    if m:
        return pd.Period(year=int(m.group(2)), month=int(m.group(1)), freq="M")

    return pd.NaT


def period_from_date(value):
    if pd.isna(value):
        return pd.NaT
    value = pd.to_datetime(value)
    return pd.Period(year=value.year, month=value.month, freq="M")


def months_diff_inclusive(start_period, end_period):
    if pd.isna(start_period) or pd.isna(end_period):
        return 0
    try:
        return max(0, (end_period.year - start_period.year) * 12 + (end_period.month - start_period.month) + 1)
    except Exception:
        return 0


def derive_regional_from_division(value):
    text = str(value)
    m = re.search(r"(?:Regi[oó]n|Region|Regin)\s*(\d+)", text, flags=re.IGNORECASE)
    if not m:
        return ""
    return "R" + str(m.group(1)).zfill(2)


def derive_zona_from_division(value):
    text = str(value)
    m = re.search(r"ZN\s*0?(\d+)", text, flags=re.IGNORECASE)
    if not m:
        return ""
    code = "ZN" + str(m.group(1)).zfill(2)
    zone_names = {
        "ZN01": "Zona Occidente ZN01",
        "ZN02": "Zona Norte ZN02",
        "ZN03": "Zona Centro ZN03",
    }
    return zone_names.get(code, code)


# ============================================================
# LECTURA DE ARCHIVOS
# ============================================================

@st.cache_data(show_spinner=False)
def read_excel_cached(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, engine="openpyxl")
    return dedupe_columns(df)


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
        for sep in [",", ";", "\t", "|", None]:
            try:
                df = pd.read_csv(
                    io.BytesIO(raw_bytes),
                    sep=sep,
                    engine="python",
                    skiprows=skiprows,
                    encoding=encoding,
                    dtype=str,
                    on_bad_lines="skip",
                )
                if df.shape[1] > 1:
                    df.columns = [str(c).strip() for c in df.columns]
                    return dedupe_columns(df)
            except Exception as e:
                last_error = e
    raise ValueError(f"No pude leer el CSV. Último error: {last_error}")


@st.cache_data(show_spinner=False)
def read_md_txt_cached(raw_bytes):
    text = raw_bytes.decode("utf-8-sig", errors="ignore")

    if "|" in text:
        rows = []
        for line in text.splitlines():
            if "|" in line and not set(line.strip()).issubset({"-", "|"}):
                parts = [p.strip() for p in line.strip().strip("|").split("|")]
                if len(parts) >= 5:
                    rows.append(parts)

        header_idx = None
        for i, row in enumerate(rows[:80]):
            nrow = [norm_text(x) for x in row]
            if any("pers" in x for x in nrow[:2]) and any("status" in x for x in nrow[:6]):
                header_idx = i
                break

        if header_idx is None:
            header_idx = 0

        header = rows[header_idx]
        width = len(header)
        data = [r for r in rows[header_idx + 1:] if len(r) == width]
        df = pd.DataFrame(data, columns=header)
        return dedupe_columns(df)

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
    c_source = find_col(df, ["Source.Name", "Source Name", "Archivo", "Nombre archivo"], required=False)
    c_id = find_col(
        df,
        ["Nº pers.", "N° pers.", "N pers.", "No pers", "Nro pers", "Núm. Personal", "Num Personal"],
        label="SAP / Nº pers. acumulado",
    )
    c_nombre = find_col(
        df,
        ["Número de personal", "Nmero de personal", "Nombre empl./cand.", "Nombre empleado", "Nombre completo", "Nombre"],
        required=False,
    )
    c_tipo = find_col(
        df,
        [
            "Clase de absentismo o presenci",
            "CLASE",
            "Txt.cl.pres./ab.3",
            "Txt.cl.pres",
            "Tipo",
            "Clase absent./pres.2",
        ],
        required=False,
    )
    c_ini = find_col(df, ["Inicio", "Inicio de validez", "Fecha Inicio", "startDate"], label="Inicio acumulado")
    c_fin = find_col(df, ["Final", "Fin de validez", "Fin", "Fecha Final", "endDate"], label="Fin acumulado")
    c_mod = find_col(df, ["Modif/el", "Modif el", "Modif.el", "Mod.", "Modificado el"], required=False)
    c_ceco = find_col(df, ["Centro de coste", "Ce.coste", "Ceco", "Tipo Ceco"], required=False)
    c_div = find_col(df, ["División de personal", "Divisin de personal", "Division de personal"], required=False)

    out = pd.DataFrame()
    out["reporte_origen"] = df[c_source].astype(str).str.strip() if c_source else ""
    out["id"] = clean_id_series(df[c_id])
    out["Nombre"] = df[c_nombre].astype(str).str.strip() if c_nombre else ""
    out["Tipo"] = df[c_tipo].astype(str).str.strip() if c_tipo else ""
    out["inicio"] = parse_date_series(df[c_ini])
    out["fin"] = parse_date_series(df[c_fin])
    out["fecha_modificacion"] = parse_date_series(df[c_mod]) if c_mod else pd.NaT
    out["Ceco"] = df[c_ceco].astype(str).str.strip() if c_ceco else ""
    out["División acumulado"] = df[c_div].astype(str).str.strip() if c_div else ""

    out = out[(out["id"] != "") & out["inicio"].notna() & out["fin"].notna()].copy()
    out["mes_ausencia"] = month_key_series(out["inicio"])
    out["mes_reporte_period"] = out["reporte_origen"].apply(parse_report_period)
    out["mes_modificacion_period"] = out["fecha_modificacion"].apply(period_from_date)
    out["key"] = make_key(out, "id", "inicio", "fin")
    return out


def normalize_it2001(df):
    c_id = find_col(
        df,
        ["Nº pers.", "N° pers.", "N pers.", "No pers", "Nro pers", "Número de personal"],
        label="SAP / Nº pers. IT2001",
    )
    c_nombre = find_col(df, ["Nom.empl./cand.", "Nombre", "Nombre empleado", "Número de personal"], required=False)
    c_tipo = find_col(df, ["Txt.cl.pres./ab. _2", "Txt.cl.pres./ab.2", "Txt.cl.pres", "Descripción"], required=False)
    c_ini = find_col(df, ["Válido de", "Valido de", "Inicio de validez", "Inicio"], label="Inicio IT2001")
    c_fin = find_col(df, ["Válido a", "Valido a", "Fin de validez", "Fin", "Final"], label="Fin IT2001")
    c_ceco = find_col(df, ["Ce.coste", "Centro de coste", "Ceco"], required=False)

    out = pd.DataFrame()
    out["id"] = clean_id_series(df[c_id])
    out["nombre_it2001"] = df[c_nombre].astype(str).str.strip() if c_nombre else ""
    out["tipo_it2001"] = df[c_tipo].astype(str).str.strip() if c_tipo else ""
    out["inicio_it2001"] = parse_date_series(df[c_ini])
    out["fin_it2001"] = parse_date_series(df[c_fin])
    out["ceco_it2001"] = df[c_ceco].astype(str).str.strip() if c_ceco else ""

    out = out[(out["id"] != "") & out["inicio_it2001"].notna() & out["fin_it2001"].notna()].copy()
    return out


def normalize_sf(df):
    c_id = find_col(df, ["ID personal", "id", "Codigo Empleado", "Código Empleado"], label="ID SF")
    c_ini = find_col(df, ["startDate", "Fecha de inicio de ausentismo", "Fecha de inicio", "Inicio"], label="Inicio SF")
    c_fin = find_col(df, ["endDate", "Fecha fin", "Fin", "Fecha final"], label="Fin SF")
    c_estado = find_col(df, ["approvalStatus", "Approval Status"], required=False)
    c_status = find_col(df, ["Status", "Current Step Status", "approvalStatus"], required=False)
    c_tipo = find_col(df, ["externalName (Label)", "Tipo", "Ausencia", "Descripción General (Picklist Label)"], required=False)
    c_comment = find_col(df, ["Workflow Steps Comments", "Comentarios", "Comments", "comentario"], required=False)

    out = pd.DataFrame()
    out["id"] = clean_id_series(df[c_id])
    out["inicio_sf"] = parse_date_series(df[c_ini])
    out["fin_sf"] = parse_date_series(df[c_fin])
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

    out = pd.DataFrame()
    out["id"] = clean_id_series(df[c_id])
    out["inicio_ts"] = parse_date_series(df[c_ini])
    out["fin_ts"] = parse_date_series(df[c_fin])
    out["comentario_ts"] = df[c_obs].astype(str).str.strip() if c_obs else ""

    out = out[(out["id"] != "") & out["inicio_ts"].notna() & out["fin_ts"].notna()].copy()
    return out


def normalize_md(df):
    c_id = find_col(df, ["Nº pers.", "N° pers.", "N pers.", "No pers", "Nro pers"], label="SAP / Nº pers. MD")
    c_nombre = find_col(df, ["Número de personal", "Nmero de personal", "Nombre empleado", "Nombre"], required=False)
    c_status = find_col(df, ["Status ocupación", "Status ocupacin", "Status ocupacion", "Estado", "Status"], required=False)
    c_div = find_col(df, ["División de personal", "Divisin de personal", "Division de personal"], required=False)
    c_area = find_col(df, ["Área de nómina", "rea de nmina", "Area de nomina"], required=False)
    c_ceco = find_col(df, ["Ce.coste", "Ceco"], required=False)
    c_desc_ceco = find_col(df, ["Centro de coste"], required=False)
    c_resp = find_col(df, ["Encargado para registro de tie", "Administrador para datos maest", "Responsable"], required=False)

    func_cols = [c for c in df.columns if "func" in norm_text(c)]
    c_cargo = func_cols[-1] if func_cols else None

    out = pd.DataFrame()
    out["id"] = clean_id_series(df[c_id])
    out["Nombre_MD"] = df[c_nombre].astype(str).str.strip() if c_nombre else ""
    out["activo"] = df[c_status].astype(str).str.strip() if c_status else "Activo"
    out["División de personal"] = df[c_div].astype(str).str.strip() if c_div else ""
    out["Área de nómina"] = df[c_area].astype(str).str.strip() if c_area else ""
    out["Ceco_MD"] = df[c_ceco].astype(str).str.strip() if c_ceco else ""
    out["Descripción de tienda"] = df[c_desc_ceco].astype(str).str.strip() if c_desc_ceco else ""
    out["Responsable"] = df[c_resp].astype(str).str.strip() if c_resp else ""
    out["Cargo_MD"] = df[c_cargo].astype(str).str.strip() if c_cargo else ""

    out["regional"] = out["División de personal"].apply(derive_regional_from_division)
    out["zona"] = out["División de personal"].apply(derive_zona_from_division)

    out = out[out["id"] != ""].copy()
    out["_prio"] = out["activo"].astype(str).str.upper().str.contains("ACTIVO", na=False).astype(int)
    out = out.sort_values(["id", "_prio"], ascending=[True, False])
    out = out.drop_duplicates("id", keep="first").drop(columns=["_prio"], errors="ignore")
    return out


# ============================================================
# CRUCES
# ============================================================

def classify_it2001(row, it_by_id):
    result = {
        "estado_it2001": "No existe en IT2001 actual",
        "tipo_it2001": "",
        "inicio_it2001": "",
        "fin_it2001": "",
    }

    emp = row["id"]
    if emp not in it_by_id:
        return result

    sub = it_by_id[emp]
    matches = sub[(sub["inicio_it2001"] <= row["fin"]) & (sub["fin_it2001"] >= row["inicio"])]

    if matches.empty:
        return result

    chosen = matches.iloc[0]
    result.update({
        "estado_it2001": "Sigue en IT2001 actual",
        "tipo_it2001": unique_join(matches["tipo_it2001"].tolist()),
        "inicio_it2001": fmt_date(chosen.get("inicio_it2001", "")),
        "fin_it2001": fmt_date(chosen.get("fin_it2001", "")),
    })
    return result


def classify_sf(row, sf_by_id):
    result = {
        "estado_hello": "Sin registro en Hello",
        "approvalStatus_SF": "",
        "Status_SF": "",
        "comentario_sf": "",
        "tipo_sf": "",
        "inicio_sf": "",
        "fin_sf": "",
    }

    emp = row["id"]
    if emp not in sf_by_id:
        return result

    sub = sf_by_id[emp]
    matches = sub[(sub["inicio_sf"] <= row["fin"]) & (sub["fin_sf"] >= row["inicio"])].copy()

    if matches.empty:
        return result

    status = matches["approvalStatus"].astype(str).str.upper()
    approved = matches[status.str.contains("APPROVED|APROBADO", regex=True, na=False)]
    rejected = matches[status.str.contains("REJECTED|RECHAZADO", regex=True, na=False)]

    if not approved.empty:
        chosen = approved.iloc[0]
        result.update({
            "estado_hello": "Aprobado en Hello",
            "approvalStatus_SF": chosen.get("approvalStatus", ""),
            "Status_SF": chosen.get("Status_SF", ""),
            "comentario_sf": unique_join(approved["comentario_sf"].tolist()),
            "tipo_sf": unique_join(approved["tipo_sf"].tolist()),
            "inicio_sf": fmt_date(chosen.get("inicio_sf", "")),
            "fin_sf": fmt_date(chosen.get("fin_sf", "")),
        })
        return result

    if not rejected.empty:
        chosen = rejected.iloc[0]
        result.update({
            "estado_hello": "Rechazado en Hello",
            "approvalStatus_SF": chosen.get("approvalStatus", ""),
            "Status_SF": chosen.get("Status_SF", ""),
            "comentario_sf": unique_join(rejected["comentario_sf"].tolist()),
            "tipo_sf": unique_join(rejected["tipo_sf"].tolist()),
            "inicio_sf": fmt_date(chosen.get("inicio_sf", "")),
            "fin_sf": fmt_date(chosen.get("fin_sf", "")),
        })
        return result

    chosen = matches.iloc[0]
    result.update({
        "estado_hello": "Pendiente en Hello",
        "approvalStatus_SF": unique_join(matches["approvalStatus"].tolist()),
        "Status_SF": unique_join(matches["Status_SF"].tolist()),
        "comentario_sf": unique_join(matches["comentario_sf"].tolist()),
        "tipo_sf": unique_join(matches["tipo_sf"].tolist()),
        "inicio_sf": fmt_date(chosen.get("inicio_sf", "")),
        "fin_sf": fmt_date(chosen.get("fin_sf", "")),
    })
    return result


def find_ts_comment(row, ts_by_id):
    emp = row["id"]
    if emp not in ts_by_id:
        return ""
    sub = ts_by_id[emp]
    matches = sub[(sub["inicio_ts"] <= row["fin"]) & (sub["fin_ts"] >= row["inicio"])]
    return unique_join(matches["comentario_ts"].tolist())


def build_output(acum_raw, it_raw, sf_raw, ts_raw, md_raw):
    logs = []

    acum = normalize_acumulado(acum_raw)
    it2001 = normalize_it2001(it_raw)
    sf = normalize_sf(sf_raw)
    ts = normalize_ts(ts_raw)
    md = normalize_md(md_raw)

    latest_report_period = None
    valid_report_periods = [p for p in acum["mes_reporte_period"].dropna().tolist() if not pd.isna(p)]
    if valid_report_periods:
        latest_report_period = max(valid_report_periods)
    else:
        valid_mod_periods = [p for p in acum["mes_modificacion_period"].dropna().tolist() if not pd.isna(p)]
        latest_report_period = max(valid_mod_periods) if valid_mod_periods else pd.Period(pd.Timestamp.today(), freq="M")

    latest_report_label = str(latest_report_period)

    logs.append({"Paso": "Acumulado", "Detalle": f"{len(acum):,} registros válidos"})
    logs.append({"Paso": "IT2001 actual", "Detalle": f"{len(it2001):,} registros válidos"})
    logs.append({"Paso": "SF/Hello", "Detalle": f"{len(sf):,} registros válidos"})
    logs.append({"Paso": "Timesoft", "Detalle": f"{len(ts):,} registros válidos"})
    logs.append({"Paso": "MD activos", "Detalle": f"{len(md):,} empleados únicos"})
    logs.append({"Paso": "Último reporte detectado", "Detalle": latest_report_label})

    # Conteo por registros reales del acumulado
    count_real = (
        acum.groupby("key")["mes_reporte_period"]
        .nunique(dropna=True)
        .rename("conteo_registros_acumulado")
        .reset_index()
    )
    fallback_count = acum.groupby("key").size().rename("conteo_filas_acumulado").reset_index()

    acum = acum.merge(count_real, on="key", how="left").merge(fallback_count, on="key", how="left")
    acum["conteo_registros_acumulado"] = acum["conteo_registros_acumulado"].fillna(0).astype(int)
    acum["conteo_filas_acumulado"] = acum["conteo_filas_acumulado"].fillna(1).astype(int)

    acum["meses_estimados_reporte"] = acum["mes_modificacion_period"].apply(
        lambda p: months_diff_inclusive(p, latest_report_period)
    )
    acum["veces_en_reporte"] = acum[["conteo_registros_acumulado", "conteo_filas_acumulado", "meses_estimados_reporte"]].max(axis=1).astype(int)
    acum["mes_ultimo_reporte"] = latest_report_label
    acum["fecha_modificacion_txt"] = acum["fecha_modificacion"].apply(fmt_date)
    acum["mes_modificacion"] = acum["mes_modificacion_period"].astype(str).replace("NaT", "")

    # Reducir a una línea por SAP + inicio + fin para la matriz
    # Conserva la fila más reciente por mes de reporte si existe.
    acum["_sort_period"] = acum["mes_reporte_period"].astype(str).replace("NaT", "")
    acum = acum.sort_values(["key", "_sort_period"], ascending=[True, False]).drop_duplicates("key", keep="first")

    # MD
    acum = acum.merge(md, on="id", how="left")
    acum["activo"] = acum["activo"].fillna("No está en MD activos actual")

    for col, fallback in [
        ("Nombre", "Nombre_MD"),
        ("Ceco", "Ceco_MD"),
    ]:
        if fallback in acum.columns:
            acum[col] = acum[col].mask(is_blank_series(acum[col]), acum[fallback].fillna(""))

    # IT2001 actual define Matriz vs Gestionados
    it_by_id = {k: v for k, v in it2001.groupby("id")}
    it_result = acum.apply(lambda r: pd.Series(classify_it2001(r, it_by_id)), axis=1)
    acum = pd.concat([acum.reset_index(drop=True), it_result.reset_index(drop=True)], axis=1)

    # SF/Hello contexto
    sf_by_id = {k: v for k, v in sf.groupby("id")}
    sf_result = acum.apply(lambda r: pd.Series(classify_sf(r, sf_by_id)), axis=1)
    acum = pd.concat([acum.reset_index(drop=True), sf_result.reset_index(drop=True)], axis=1)

    # Timesoft comentario
    ts_by_id = {k: v for k, v in ts.groupby("id")}
    acum["comentario_ts"] = acum.apply(lambda r: find_ts_comment(r, ts_by_id), axis=1)

    # Observación
    def make_observation(r):
        if r["estado_it2001"] == "No existe en IT2001 actual":
            if r["estado_hello"] == "Aprobado en Hello":
                return "No existe en IT2001 actual / aprobado en Hello"
            if r["estado_hello"] == "Rechazado en Hello":
                return "No existe en IT2001 actual / revisar si fue eliminado o corregido por salario"
            return "No existe en IT2001 actual / posible eliminación o corrección por salario"

        # Sigue vivo en IT2001
        if r["estado_hello"] == "Rechazado en Hello":
            return "Sigue en IT2001 y está rechazado en Hello"
        if r["estado_hello"] == "Pendiente en Hello":
            return "Sigue en IT2001 pendiente en Hello"
        if r["estado_hello"] == "Aprobado en Hello":
            return "Sigue en IT2001 y aprobado en Hello"
        return "Sigue en IT2001 sin registro en Hello"

    acum["Observación"] = acum.apply(make_observation, axis=1)
    acum["inicio_txt"] = acum["inicio"].apply(fmt_date)
    acum["fin_txt"] = acum["fin"].apply(fmt_date)

    final_cols = [
        "id", "Nombre", "Tipo", "inicio_txt", "fin_txt", "mes_ausencia",
        "fecha_modificacion_txt", "mes_modificacion", "mes_ultimo_reporte",
        "veces_en_reporte", "conteo_registros_acumulado", "meses_estimados_reporte",
        "estado_it2001", "tipo_it2001", "inicio_it2001", "fin_it2001",
        "activo", "zona", "regional",
        "estado_hello", "approvalStatus_SF", "Status_SF",
        "comentario_ts", "comentario_sf", "Observación",
        "Ceco", "Descripción de tienda", "Responsable", "Cargo_MD",
        "División de personal", "Área de nómina",
    ]

    for col in final_cols:
        if col not in acum.columns:
            acum[col] = ""

    salida = acum[final_cols].copy()

    matriz = salida[salida["estado_it2001"] == "Sigue en IT2001 actual"].copy()
    gestionados = salida[salida["estado_it2001"] != "Sigue en IT2001 actual"].copy()

    conteo = acum[[
        "id", "Nombre", "inicio_txt", "fin_txt", "fecha_modificacion_txt",
        "mes_modificacion", "mes_ultimo_reporte", "conteo_registros_acumulado",
        "meses_estimados_reporte", "veces_en_reporte", "reporte_origen"
    ]].copy()

    logs.append({"Paso": "Matriz", "Detalle": f"{len(matriz):,} registros que siguen en IT2001"})
    logs.append({"Paso": "Gestionados", "Detalle": f"{len(gestionados):,} registros que ya no están en IT2001"})
    logs.append({"Paso": "Criterio de veces", "Detalle": "max(conteo acumulado por mes, conteo filas, meses entre modificación y último reporte)"})

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
                if col in {"comentario_ts", "comentario_sf", "Observación"}:
                    ws.set_column(i, i, 48, wrap_fmt)
                else:
                    ws.set_column(i, i, width)

    output.seek(0)
    return output.getvalue()


# ============================================================
# UI
# ============================================================

st.subheader("1. Carga de insumos")

c1, c2 = st.columns(2)

with c1:
    file_acum = st.file_uploader("Reporte_Aus_Acumulado.xlsx", type=["xlsx", "xlsm"])
    file_it = st.file_uploader("IT2001 actual a hoy.xlsx", type=["xlsx", "xlsm"])
    file_ts = st.file_uploader("REP_AUS_TS.xlsx", type=["xlsx", "xlsm"])

with c2:
    file_sf = st.file_uploader("Ausnom_total_SF.csv", type=["csv", "txt"])
    file_md = st.file_uploader("MD activos actual (.txt / .csv / .xlsx)", type=["txt", "csv", "xlsx", "xlsm"])

st.divider()
st.subheader("2. Generar Excel de salida")

if not all([file_acum, file_it, file_sf, file_ts, file_md]):
    st.info("Carga los 5 archivos para habilitar la generación.")
else:
    if st.button("Generar Excel Salida_Ausentismos.xlsx", type="primary", use_container_width=True):
        try:
            with st.spinner("Procesando archivos..."):
                acum_raw = read_uploaded(file_acum, "excel")
                it_raw = read_uploaded(file_it, "excel")
                sf_raw = read_uploaded(file_sf, "sf")
                ts_raw = read_uploaded(file_ts, "excel")
                md_raw = read_uploaded(file_md, "md")

                matriz, gestionados, conteo, log_df = build_output(acum_raw, it_raw, sf_raw, ts_raw, md_raw)
                excel_bytes = write_excel(matriz, gestionados, conteo, log_df)

            st.success("Archivo generado correctamente.")

            k1, k2, k3 = st.columns(3)
            k1.metric("Matriz", f"{len(matriz):,}")
            k2.metric("Gestionados", f"{len(gestionados):,}")
            k3.metric("Conteo", f"{len(conteo):,}")

            st.download_button(
                "Descargar Salida_Ausentismos.xlsx",
                data=excel_bytes,
                file_name="Salida_Ausentismos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            with st.expander("Vista previa Matriz"):
                st.dataframe(matriz.head(100), use_container_width=True)

            with st.expander("Vista previa Gestionados"):
                st.dataframe(gestionados.head(100), use_container_width=True)

            with st.expander("Log"):
                st.dataframe(log_df, use_container_width=True)

        except Exception as e:
            st.error("No se pudo generar el archivo.")
            st.exception(e)

st.caption("🦜 Creado por Andrés Huérfano Dávila - Nómina JMC")
