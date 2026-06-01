"""
export_json.py — lee el Excel maestro, aplica las reglas y genera data/finanzas.json.

Cambios v2:
  • Usa reglas.py (basado en mapping_propuesto.xlsx editado por el usuario)
  • Filtra filas 'Saldo Inicial' / 'Saldo Final' (eran manuales, ya no se necesitan)
  • Calcula saldo inicial / final dinámicamente desde la columna 'Saldo' del banco
  • Genera estructura 'cuadro' con tabla cruzada categorías × meses (CA y CC)
"""

import sys
import json
import shutil
import tempfile
import os
import warnings
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from config import EXCEL_PATH, SHEET_CA, SHEET_CC, JSON_OUTPUT, APERTURA_CA
from reglas import clasificar, cargar_reglas

warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
warnings.filterwarnings("ignore", category=UserWarning)


# ── utilidades ────────────────────────────────────────────────────────────────

INGRESO_CATS = {"Sueldo JM", "Sueldo Pili", "Otros ingresos",
                # ingresos viejos del CC (que se preservan)
                "Sueldo", "Sueldo Pili", "Noviembre", "CI SAS", "Dayman",
                "Zureo", "G Wolman", "Uruing", "Consultax", "Otros ingresos"}
SISTEMA_CATS = {"No va", "Traspaso", "Traspasos internos"}

# Mapping categorías viejas (CA 2025) → nuevas, para traducir al vuelo en el dashboard.
# El Excel queda intacto; esto es solo para visualización consistente.
CAT_VIEJA_NUEVA = {
    "Supermercado":                 ("Supermercado",          ""),
    "Ocio, social, entrenimento":   ("Social / Amigos",       ""),
    "Extras Casa":                  ("Hogar - Mejoras",       ""),
    "No va":                        ("No va",                 ""),
    "Farmacia":                     ("Salud",                 "Farmacia"),
    "Comida oficina":               ("Comida trabajo",        ""),
    "Nafta":                        ("Estación de servicio",  "Combustible"),
    "Paquete Personal Bank":        ("Servicios",             "Personal Bank"),
    "ANTEL":                        ("Servicios",             "ANTEL"),
    "Otros":                        ("Varios",                ""),
    "Caja de Ahorro":               ("No va",                 ""),
    "Cuenta Corriente":             ("No va",                 ""),
    "Salud (Dentista, BCBS, ots)":  ("Salud",                 "Dentista"),
    "Limpieza Casa":                ("Niñera",                ""),
    "Sueldo Pili":                  ("Sueldo Pili",           ""),
    "Sueldo":                       ("Sueldo JM",             ""),
    "Honda":                        ("Cuotas",                "Honda"),
    "CJPPU":                        ("CJPPU",                 ""),
    "Celerio":                      ("Cuotas",                "Celerio"),
    "Seguro de vida":               ("Seguro de vida",        ""),
    "UTE":                          ("Servicios",             "UTE"),
    "OSE":                          ("Servicios",             "OSE"),
    "Gastos comunes":               ("Gastos comunes",        ""),
    "Jardinero":                    ("Jardín",                ""),
    "Cosmética":                    ("Cosmética",             ""),
    "CADE":                         ("Servicios",             "CADE"),
    "Traspaso a Santander Pili":    ("Traspaso",              ""),
    "Traspaso de Santander Pili":   ("Traspaso",              ""),
    "Traspaso de Itau Pili":        ("Traspaso",              ""),
    "Traspaso de CC":               ("Traspaso",              ""),
    "Ahorros":                      ("Ahorros",               ""),
}

# Fallback para CC viejo (se usa SOLO si lees del Planilla Madre original).
CC_FALLBACK_OLD_CAT = {
    "Cuenta Corriente":   ("No va",            ""),
    "Noviembre":          ("Otros ingresos",   "Noviembre"),
    "CI SAS":             ("Otros ingresos",   "CI SAS"),
    "Dayman":             ("Otros ingresos",   "Dayman"),
    "Zureo":              ("Otros ingresos",   "Zureo"),
    "G Wolman":           ("Otros ingresos",   "G Wolman"),
    "Uruing":             ("Otros ingresos",   "Uruing"),
    "Consultax":          ("Otros ingresos",   "Consultax"),
    "Otros ingresos":     ("Otros ingresos",   "Otro"),
    "BPS":                ("Impuestos",        "BPS"),
    "DGI":                ("Impuestos",        "DGI"),
    "CJPPU":              ("CJPPU",            ""),
    "Ahorros Tenere":     ("Ahorros",          "Tenere"),
    "Convivencia/ Retiro":("Social / Amigos",  "Retiro"),
    "Service Moto":       ("Varios",           "Service Moto"),
    "Extras":             ("Varios",           ""),
    "Traspasos internos": ("Traspaso",         ""),
    "CADE":               ("Servicios",        "CADE"),
}


def to_number(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float, np.number)): return float(x)
    try: return float(str(x).strip().replace(".", "").replace(",", "."))
    except: return 0.0


def safe(v):
    if v is None: return None
    if isinstance(v, float) and np.isnan(v): return None
    if isinstance(v, (np.integer,)): return int(v)
    if isinstance(v, (np.floating,)): return round(float(v), 2)
    if isinstance(v, pd.Timestamp): return str(v.date())
    return v


def mes_label(period_str):
    meses_es = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    y, m = period_str.split("-")
    return f"{meses_es[int(m)-1]} {y[2:]}"


def read_excel_safe(path, sheet):
    try:
        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except PermissionError:
        tmp = os.path.join(tempfile.gettempdir(), f"_temp_{Path(path).name}")
        shutil.copy2(path, tmp)
        return pd.read_excel(tmp, sheet_name=sheet, engine="openpyxl")


# ── carga y clasificación ─────────────────────────────────────────────────────

def cargar_hoja(excel_path, sheet_name, traducir_2025=False):
    """
    Lee la hoja del libro NUEVO (con columnas Categoría / Subcategoría).
    Si traducir_2025=True (para CA), traduce las categorías viejas de 2025 al
    esquema nuevo en memoria (el Excel queda intacto).
    """
    df = read_excel_safe(excel_path, sheet_name)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)
    for c in ["Débito", "Crédito", "Saldo"]:
        df[c] = df[c].apply(to_number) if c in df.columns else 0.0
    df["Monto"]    = df["Crédito"] - df["Débito"]
    df["Concepto"] = df["Concepto"].fillna("").astype(str).str.strip()
    df["Tipo"]     = df["Tipo"].fillna("").astype(str).str.strip() if "Tipo" in df.columns else ""
    df["Origen"]   = df["Origen"].fillna("").astype(str).str.strip() if "Origen" in df.columns else ""

    # Las cols Categoría/Subcategoría YA vienen en el libro nuevo
    df["cat"] = df["Categoría"].fillna("").astype(str).str.strip() if "Categoría" in df.columns else ""
    df["sub"] = df["Subcategoría"].fillna("").astype(str).str.strip() if "Subcategoría" in df.columns else ""

    df = df[df["Fecha"].notna() & (df["Fecha"].dt.year >= 2020)].copy()
    df["mes"] = df["Fecha"].dt.to_period("M")

    # Traducir categorías viejas (CA 2025) al esquema nuevo
    if traducir_2025:
        for idx, row in df.iterrows():
            if row["Fecha"].year < 2026 and row["cat"] in CAT_VIEJA_NUEVA:
                nc, ns = CAT_VIEJA_NUEVA[row["cat"]]
                df.at[idx, "cat"] = nc
                if ns and not row["sub"]:
                    df.at[idx, "sub"] = ns

    df.loc[df["cat"].eq(""), "cat"] = "Sin clasificar"
    return df


# ── saldos por mes ────────────────────────────────────────────────────────────

def calcular_saldos_mov(df, apertura, excluir_cats=("No va",)):
    """
    {periodo: (saldo_inicial, saldo_final)} basado en MOVIMIENTOS (no en la
    columna Saldo del banco). Incluye TODOS los orígenes (Caja de Ahorro +
    Tarjeta de crédito, porque la tarjeta también sale de la caja). Excluye las
    categorías en excluir_cats (por defecto "No va").

      saldo_final[mes] = apertura + suma acumulada de Monto (excl excluir_cats)
                         hasta el fin de ese mes.
    """
    d = df[~df["cat"].isin(list(excluir_cats))].sort_values("Fecha")
    meses = sorted(df["mes"].dropna().unique().astype(str))
    out = {}
    prev_sf = float(apertura)
    for m in meses:
        si = prev_sf
        neto = float(d[d["mes"].astype(str) == m]["Monto"].sum())
        sf = si + neto
        out[m] = (round(si, 2), round(sf, 2))
        prev_sf = sf
    return out


def calcular_saldos(df, filtrar_origen=None):
    """
    {periodo: (saldo_inicial, saldo_final)} basado en la columna Saldo del banco.
    Se usa para la Cuenta Corriente (manual). Para Caja de Ahorro se usa
    calcular_saldos_mov().
    """
    sub = df[df["Origen"] == filtrar_origen].copy() if filtrar_origen else df.copy()
    df_ord = sub.sort_values("Fecha")
    sf_por_mes = {}
    for mes, grp in df_ord.groupby("mes"):
        last = grp.sort_values("Fecha").iloc[-1]
        s = float(last["Saldo"]) if last["Saldo"] and not pd.isna(last["Saldo"]) else None
        sf_por_mes[str(mes)] = s

    meses_ord = sorted(sf_por_mes.keys())
    out = {}
    prev_sf = None
    for m in meses_ord:
        si = prev_sf
        sf = sf_por_mes[m]
        grp = df_ord[df_ord["mes"].astype(str) == m]
        movs_neto = float(grp["Monto"].sum())
        if sf is None and si is not None:
            sf = si + movs_neto
        if si is None and sf is not None:
            si = sf - movs_neto
        out[m] = (si, sf)
        prev_sf = sf
    return out


# ── agregación por mes (vista resumen) ────────────────────────────────────────

def _split_ing_gasto(df):
    """
    Devuelve (mask_ingreso, mask_gasto). Criterio:
      • Tipo == 'Ingresos' → ingreso (independiente de la categoría)
      • Tipo == 'Gastos'   → gasto
      • Tipo vacío → fallback: si cat ∈ INGRESO_CATS → ingreso; sino gasto
      • SISTEMA_CATS (No va, Traspaso) → excluido de ambos
    """
    excl  = df["cat"].isin(SISTEMA_CATS)
    tipo  = df["Tipo"].fillna("").str.strip()
    is_ing = (
        (tipo == "Ingresos") |
        ((tipo == "") & df["cat"].isin(INGRESO_CATS))
    ) & ~excl
    is_gas = ~is_ing & ~excl & (df["cat"] != "")
    return is_ing, is_gas


def _totalizar(cg, claves):
    """
    Agrupa las filas de cg por las columnas `claves` y suma el Monto.
    Devuelve lista de dicts con: f (fecha más reciente), c (concepto
    representativo), m (suma), d2 (sub), cat, origen, n (cant. de movs).
    """
    out = []
    for valores, g in cg.groupby(claves, sort=False, dropna=False):
        if not isinstance(valores, tuple):
            valores = (valores,)
        kv = dict(zip(claves, valores))
        # concepto representativo = el más frecuente del grupo
        modo = g["Concepto"].mode()
        concepto_rep = str(modo.iloc[0]) if len(modo) else str(g["Concepto"].iloc[0])
        out.append({
            "f":      str(g["Fecha"].max().date()),
            "c":      concepto_rep[:60],
            "m":      round(float(g["Monto"].sum()), 2),
            "d2":     str(kv.get("sub", "")) if kv.get("sub") else "",
            "cat":    str(kv.get("cat", "")),
            "origen": str(kv.get("Origen", "")),
            "n":      int(len(g)),
        })
    return out


def agg_meses(df):
    out = []
    is_ing_all, is_gas_all = _split_ing_gasto(df)
    df = df.assign(_is_ing=is_ing_all, _is_gas=is_gas_all)
    for mes, grp in df.groupby("mes", sort=True):
        gastos_df   = grp[grp["_is_gas"]]
        ingresos_df = grp[grp["_is_ing"]]

        total_gastos   = round(-gastos_df["Monto"].sum(), 2)
        total_ingresos = round(ingresos_df["Monto"].sum(), 2)

        cats_out = {}
        for cat, cg in gastos_df.groupby("cat"):
            # Totalizar por (Origen, Subcategoría) dentro de la categoría
            movs = _totalizar(cg, ["Origen", "sub"])
            movs.sort(key=lambda x: x["m"])  # mayor gasto (más negativo) primero
            sub = {}
            for s, sg in cg.groupby("sub"):
                if s:
                    sub[s] = round(-sg["Monto"].sum(), 2)
            cats_out[cat] = {
                "total": round(-cg["Monto"].sum(), 2),
                "sub":   sub,
                "movs":  movs,
            }

        # top movs: totalizar por (Origen, cat, sub) y tomar los 15 mayores >500
        top_all = _totalizar(gastos_df, ["Origen", "cat", "sub"])
        top = sorted([t for t in top_all if abs(t["m"]) > 500],
                     key=lambda x: x["m"])[:15]
        for t in top:
            t["desc"] = t["cat"]

        out.append({
            "id":             str(mes),
            "label":          mes_label(str(mes)),
            "gastos_total":   total_gastos,
            "ingresos_total": total_ingresos,
            "categorias":     cats_out,
            "top_movs":       top,
        })
    return out


# ── cuadro cruzado (categorías × meses) ───────────────────────────────────────

def build_cuadro(df, saldos):
    """
    Estructura:
    {
      "meses": ["2025-01", ...],
      "labels": ["Ene 25", ...],
      "saldo_inicial": [...],
      "saldo_final": [...],
      "ingresos_total": [...],
      "gastos_total": [...],
      "ingresos": [{"cat": "Sueldo JM", "valores": [...]}, ...],
      "gastos":   [{"cat": "Servicios", "valores": [...], "subcats": {"ANTEL": [...], ...}}, ...]
    }
    """
    meses = sorted(df["mes"].dropna().unique().astype(str))
    labels = [mes_label(m) for m in meses]
    n = len(meses)
    idx = {m: i for i, m in enumerate(meses)}

    df["mes_str"] = df["mes"].astype(str)

    # Saldos
    sin_arr = [saldos.get(m, (None, None))[0] for m in meses]
    sfn_arr = [saldos.get(m, (None, None))[1] for m in meses]

    # Determinar qué es ingreso vs gasto usando Tipo + cat
    is_ing, is_gas = _split_ing_gasto(df)

    # Ingresos por categoría (con subcategorías opcionales)
    ing_df = df[is_ing]
    ingresos = []
    for cat, cg in ing_df.groupby("cat"):
        valores = [0.0] * n
        for mes_str, mg in cg.groupby("mes_str"):
            valores[idx[mes_str]] = round(float(mg["Monto"].sum()), 2)
        subcats = {}
        for sub, sg in cg.groupby("sub"):
            if not sub: continue
            sv = [0.0] * n
            for mes_str, smg in sg.groupby("mes_str"):
                sv[idx[mes_str]] = round(float(smg["Monto"].sum()), 2)
            subcats[sub] = sv
        ingresos.append({"cat": cat, "valores": valores, "subcats": subcats,
                         "total": round(sum(valores), 2)})
    ingresos.sort(key=lambda x: -x["total"])
    ing_total = [round(sum(c["valores"][i] for c in ingresos), 2) for i in range(n)]

    # Gastos por categoría con subcategorías
    gasto_df = df[is_gas & (df["cat"] != "Sin clasificar")]
    gastos = []
    for cat, cg in gasto_df.groupby("cat"):
        valores = [0.0] * n
        for mes_str, mg in cg.groupby("mes_str"):
            valores[idx[mes_str]] = round(-float(mg["Monto"].sum()), 2)  # gastos como positivos
        subcats = {}
        for sub, sg in cg.groupby("sub"):
            if not sub: continue
            sub_vals = [0.0] * n
            for mes_str, smg in sg.groupby("mes_str"):
                sub_vals[idx[mes_str]] = round(-float(smg["Monto"].sum()), 2)
            subcats[sub] = sub_vals
        gastos.append({
            "cat":     cat,
            "valores": valores,
            "subcats": subcats,
            "total":   round(sum(valores), 2),
        })
    gastos.sort(key=lambda x: -x["total"])
    gas_total = [round(sum(c["valores"][i] for c in gastos), 2) for i in range(n)]

    # Sin clasificar (separado para que se vea)
    sc_df = df[df["cat"] == "Sin clasificar"]
    sc_vals = [0.0] * n
    sc_movs_por_mes = {m: [] for m in meses}
    if len(sc_df):
        for mes_str, mg in sc_df.groupby("mes_str"):
            sc_vals[idx[mes_str]] = round(-float(mg["Monto"].sum()), 2)
            for _, r in mg.iterrows():
                sc_movs_por_mes[mes_str].append({
                    "f": str(r["Fecha"].date()),
                    "c": str(r["Concepto"])[:60],
                    "m": round(float(r["Monto"]), 2),
                })

    return {
        "meses":          meses,
        "labels":         labels,
        "saldo_inicial":  sin_arr,
        "saldo_final":    sfn_arr,
        "ingresos":       ingresos,
        "ingresos_total": ing_total,
        "gastos":         gastos,
        "gastos_total":   gas_total,
        "sin_clasificar": {
            "valores":      sc_vals,
            "movs_por_mes": sc_movs_por_mes,
        },
    }


# ── promedios y serie histórica ───────────────────────────────────────────────

def calc_promedios(meses_list, n=12):
    ult = meses_list[-n:] if len(meses_list) >= n else meses_list
    tot, cnt = {}, {}
    for m in ult:
        for cat, d in m["categorias"].items():
            tot[cat] = tot.get(cat, 0) + d["total"]
            cnt[cat] = cnt.get(cat, 0) + 1
    return {c: round(tot[c] / cnt[c], 2) for c in tot}


def calc_historico(meses_list):
    labels = [m["label"] for m in meses_list]
    gastos = [m["gastos_total"] for m in meses_list]
    ing    = [m["ingresos_total"] for m in meses_list]
    cat_tot = {}
    for m in meses_list:
        for cat, d in m["categorias"].items():
            cat_tot[cat] = cat_tot.get(cat, 0) + d["total"]
    top = sorted(cat_tot, key=cat_tot.get, reverse=True)[:8]
    por_cat = {c: [m["categorias"].get(c, {}).get("total", 0) for m in meses_list] for c in top}
    return {"labels": labels, "gastos": gastos, "ingresos": ing, "por_cat": por_cat}


# ── main ──────────────────────────────────────────────────────────────────────

def export_json(excel_path=EXCEL_PATH, output=JSON_OUTPUT, verbose=True):
    if verbose: print("Cargando reglas...")
    cargar_reglas()

    if verbose: print("Procesando Caja de Ahorro...")
    ca       = cargar_hoja(excel_path, SHEET_CA, traducir_2025=True)
    # Saldo por movimientos (CA + Tarjeta), excluyendo "No va". Sin usar col Saldo.
    saldos_ca = calcular_saldos_mov(ca, APERTURA_CA, excluir_cats=("No va",))
    meses_ca = agg_meses(ca)
    cuadro_ca = build_cuadro(ca, saldos_ca)

    if verbose: print("Procesando Tarjeta de Crédito (subset CA)...")
    tc = ca[ca["Origen"].astype(str).str.contains("Tarjeta", case=False, na=False)].copy()
    cuadro_tc = build_cuadro(tc, {})  # TC no tiene saldo bancario propio

    if verbose: print("Procesando Cuenta Corriente...")
    cc       = cargar_hoja(excel_path, SHEET_CC, traducir_2025=False)
    saldos_cc = calcular_saldos(cc)
    meses_cc = agg_meses(cc)
    cuadro_cc = build_cuadro(cc, saldos_cc)

    sin_clas_total = int((ca["cat"] == "Sin clasificar").sum() + (cc["cat"] == "Sin clasificar").sum())

    payload = {
        "generado":     datetime.now().isoformat(timespec="seconds"),
        "n_revisar":    sin_clas_total,
        "meses":        meses_ca,
        "promedios":    calc_promedios(meses_ca),
        "historico":    calc_historico(meses_ca),
        "cuadro":       {"ca": cuadro_ca, "tc": cuadro_tc, "cc": cuadro_cc},
        "cc": {
            "meses": [
                {
                    "id":       str(m["id"]),
                    "label":    m["label"],
                    "ingresos": m["ingresos_total"],
                    "egresos":  m["gastos_total"],
                    "saldo":    saldos_cc.get(str(m["id"]), (None, None))[1],
                }
                for m in meses_cc
            ]
        },
    }

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, default=safe, ensure_ascii=False, indent=2),
                      encoding="utf-8")

    if verbose:
        try:
            print(f"OK JSON ({len(meses_ca)} meses CA, {len(meses_cc)} meses CC, {sin_clas_total} sin clasificar)")
        except UnicodeEncodeError:
            print(f"OK JSON ({len(meses_ca)} meses)")
    return output


if __name__ == "__main__":
    export_json()
