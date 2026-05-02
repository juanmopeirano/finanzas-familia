"""
reglas.py — motor de clasificación basado en mapping_propuesto.xlsx (hoja 'Reglas').

Cada regla tiene 5 columnas:
  Patrón texto | Monto desde | Monto hasta | Categoría | Subcategoría

Lógica de match:
  • Si Patrón texto != vacío: chequear que esté como substring (case-insensitive)
    en el Concepto del banco.
  • Si Monto desde / hasta != vacío: chequear que el monto del movimiento
    (max(|Débito|, |Crédito|)) caiga en el rango.
  • Lógica AND: TODAS las condiciones definidas deben cumplirse.

Prioridad cuando varias reglas matchean:
  1) Las que tienen MÁS condiciones (texto Y monto > solo una)
  2) La de patrón texto más largo (más específica)
  3) Orden de la hoja (la primera gana)
"""

import sys
import os
import re
import shutil
import tempfile
import warnings
from pathlib import Path
from functools import lru_cache

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from config import PROJECT_DIR

warnings.filterwarnings("ignore", message="Data Validation extension is not supported")

MAPPING_FILE = PROJECT_DIR / "mapping_propuesto.xlsx"


# ── helpers ──────────────────────────────────────────────────────────────────

def _read_safe(path, sheet):
    try:
        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except PermissionError:
        tmp = os.path.join(tempfile.gettempdir(), f"_reglas_{Path(path).name}")
        shutil.copy2(path, tmp)
        return pd.read_excel(tmp, sheet_name=sheet, engine="openpyxl")


def _to_float(x):
    if x is None or pd.isna(x):
        return None
    try:
        return float(x)
    except Exception:
        return None


def _to_str(x):
    if x is None or pd.isna(x):
        return ""
    s = str(x).strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s


# Regex para normalizar el concepto del banco a la forma en que se generaron
# las reglas migradas (sin números, sin puntuación, sin espacios duplicados).
_RE_PUNCT  = re.compile(r"[\.\*\-/]")
_RE_DIGITS = re.compile(r"\d+")
_RE_SPACES = re.compile(r"\s+")


def _normalize_concepto(s):
    """Devuelve versión 'limpia': sin números, sin puntuación común, espacios colapsados."""
    s = _RE_PUNCT.sub(" ", s)
    s = _RE_DIGITS.sub(" ", s)
    s = _RE_SPACES.sub(" ", s).strip()
    return s


# ── carga de reglas ──────────────────────────────────────────────────────────

_REGLAS = {"data": None, "mtime": None}


def cargar_reglas(force=False):
    """Lee la hoja 'Reglas' del mapping y devuelve la lista pre-ordenada por
    especificidad descendente (las reglas más específicas primero)."""
    # Recargar si el archivo cambió (útil cuando el usuario edita y vuelve a correr)
    try:
        mtime = MAPPING_FILE.stat().st_mtime
    except FileNotFoundError:
        raise FileNotFoundError(f"No se encuentra {MAPPING_FILE}. Corré scripts/migrar_mapping.py primero.")

    if _REGLAS["data"] is not None and not force and _REGLAS["mtime"] == mtime:
        return _REGLAS["data"]

    df = _read_safe(MAPPING_FILE, "Reglas")
    reglas = []
    # Detectar si las columnas Origen/Tipo existen (compat con archivos viejos)
    has_origen_col = "Origen" in df.columns
    has_tipo_col   = "Tipo"   in df.columns

    for orig_idx, r in enumerate(df.itertuples(index=False)):
        # Columnas: Patrón texto, Monto desde, Monto hasta, Categoría, Subcategoría, [Origen, Tipo]
        patron = _to_str(r[0]).upper()
        desde  = _to_float(r[1])
        hasta  = _to_float(r[2])
        cat    = _to_str(r[3])
        sub    = _to_str(r[4])
        origen = _to_str(r[5]) if has_origen_col and len(r) > 5 else ""
        tipo   = _to_str(r[6]) if has_tipo_col   and len(r) > 6 else ""

        # Saltear filas sin Categoría
        if not cat:
            continue

        has_text = bool(patron)
        has_amt  = (desde is not None) or (hasta is not None)
        # Saltear reglas vacías (sin texto ni monto)
        if not has_text and not has_amt:
            continue

        especificidad = (1 if has_text else 0) + (1 if has_amt else 0)

        reglas.append({
            "patron": patron,
            "desde":  desde,
            "hasta":  hasta,
            "cat":    cat,
            "sub":    sub,
            "origen": origen,
            "tipo":   tipo,
            "espec":  especificidad,
            "orden":  orig_idx,
        })

    # Sort: especificidad DESC, longitud de patrón DESC, orden original ASC
    reglas.sort(key=lambda r: (-r["espec"], -len(r["patron"]), r["orden"]))

    _REGLAS["data"]  = reglas
    _REGLAS["mtime"] = mtime
    # Limpiar cache de clasificar() porque cambiaron las reglas
    clasificar.cache_clear()
    return reglas


# ── clasificación ────────────────────────────────────────────────────────────

@lru_cache(maxsize=20000)
def clasificar(concepto, debito=0.0, credito=0.0):
    """
    Clasifica un movimiento del banco.

    Args:
      concepto: descripción del banco (string)
      debito:   monto del débito (float, 0 si no hay)
      credito:  monto del crédito (float, 0 si no hay)

    Returns:
      (categoría, subcategoría, origen, tipo)
      - origen y tipo pueden venir vacíos si la regla no los especifica
      - ('Sin clasificar', '', '', '') si nada matchea
    """
    reglas = cargar_reglas()
    if not concepto:
        return ("Sin clasificar", "", "", "")

    concepto_upper = str(concepto).upper()
    # Versión normalizada (sin números, espacios colapsados) — para que reglas
    # heredadas de la migración (como "TRASPASO A ILINK") matcheen contra
    # conceptos reales como "TRASPASO A  1778034ILINK"
    concepto_norm  = _normalize_concepto(concepto_upper)

    monto_abs = max(abs(debito or 0), abs(credito or 0))

    for r in reglas:
        # Condición de texto: matchea si el patrón está en el concepto original
        # O en el concepto normalizado
        if r["patron"]:
            if r["patron"] not in concepto_upper and r["patron"] not in concepto_norm:
                continue

        # Condición de monto
        if r["desde"] is not None and monto_abs < r["desde"]:
            continue
        if r["hasta"] is not None and monto_abs > r["hasta"]:
            continue

        # Match!
        return (r["cat"], r["sub"], r["origen"], r["tipo"])

    return ("Sin clasificar", "", "", "")


# ── test sobre el histórico ──────────────────────────────────────────────────

def test_cobertura():
    """Clasifica todos los movs del Excel maestro y reporta cobertura."""
    from config import EXCEL_PATH, SHEET_CA

    print("Leyendo Excel maestro...")
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_CA, engine="openpyxl")
    except PermissionError:
        tmp = os.path.join(tempfile.gettempdir(), "planilla_temp.xlsx")
        shutil.copy2(EXCEL_PATH, tmp)
        df = pd.read_excel(tmp, sheet_name=SHEET_CA, engine="openpyxl")

    df["Concepto"] = df["Concepto"].fillna("").astype(str)
    df["Débito"]   = pd.to_numeric(df["Débito"],   errors="coerce").fillna(0)
    df["Crédito"]  = pd.to_numeric(df["Crédito"],  errors="coerce").fillna(0)
    print(f"  {len(df)} movimientos totales\n")

    print("Cargando reglas...")
    reglas = cargar_reglas(force=True)
    n_text_y_monto = sum(1 for r in reglas if r["espec"] == 2)
    n_solo_uno     = sum(1 for r in reglas if r["espec"] == 1)
    print(f"  {len(reglas)} reglas: {n_text_y_monto} con texto+monto, {n_solo_uno} con solo una condición")

    print("\nClasificando...")
    df[["cat_pred", "sub_pred", "origen_pred", "tipo_pred"]] = df.apply(
        lambda r: pd.Series(clasificar(r["Concepto"], r["Débito"], r["Crédito"])),
        axis=1,
    )

    sin_clas = (df["cat_pred"] == "Sin clasificar").sum()
    print(f"\n=== Cobertura ===")
    print(f"Total movimientos:        {len(df)}")
    print(f"Auto-clasificados:        {len(df) - sin_clas} ({100*(len(df)-sin_clas)/len(df):.1f}%)")
    print(f"Sin clasificar (manual):  {sin_clas} ({100*sin_clas/len(df):.1f}%)\n")

    print("Distribución de categorías:")
    print(df["cat_pred"].value_counts().to_string())

    if sin_clas > 0:
        print(f"\nEjemplos de 'Sin clasificar' (top 10 por frecuencia):")
        sin_df = df[df["cat_pred"] == "Sin clasificar"].copy()
        sin_df["cn"] = sin_df["Concepto"].str.upper().str.strip()
        print(sin_df["cn"].value_counts().head(10).to_string())


if __name__ == "__main__":
    test_cobertura()
