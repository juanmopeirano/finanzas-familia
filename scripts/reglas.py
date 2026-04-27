"""
reglas.py — motor de clasificación basado en 'mapping_propuesto.xlsx'.

Carga las reglas que vos editaste y clasifica cualquier Concepto del banco
en (Categoría, Subcategoría).

Estrategia:
  1. Match exacto por concepto normalizado (hoja 'Sin patrón' — más específico)
  2. Match por token contra patrones (hoja 'Patrones' — más amplio)
  3. Si nada matchea → ('Sin clasificar', '')
"""

import sys
import re
import warnings
from pathlib import Path
from functools import lru_cache

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from config import PROJECT_DIR

# silencia el warning ruidoso de openpyxl sobre data validation
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")

MAPPING_FILE = PROJECT_DIR / "mapping_propuesto.xlsx"

STOP = {"COMPRA","REDIVA","PAGO","TRANSFERENCIA","DEBITO","CREDITO","POS","WEB",
        "TARJETA","ITAU","CAJA","DE","AHORRO","CRE","DEB","CAMBIOSST","CAMBIOSOP",
        "VARIOS","VISA","LINK","ILINK","TRASPASO","DESDE","HASTA","CUENTA","A","EN","LA","EL"}


def normalizar(s):
    """Normaliza un Concepto del banco: mayúsculas, sin números/puntuación."""
    s = "" if pd.isna(s) else str(s)
    s = s.upper()
    for ch in ".*-/": s = s.replace(ch, " ")
    s = re.sub(r"\d+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tokens(s):
    """Devuelve tokens significativos (sin stopwords del banco, >2 chars)."""
    return [t for t in s.split() if t not in STOP and len(t) > 2]


# ── carga única (cache) ──────────────────────────────────────────────────────

_REGLAS = {"sin_patron": None, "patrones": None}


def cargar_reglas(mapping_file=MAPPING_FILE, force=False):
    """Carga el mapping del Excel. Devuelve (sin_patron_dict, patrones_list)."""
    if _REGLAS["sin_patron"] is not None and not force:
        return _REGLAS["sin_patron"], _REGLAS["patrones"]

    # Hoja 'Sin patrón' → exact match por concepto normalizado
    df_sp = pd.read_excel(mapping_file, sheet_name="Sin patrón")
    df_sp = df_sp[df_sp["Categoría a asignar"].notna()].copy()
    sin_patron = {}
    for _, r in df_sp.iterrows():
        cn = str(r["Concepto normalizado"]).strip().upper()
        cat = str(r["Categoría a asignar"]).strip()
        sub = str(r["Subcategoría"]).strip() if pd.notna(r["Subcategoría"]) else ""
        if sub.lower() in ("nan", "none"):
            sub = ""
        sin_patron[cn] = (cat, sub)

    # Hoja 'Patrones' → token match (ordenados por specificidad: más matches primero)
    df_p = pd.read_excel(mapping_file, sheet_name="Patrones")
    df_p = df_p[df_p["Categoría"].notna()].copy()
    df_p = df_p.sort_values("# matches", ascending=False)
    patrones = []
    for _, r in df_p.iterrows():
        tok = str(r["Patrón (token banco)"]).strip().upper()
        cat = str(r["Categoría"]).strip()
        sub = str(r["Subcategoría"]).strip() if pd.notna(r["Subcategoría"]) else ""
        if sub.lower() in ("nan", "none"):
            sub = ""
        patrones.append((tok, cat, sub))

    _REGLAS["sin_patron"] = sin_patron
    _REGLAS["patrones"]   = patrones
    return sin_patron, patrones


# ── clasificación ────────────────────────────────────────────────────────────

@lru_cache(maxsize=10000)
def clasificar(concepto):
    """
    Devuelve (categoria, subcategoria) para un Concepto del banco.
    Si nada matchea, devuelve ('Sin clasificar', '').
    """
    sin_patron, patrones = cargar_reglas()
    cn = normalizar(concepto)
    if not cn:
        return ("Sin clasificar", "")

    # 1) Match exacto por concepto normalizado
    if cn in sin_patron:
        return sin_patron[cn]

    # 2) Match por token (primer patrón que aparezca como token)
    toks = set(tokens(cn))
    for tok, cat, sub in patrones:
        if tok in toks:
            return (cat, sub)

    return ("Sin clasificar", "")


# ── test de cobertura sobre el histórico ─────────────────────────────────────

def test_cobertura():
    from config import EXCEL_PATH, SHEET_CA
    import shutil, tempfile, os

    print("Leyendo Excel maestro...")
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_CA, engine="openpyxl")
    except PermissionError:
        tmp = os.path.join(tempfile.gettempdir(), "planilla_temp.xlsx")
        shutil.copy2(EXCEL_PATH, tmp)
        df = pd.read_excel(tmp, sheet_name=SHEET_CA, engine="openpyxl")

    df["Concepto"] = df["Concepto"].fillna("").astype(str)
    print(f"  {len(df)} movimientos totales\n")

    print("Clasificando...")
    cargar_reglas()
    df[["cat_pred", "sub_pred"]] = df["Concepto"].apply(lambda c: pd.Series(clasificar(c)))

    sin_clas = (df["cat_pred"] == "Sin clasificar").sum()
    print(f"\n=== Cobertura ===")
    print(f"Total movimientos:           {len(df)}")
    print(f"Auto-clasificados:           {len(df) - sin_clas} ({100*(len(df)-sin_clas)/len(df):.1f}%)")
    print(f"Sin clasificar (manual):     {sin_clas} ({100*sin_clas/len(df):.1f}%)\n")

    print("Distribución de categorías:")
    print(df["cat_pred"].value_counts().to_string())

    if sin_clas > 0:
        print(f"\nEjemplos de 'Sin clasificar' (top 10 por frecuencia):")
        sin_df = df[df["cat_pred"] == "Sin clasificar"]
        sin_df["cn"] = sin_df["Concepto"].apply(normalizar)
        print(sin_df["cn"].value_counts().head(10).to_string())


if __name__ == "__main__":
    test_cobertura()
