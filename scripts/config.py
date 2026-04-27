from pathlib import Path

# Libro NUEVO (fuente de verdad de ahora en adelante)
EXCEL_PATH   = Path(r"C:\Users\jmpei\OneDrive\Finanzas Pili JM\Planilla Familia.xlsx")
SHEET_CA     = "CA - Caja de Ahorro"
SHEET_CC     = "CC - Cuenta Corriente"

# Libro VIEJO (solo para regenerar el libro nuevo desde cero si hace falta)
EXCEL_VIEJO       = Path(r"C:\Users\jmpei\OneDrive\Finanzas Pili JM\Planilla Madre.xlsx")
SHEET_CA_VIEJO    = "ITAU CA 9684937"
SHEET_CC_VIEJO    = "ITAU CC 1778034"

PROJECT_DIR  = Path(__file__).parent.parent
JSON_OUTPUT  = PROJECT_DIR / "data" / "finanzas.json"

CONFIANZA_MINIMA = 0.82
