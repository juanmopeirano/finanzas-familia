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

# Saldo de la Caja de Ahorro al inicio de los datos (1-ene-2025), ANTES del
# primer movimiento. El saldo de la app se calcula como:
#   APERTURA_CA + suma acumulada de TODOS los movimientos (Caja de Ahorro +
#   Tarjeta de crédito), excluyendo solo la categoría "No va".
# No se usa la columna Saldo del banco.
APERTURA_CA = 30307.94
