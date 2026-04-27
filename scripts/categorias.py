# ── Árbol de categorías ─────────────────────────────────────────────────────
# Cada categoría principal puede tener subcategorías opcionales.
# Si no se usa subcategoría, la lista queda vacía.

GASTOS_CATS = {
    # Fijos
    "Supermercado":        [],
    "Servicios":           ["ANTEL", "UTE", "OSE", "Personal Bank", "CADE", "Otro"],
    "Cuotas":              ["Honda", "Celerio", "Otro"],
    "Gastos comunes":      [],
    "Limpieza / Jardín":   ["Limpieza", "Jardinero"],
    "Seguro de vida":      [],
    "CJPPU":               [],
    # Variables esenciales
    "Salud":               ["Farmacia", "Dentista", "BCBS", "Otro"],
    "Nafta":               [],
    "Comida trabajo":      [],
    # Variables discrecionales
    "Delivery / Pedidos":  [],
    "Salidas / Ocio":      [],
    "Regalos":             [],
    "Viajes":              [],
    "Hogar - Mejoras":     [],
    "Ropa":                [],
    "Cosmética":           [],
    "Deportes / Gym":      [],
    "Varios":              [],
    "Ahorros":             [],
}

INGRESOS_CATS = {"Sueldo JM", "Sueldo Pili"}

# Tipos/Descripciones que se excluyen del cálculo de gastos
EXCLUIR_DESC  = {"No va", "Traspaso", ""}
EXCLUIR_TIPO  = {"Saldo Inicial", "Saldo Final", "Traspasos"}

# ── Migración old → new ──────────────────────────────────────────────────────
# Formato: "descripción vieja": ("nueva principal", "nueva subcategoría o None")
MAPEO = {
    "Supermercado":                 ("Supermercado",      None),
    "Ocio, social, entrenimento":   ("Salidas / Ocio",    None),
    "Extras Casa":                  ("Hogar - Mejoras",   None),
    "No va":                        ("No va",             None),
    "Farmacia":                     ("Salud",             "Farmacia"),
    "Comida oficina":               ("Comida trabajo",    None),
    "Nafta":                        ("Nafta",             None),
    "Paquete Personal Bank":        ("Servicios",         "Personal Bank"),
    "ANTEL":                        ("Servicios",         "ANTEL"),
    "Otros":                        ("Varios",            None),
    "Salud (Dentista, BCBS, ots)":  ("Salud",             "Dentista"),
    "Limpieza Casa":                ("Limpieza / Jardín", "Limpieza"),
    "Sueldo Pili":                  ("Sueldo Pili",       None),
    "Sueldo":                       ("Sueldo JM",         None),
    "Honda":                        ("Cuotas",            "Honda"),
    "CJPPU":                        ("CJPPU",             None),
    "Celerio":                      ("Cuotas",            "Celerio"),
    "Seguro de vida":               ("Seguro de vida",    None),
    "UTE":                          ("Servicios",         "UTE"),
    "OSE":                          ("Servicios",         "OSE"),
    "Gastos comunes":               ("Gastos comunes",    None),
    "Jardinero":                    ("Limpieza / Jardín", "Jardinero"),
    "Cosmética":                    ("Cosmética",         None),
    "CADE":                         ("Servicios",         "CADE"),
    "Traspaso a Santander Pili":    ("Traspaso",          None),
    "Traspaso de Santander Pili":   ("Traspaso",          None),
    "Traspaso de Itau Pili":        ("Traspaso",          None),
    "Traspaso de CC":               ("Traspaso",          None),
    "Ahorros":                      ("Ahorros",           None),
    "Caja de Ahorro":               ("No va",             None),
    "Cuenta Corriente":             ("No va",             None),
}

# Stopwords para el vectorizador de texto del banco
STOP_BANCO = {
    "COMPRA","REDIVA","PAGO","TRANSFERENCIA","DEBITO","CREDITO","POS","WEB",
    "TARJETA","ITAU","CAJA","DE","AHORRO","CRE","DEB","CAMBIOSST","CAMBIOSOP",
    "VARIOS","VISA","LINK","ILINK","TRASPASO","DESDE","HASTA","CUENTA"
}
