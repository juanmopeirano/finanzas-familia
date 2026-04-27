# Finanzas JM & Pili — Dashboard familiar

Tablero web (PWA) para visualizar las finanzas familiares mes a mes desde el celular o la compu.
Se alimenta del Excel `Planilla Madre.xlsx`.

## Cómo funciona

```
Descarga manual del banco (CSV/Excel)
        ↓
 actualizar.py archivo.xlsx
        ↓
 import_bank.py    (agrega filas nuevas, deduplica)
        ↓
 classify.py       (ML sugiere categorías)
        ↓
 export_json.py    (genera data/finanzas.json)
        ↓
 git push          (actualiza el dashboard online)
```

## Estructura

```
finanzas-familia/
├── index.html, styles.css, app.js   ← dashboard PWA
├── manifest.json, sw.js, icon.svg   ← assets PWA
├── data/
│   └── finanzas.json                ← datos generados
└── scripts/
    ├── config.py                    ← rutas y configuración
    ├── categorias.py                ← árbol de categorías + mapeo
    ├── import_bank.py               ← importa CSV/Excel del banco
    ├── classify.py                  ← clasificador ML
    ├── export_json.py               ← genera el JSON del dashboard
    └── actualizar.py                ← script principal (corre todo)
```

## Uso típico

**Cuando descargás el estado de cuenta del banco:**

```bash
python scripts/actualizar.py "C:\Users\jmpei\Downloads\estado_cuenta.xlsx"
```

Esto hace:
1. Lee el archivo del banco, deduplica y agrega solo lo nuevo al Excel
2. Corre el ML — clasifica automáticamente lo que pueda, marca en amarillo lo que necesita revisión
3. Regenera `data/finanzas.json`

**Si solo querés regenerar el JSON (después de corregir manualmente en el Excel):**

```bash
python scripts/actualizar.py --solo-export
```

**Solo correr el clasificador:**

```bash
python scripts/classify.py
```

## Categorías

**Fijos:** Supermercado, Servicios, Cuotas, Gastos comunes, Limpieza/Jardín, Seguro de vida, CJPPU
**Variables esenciales:** Salud, Nafta, Comida trabajo
**Variables discrecionales:** Delivery/Pedidos, Salidas/Ocio, Regalos, Viajes, Hogar-Mejoras, Ropa, Cosmética, Deportes/Gym, Varios, Ahorros

Cada categoría puede tener una **subcategoría** (Descripción 2) para detalle. Ejemplo: `Servicios → ANTEL`.

## Flujo de revisión manual

1. Después de correr `actualizar.py`, abrir `Planilla Madre.xlsx`.
2. Filtrar `ML_Revisar = TRUE` (filas amarillas) → corregir descripción si hace falta.
3. Re-correr `python scripts/actualizar.py --solo-export` para regenerar el JSON.
4. `git add data/finanzas.json && git commit -m "update" && git push`.

## Subir a GitHub Pages

```bash
cd finanzas-familia
git init
git add .
git commit -m "initial commit"
gh repo create finanzas-familia --private --source=. --push
```

Luego en GitHub: **Settings → Pages → Source: main / root**.
La URL queda algo como `https://<tu-usuario>.github.io/finanzas-familia/`.

> El repo es privado, pero **GitHub Pages sirve el contenido públicamente**. Si querés
> que solo Pili y vos lo vean, mantené el link sin compartir y considerá agregar
> autenticación con un servicio externo (Cloudflare Access tiene plan gratuito).

## Dependencias Python

```bash
pip install pandas openpyxl scikit-learn
```
