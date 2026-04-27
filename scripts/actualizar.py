"""
actualizar.py — script principal. Hace todo el ciclo:
  1. (opcional) importa nuevo archivo del banco al Excel maestro
  2. corre el clasificador ML sobre filas sin categoría
  3. exporta data/finanzas.json para el dashboard

Uso:
  python actualizar.py                     # solo clasifica + exporta JSON
  python actualizar.py archivo_banco.xlsx  # importa + clasifica + exporta
  python actualizar.py --solo-export       # solo regenera el JSON
"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from config import EXCEL_PATH


def main():
    parser = argparse.ArgumentParser(description="Pipeline finanzas: importar + clasificar + exportar")
    parser.add_argument("archivo", nargs="?", help="Ruta al CSV/Excel del banco (opcional)")
    parser.add_argument("--solo-export", action="store_true", help="Solo regenerar el JSON")
    parser.add_argument("--no-clasificar", action="store_true", help="Saltear paso ML")
    args = parser.parse_args()

    if args.solo_export:
        print("─── Solo exportar JSON ───")
        from export_json import export_json
        export_json()
        return

    # 1) Importar
    if args.archivo:
        print("─── 1. Importar banco ───")
        from import_bank import import_bank
        added = import_bank(args.archivo)
        if added == 0:
            print("(sin nuevos movimientos)")
    else:
        print("(sin archivo del banco para importar)")

    # 2) Clasificar
    if not args.no_clasificar:
        print("\n─── 2. Clasificar con ML ───")
        from classify import classify
        classify()

    # 3) Exportar JSON
    print("\n─── 3. Exportar JSON ───")
    from export_json import export_json
    export_json()

    print("\n✓ Listo. Hacé commit y push del JSON al repo de GitHub para actualizar el dashboard.")


if __name__ == "__main__":
    main()
