"""
actualizar.py — script principal del pipeline mensual.

Flujo:
  1. Mira la carpeta 'NUEVOS del banco/' por archivos para importar.
  2. Importa cada archivo → agrega movs nuevos al Excel + auto-clasifica.
  3. Mueve los archivos procesados a 'PROCESADOS/' con timestamp.
  4. Regenera data/finanzas.json para el dashboard.

Uso:
  python actualizar.py                # corre el ciclo completo
  python actualizar.py --solo-export  # solo regenera el JSON
"""

import sys
import shutil
import argparse
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from config import PROJECT_DIR

CARPETA_NUEVOS    = PROJECT_DIR / "NUEVOS del banco"
CARPETA_PROCESADO = PROJECT_DIR / "PROCESADOS"
EXTENSIONES_OK    = {".xlsx", ".xls", ".xlsm", ".csv"}


def _archivos_nuevos():
    if not CARPETA_NUEVOS.exists():
        return []
    return sorted([
        f for f in CARPETA_NUEVOS.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSIONES_OK
    ])


def _mover_a_procesado(archivo: Path):
    CARPETA_PROCESADO.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = CARPETA_PROCESADO / f"{ts}_{archivo.name}"
    shutil.move(str(archivo), str(destino))
    return destino


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--solo-export", action="store_true")
    args = parser.parse_args()

    total_nuevos     = 0
    total_sin_clasif = 0

    if not args.solo_export:
        archivos = _archivos_nuevos()
        if not archivos:
            print("Sin archivos nuevos en 'NUEVOS del banco/'.")
            print("(Si vas a re-publicar correcciones manuales del Excel, esta bien.)")
        else:
            from import_bank import import_bank
            for arch in archivos:
                print(f"\n--- Importando: {arch.name} ---")
                try:
                    added, sin_c = import_bank(arch)
                    total_nuevos     += added
                    total_sin_clasif += sin_c
                    if added > 0:
                        destino = _mover_a_procesado(arch)
                        print(f"  -> movido a PROCESADOS/{destino.name}")
                    else:
                        # nada nuevo: igual lo movemos para no reprocesar siempre
                        destino = _mover_a_procesado(arch)
                        print(f"  -> movido a PROCESADOS/{destino.name} (sin cambios)")
                except Exception as e:
                    print(f"  ERROR procesando {arch.name}: {e}")

    # Generar JSON
    print("\n--- Regenerando JSON del dashboard ---")
    from export_json import export_json
    export_json()

    # Resumen
    print("\n" + "=" * 50)
    if total_nuevos:
        print(f"OK: {total_nuevos} movimientos nuevos importados.")
        if total_sin_clasif:
            print(f"!!  {total_sin_clasif} requieren revision manual en el Excel (amarillo).")
        else:
            print(f"    Todos clasificados automaticamente.")
    else:
        print("OK: dashboard regenerado.")
    print("=" * 50)

    # Devolver código según haya sin clasificar
    return 1 if total_sin_clasif else 0


if __name__ == "__main__":
    sys.exit(main())
