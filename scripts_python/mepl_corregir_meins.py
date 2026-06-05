# -*- coding: utf-8 -*-
"""
mepl_corregir_meins.py
----------------------
Recorre el Sheet JUNIO 2026 buscando filas donde col A = col B
(MEIN = ORDEN), lo que indica que SAP no genero la MEIN correctamente.

Para cada fila con ese problema:
  1. Llama al VBS MEPL_BUSCAR_MEIN.vbs con la orden
  2. Obtiene el numero de MEIN real de SAP
  3. Reemplaza col A con el numero correcto

Al final valida duplicados y elimina filas repetidas avisando cuales.

Ejecutar DESPUES de supabase_sync.py.
"""

import sys
import subprocess
import re
import requests
from pathlib import Path

# ========= CONFIGURACION =========
SUPABASE_URL = "https://zynvptlesftdpyijxvpv.supabase.co"
SUPABASE_KEY = "sb_publishable_a5BtY5K2O8w2rPBlIeKRUg_u6z478MX"

SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"
HOJA_JUNIO     = "JUNIO 2026"

BASE_DIR         = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token_mepl.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

VBS_BUSCAR_MEIN = r"C:\Users\usuario\AppData\Roaming\SAP\SAP GUI\Scripts\MEPL_BUSCAR_MEIN.vbs"
# =================================


def asegurar_librerias():
    try:
        import google.auth  # noqa
    except Exception:
        print("Instalando librerias...")
        subprocess.check_call([sys.executable, "-m", "pip", "install",
            "google-auth", "google-auth-oauthlib", "google-api-python-client"])


def crear_servicio_sheets():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return build("sheets", "v4", credentials=creds)


def buscar_mein_en_sap(orden):
    """Llama al VBS para obtener el numero de MEIN real de una orden."""
    try:
        resultado = subprocess.run(
            ["cscript", "//NoLogo", VBS_BUSCAR_MEIN, orden],
            capture_output=True, text=True, timeout=30
        )
        stdout = resultado.stdout.strip()
        if "OK" in stdout:
            match = re.search(r'MEIN:\s*(\d{6,12})', stdout)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"    ERROR VBS: {e}")
    return None


def obtener_sheet_id(service):
    """Obtiene el sheetId de JUNIO 2026."""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == HOJA_JUNIO:
            return s["properties"]["sheetId"]
    return 0


def main():
    print("=" * 55)
    print("   MEPL — Corregir MEINs iguales a orden en JUNIO")
    print("=" * 55)
    print("")

    asegurar_librerias()
    service = crear_servicio_sheets()

    # ── Paso 1: Leer Sheet JUNIO ────────────────────
    print("Leyendo Sheet JUNIO 2026...")
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA_JUNIO}'!A:D"
    ).execute()
    filas = result.get("values", [])[1:]  # saltar header
    print(f"  {len(filas)} filas encontradas")

    # Detectar filas con mein = orden
    a_corregir = []
    for idx, row in enumerate(filas):
        fila_num = idx + 2
        mein  = row[0].strip() if len(row) > 0 and row[0] else ""
        orden = row[1].strip() if len(row) > 1 and row[1] else ""
        dom   = row[2].strip() if len(row) > 2 and row[2] else ""
        tipo  = row[3].strip() if len(row) > 3 and row[3] else ""

        if mein and orden and mein == orden:
            a_corregir.append({
                "fila": fila_num,
                "orden": orden,
                "dom": dom,
                "tipo": tipo
            })

    if not a_corregir:
        print("\nNo hay filas con MEIN = ORDEN. Todo correcto.")
    else:
        print(f"\n  {len(a_corregir)} fila(s) con MEIN = ORDEN:")
        for c in a_corregir:
            print(f"  - fila {c['fila']} | Orden {c['orden']} | {c['dom']} | {c['tipo']}")

        resp = input(f"\nBuscar MEINs reales en SAP para estas {len(a_corregir)} filas? (S/N): ").strip().upper()
        if resp not in ("S", "SI"):
            print("Cancelado.")
            input("\nPresione Enter para salir...")
            return

        # ── Paso 2: Buscar y corregir ───────────────────
        print("\nBuscando MEINs en SAP...")
        corregidas = 0
        sin_mein   = []

        for c in a_corregir:
            print(f"\n  fila {c['fila']} — Orden {c['orden']} | {c['dom']} | {c['tipo']}")
            mein_real = buscar_mein_en_sap(c["orden"])

            if mein_real and mein_real != c["orden"]:
                service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"'{HOJA_JUNIO}'!A{c['fila']}",
                    valueInputOption="RAW",
                    body={"values": [[mein_real]]}
                ).execute()
                print(f"    → MEIN corregida: {mein_real}")
                corregidas += 1
            else:
                print(f"    → No se encontro MEIN en SAP")
                sin_mein.append(c)

        print(f"\n  Corregidas: {corregidas} | Sin MEIN en SAP: {len(sin_mein)}")
        if sin_mein:
            print("  PENDIENTES (no encontradas en SAP):")
            for s in sin_mein:
                print(f"    fila {s['fila']} — Orden {s['orden']} | {s['dom']} | {s['tipo']}")

    # ── Paso 3: Validar y eliminar duplicados ───────
    print("\nValidando duplicados en el Sheet JUNIO...")
    result2 = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA_JUNIO}'!A:D"
    ).execute()
    filas2 = result2.get("values", [])[1:]

    vistos   = {}
    a_borrar = []

    for idx, row in enumerate(filas2):
        fila_num = idx + 2
        mein  = row[0].strip() if len(row) > 0 and row[0] else ""
        dom   = row[2].strip() if len(row) > 2 and row[2] else ""
        tipo  = row[3].strip() if len(row) > 3 and row[3] else ""
        if not mein:
            continue
        if mein in vistos:
            print(f"  [DUPLICADO] MEIN {mein} en fila {fila_num} (primera vez fila {vistos[mein]}) | {dom} | {tipo} → ELIMINANDO")
            a_borrar.append(fila_num)
        else:
            vistos[mein] = fila_num

    if not a_borrar:
        print("  Sin duplicados.")
    else:
        sheet_id = obtener_sheet_id(service)
        for fila_num in sorted(a_borrar, reverse=True):
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={"requests": [{
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": fila_num - 1,
                            "endIndex": fila_num
                        }
                    }
                }]}
            ).execute()
        print(f"  {len(a_borrar)} fila(s) duplicada(s) eliminada(s).")

    print("")
    print("=" * 55)
    print("Completado.")
    print("=" * 55)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
