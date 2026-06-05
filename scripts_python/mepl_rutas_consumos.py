# -*- coding: utf-8 -*-
"""
mepl_rutas_consumos.py
----------------------
Completa automáticamente las columnas G (TITULO CERTIFICADO)
y H (RUTA CERTIFICADO) del Sheets MEPL_RESPUESTAS para las
filas de AGUA, ELECTRICIDAD e INSTALACIONES DE GAS.

- AGUA       → G: CONSUMO          / H: ruta automática
- ELECTRICA  → G: CONSUMO          / H: ruta automática
- GAS        → G: CONSUMO ||CERTIFICADO DE HERMETICIDAD
               H: ruta consumo automática || (ruta hermeticidad vacía — completar a mano)

Ejecutar una vez por mes antes de correr el launcher.
"""

import sys
import subprocess
from pathlib import Path
from datetime import datetime

# ========= CONFIGURACION =========
SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"
HOJA           = "MEPL_RESPUESTAS"

BASE_CONSUMOS  = r"C:\Users\usuario\Desktop\jorge ex\CONSUMOS"
# La ruta completa queda: BASE_CONSUMOS\{AÑO}\CONSUMOS DE JORGE\{DOMICILIO}\{TIPO}\{DOMICILIO}.xlsx

BASE_DIR         = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token_mepl.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Columnas (0-based)
COL_DOMICILIO   = 2   # C
COL_TIPO        = 3   # D
COL_RESPUESTAS  = 4   # E
COL_RESULTADO   = 5   # F
COL_TITULO_CERT = 6   # G
COL_RUTA_CERT   = 7   # H
# =================================

TIPOS_AGUA      = ["INSTALACION DE AGUA", "INSTALACION AGUA"]
TIPOS_ELECTRICA = ["INSTALACION ELECTRICA", "INSTALACION ELECTRICA", "ELECTRICA"]
TIPOS_GAS       = ["INSTALACIONES DE GAS", "INSTALACION DE GAS", "GAS"]


def asegurar_librerias():
    try:
        import google.auth  # noqa
    except Exception:
        print("Instalando librerias...")
        subprocess.check_call([sys.executable, "-m", "pip", "install",
            "google-auth", "google-auth-oauthlib", "google-api-python-client"])


def crear_servicio():
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


def leer_filas(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA}'!A:H"
    ).execute()
    return result.get("values", [])


def escribir_celda(service, fila_sheets, col_letra, valor):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA}'!{col_letra}{fila_sheets}",
        valueInputOption="RAW",
        body={"values": [[valor]]}
    ).execute()


def limpiar_domicilio(domicilio):
    """Elimina la barra / del domicilio para usar en la ruta de carpeta."""
    return domicilio.replace("/", "").replace("  ", " ").strip()


def armar_ruta(domicilio, subcarpeta, anio):
    dom_limpio = limpiar_domicilio(domicilio)
    return (
        f"{BASE_CONSUMOS}\\{anio}\\CONSUMOS DE JORGE\\"
        f"{dom_limpio}\\{subcarpeta}\\{dom_limpio}.xlsx"
    )


def detectar_tipo(tipo):
    tipo_upper = tipo.upper()
    for t in TIPOS_GAS:
        if t in tipo_upper:
            return "GAS"
    for t in TIPOS_AGUA:
        if t in tipo_upper:
            return "AGUA"
    for t in TIPOS_ELECTRICA:
        if t in tipo_upper:
            return "ELECTRICA"
    return None


def main():
    print("=" * 50)
    print("   MEPL — Autocompletar rutas de consumos")
    print("=" * 50)
    print("")

    anio = str(datetime.now().year)
    print(f"Año actual: {anio}")
    print("")

    asegurar_librerias()
    service = crear_servicio()

    print("Leyendo Sheets...")
    filas = leer_filas(service)
    print(f"Total filas: {len(filas) - 1}")
    print("")

    actualizadas = 0
    saltadas     = 0

    for idx, row in enumerate(filas):
        if idx == 0:
            continue  # saltar header

        domicilio   = row[COL_DOMICILIO].strip() if len(row) > COL_DOMICILIO else ""
        tipo        = row[COL_TIPO].strip()       if len(row) > COL_TIPO       else ""
        resultado   = row[COL_RESULTADO].strip()  if len(row) > COL_RESULTADO  else ""
        titulo_cert = row[COL_TITULO_CERT].strip() if len(row) > COL_TITULO_CERT else ""
        ruta_cert   = row[COL_RUTA_CERT].strip()   if len(row) > COL_RUTA_CERT   else ""

        # Solo procesar filas con resultado OK y sin ruta ya cargada
        if resultado.upper() != "OK":
            continue
        if ruta_cert:
            saltadas += 1
            continue

        categoria = detectar_tipo(tipo)
        if not categoria:
            continue

        fila_sheets = idx + 1

        if categoria == "AGUA":
            titulo = "CONSUMO"
            ruta   = armar_ruta(domicilio, "AGUA", anio)
            escribir_celda(service, fila_sheets, "G", titulo)
            escribir_celda(service, fila_sheets, "H", ruta)
            print(f"  [AGUA]      fila {fila_sheets} — {domicilio}")
            actualizadas += 1

        elif categoria == "ELECTRICA":
            titulo = "CONSUMO"
            ruta   = armar_ruta(domicilio, "ELECTRICIDAD", anio)
            escribir_celda(service, fila_sheets, "G", titulo)
            escribir_celda(service, fila_sheets, "H", ruta)
            print(f"  [ELECTRICA] fila {fila_sheets} — {domicilio}")
            actualizadas += 1

        elif categoria == "GAS":
            titulo = "CONSUMO ||CERTIFICADO DE HERMETICIDAD"
            ruta   = armar_ruta(domicilio, "GAS", anio) + " || "
            escribir_celda(service, fila_sheets, "G", titulo)
            escribir_celda(service, fila_sheets, "H", ruta)
            print(f"  [GAS]       fila {fila_sheets} — {domicilio}  ← completar ruta hermeticidad a mano")
            actualizadas += 1

    print("")
    print("=" * 50)
    print(f"Completado: {actualizadas} fila(s) actualizada(s) | {saltadas} ya tenían ruta")
    print("=" * 50)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
