# -*- coding: utf-8 -*-
"""
mepl_estructura.py
------------------
Lee los PDFs de informes de estructura de la carpeta INFORMES\ESTRUCTURA,
y completa automáticamente en el Sheets:
  - Col E (RESPUESTAS): "se adjunta informe||se adjunta informe||..." (5 veces)
  - Col G (TITULO CERT): "INFORME DE ESTRUCTURA"
  - Col H (RUTA CERT):   ruta completa del PDF
"""

import sys
import subprocess
import re
from pathlib import Path

# ========= CONFIGURACION =========
SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"
HOJA           = "MAYO 2026"

BASE_ESTRUCTURA = r"C:\Users\usuario\Desktop\jorge ex\INFORMES\ESTRUCTURA"

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

RESPUESTA_FIJA = "||".join(["se adjunta informe"] * 5)
TITULO_CERT    = "INFORME DE ESTRUCTURA"
# =================================


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


def normalizar(texto):
    texto = texto.upper()
    prefijos = [
        r'\bAVDA\b\.?', r'\bAV\b\.?', r'\bAVENIDA\b',
        r'\bCALLE\b', r'\bDR\b\.?', r'\bDOCTOR\b',
        r'\bGENERAL\b', r'\bGRAL\b\.?',
    ]
    for p in prefijos:
        texto = re.sub(p, '', texto)
    texto = re.sub(r'[.,\-/]', ' ', texto)  # incluir slash
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto.lower()


def buscar_pdf_por_domicilio(domicilio_sheet, pdfs):
    clave = normalizar(domicilio_sheet)
    palabras = [p for p in clave.split() if len(p) > 2 or p.isdigit()]
    numeros        = [p for p in palabras if p.isdigit()]
    palabras_calle = [p for p in palabras if not p.isdigit() and len(p) > 3]

    mejor_pdf   = None
    mejor_score = 0

    for pdf_path in pdfs:
        nombre = normalizar(pdf_path.stem)
        # Para números, verificar si alguno de los números del domicilio
        # aparece en cualquier parte del nombre (no solo como palabra completa)
        score_num = sum(1 for n in numeros if n in nombre)
        score_cal = sum(1 for p in palabras_calle if p in nombre)
        score = score_num * 2 + score_cal
        if score > mejor_score:
            mejor_score = score
            mejor_pdf = pdf_path

    return mejor_pdf if mejor_score >= 2 else None


def main():
    print("=" * 55)
    print("   MEPL — Autocompletar Estructura desde PDFs")
    print("=" * 55)
    print("")

    carpeta = Path(BASE_ESTRUCTURA)
    if not carpeta.exists():
        print(f"ERROR: No se encontró la carpeta: {carpeta}")
        input("\nPresione Enter para salir...")
        return

    pdfs = list(carpeta.glob("*.pdf"))
    print(f"PDFs encontrados: {len(pdfs)}")
    print("")

    if not pdfs:
        print("No hay PDFs en la carpeta.")
        input("\nPresione Enter para salir...")
        return

    asegurar_librerias()
    service = crear_servicio()

    print("Leyendo Sheets...")
    filas = leer_filas(service)
    print(f"Total filas: {len(filas) - 1}")
    print("")

    actualizadas = 0
    saltadas     = 0
    sin_pdf      = 0
    faltantes    = []

    for idx, row in enumerate(filas):
        if idx == 0:
            continue

        domicilio  = row[COL_DOMICILIO].strip() if len(row) > COL_DOMICILIO else ""
        tipo       = row[COL_TIPO].strip()       if len(row) > COL_TIPO       else ""
        respuestas = row[COL_RESPUESTAS].strip() if len(row) > COL_RESPUESTAS else ""
        titulo_actual = row[COL_TITULO_CERT].strip() if len(row) > COL_TITULO_CERT else ""
        ruta_actual   = row[COL_RUTA_CERT].strip()   if len(row) > COL_RUTA_CERT   else ""

        if "ESTRUCTURA" not in tipo.upper():
            continue

        # Saltar si ya tiene el informe
        if TITULO_CERT in titulo_actual:
            saltadas += 1
            continue

        # Buscar PDF
        pdf = buscar_pdf_por_domicilio(domicilio, pdfs)
        if not pdf:
            print(f"  [SIN PDF]   {domicilio}")
            faltantes.append(domicilio)
            sin_pdf += 1
            continue

        fila_sheets = idx + 1

        # Escribir respuesta si no tiene
        if not respuestas:
            escribir_celda(service, fila_sheets, "E", RESPUESTA_FIJA)

        # Agregar título y ruta a lo existente
        titulo_final = (titulo_actual + "||" + TITULO_CERT) if titulo_actual else TITULO_CERT
        ruta_final   = (ruta_actual   + "||" + str(pdf))    if ruta_actual   else str(pdf)

        escribir_celda(service, fila_sheets, "G", titulo_final)
        escribir_celda(service, fila_sheets, "H", ruta_final)

        print(f"  [OK] fila {fila_sheets} — {domicilio} — {pdf.name}")
        actualizadas += 1

    print("")
    print("=" * 55)
    print(f"Actualizadas : {actualizadas}")
    print(f"Saltadas     : {saltadas}")
    print(f"Sin PDF      : {sin_pdf}")

    if faltantes:
        print("")
        print("FALTANTES:")
        for d in faltantes:
            print(f"  {d}")

    print("=" * 55)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
