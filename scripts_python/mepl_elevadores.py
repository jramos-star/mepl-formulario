# -*- coding: utf-8 -*-
"""
mepl_elevadores.py
------------------
Lee los certificados (JPEG) de elevadores y plataformas elevadoras,
y completa automáticamente en el Sheets:
  - Col E (RESPUESTAS): "se realizo mantenimiento||se realizo mantenimiento"
  - Col G (TITULO CERT): "CERTIFICADO DE ELEVADORES" o "CERTIFICADO DE PLATAFORMAS ELEVADORAS"
  - Col H (RUTA CERT):   ruta completa del archivo

Carpeta: ASCENSORES Y ELEVADORES/2026/{MES}
"""

import sys
import subprocess
import re
from pathlib import Path
from datetime import datetime

# ========= CONFIGURACION =========
SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"
HOJA           = "MAYO 2026"

BASE_ELEVADORES = r"C:\Users\usuario\Desktop\jorge ex\ASCENSORES Y ELEVADORES\2026"

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

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

TIPOS = {
    "ELEVADORES":            "CERTIFICADO DE ELEVADORES",
    "PLATAFORMAS ELEVADORAS": "CERTIFICADO DE PLATAFORMAS ELEVADORAS"
}
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
        r'\bJUAN\b', r'\bB\b\.?',
    ]
    for p in prefijos:
        texto = re.sub(p, '', texto)
    texto = re.sub(r'[.,\-]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto.lower()


def buscar_archivo_por_domicilio(domicilio, archivos):
    """Busca el archivo que mejor coincide con el domicilio."""
    clave = normalizar(domicilio)
    palabras = [p for p in clave.split() if len(p) > 2 or p.isdigit()]
    numeros       = [p for p in palabras if p.isdigit()]
    palabras_calle = [p for p in palabras if not p.isdigit() and len(p) > 3]

    mejor = None
    mejor_score = 0

    for arch in archivos:
        nombre = normalizar(arch.stem)
        score_num = sum(1 for n in numeros if n in nombre.split())
        score_cal = sum(1 for p in palabras_calle if p in nombre)
        score = score_num * 2 + score_cal
        if score > mejor_score:
            mejor_score = score
            mejor = arch

    return mejor if mejor_score >= 1 else None


def main():
    print("=" * 55)
    print("   MEPL — Autocompletar Elevadores desde archivos")
    print("=" * 55)
    print("")

    ahora = datetime.now()
    mes_nombre = MESES_ES[ahora.month]
    carpeta_mes = Path(BASE_ELEVADORES) / mes_nombre

    print(f"Carpeta: {carpeta_mes}")

    if not carpeta_mes.exists():
        print(f"\nERROR: No se encontró la carpeta: {carpeta_mes}")
        input("\nPresione Enter para salir...")
        return

    # Buscar archivos JPEG y PDF
    archivos = list(carpeta_mes.glob("*.jpeg")) + \
               list(carpeta_mes.glob("*.jpg"))  + \
               list(carpeta_mes.glob("*.pdf"))
    print(f"Archivos encontrados: {len(archivos)}")
    print("")

    if not archivos:
        print("No hay archivos en la carpeta.")
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
    sin_archivo  = 0
    faltantes    = []

    for idx, row in enumerate(filas):
        if idx == 0:
            continue

        domicilio  = row[COL_DOMICILIO].strip()  if len(row) > COL_DOMICILIO  else ""
        tipo       = row[COL_TIPO].strip()        if len(row) > COL_TIPO        else ""
        respuestas = row[COL_RESPUESTAS].strip()  if len(row) > COL_RESPUESTAS  else ""
        ruta_cert  = row[COL_RUTA_CERT].strip()   if len(row) > COL_RUTA_CERT   else ""
        titulo_actual = row[COL_TITULO_CERT].strip() if len(row) > COL_TITULO_CERT else ""

        # Solo ELEVADORES y PLATAFORMAS ELEVADORAS
        tipo_upper = tipo.upper()
        titulo_cert = None
        for key, val in TIPOS.items():
            if key in tipo_upper:
                titulo_cert = val
                break

        if not titulo_cert:
            continue

        # Saltar si ya tiene certificado de este tipo
        if titulo_cert in titulo_actual:
            saltadas += 1
            continue

        # Buscar archivo
        archivo = buscar_archivo_por_domicilio(domicilio, archivos)
        if not archivo:
            print(f"  [SIN ARCHIVO] {domicilio} — {tipo}")
            faltantes.append({"domicilio": domicilio, "tipo": tipo})
            sin_archivo += 1
            continue

        fila_sheets = idx + 1

        # Respuesta fija: dos tareas
        if not respuestas:
            escribir_celda(service, fila_sheets, "E", "se realizo mantenimiento||se realizo mantenimiento")

        # Agregar título y ruta a lo existente
        if titulo_actual:
            titulo_final = titulo_actual + "||" + titulo_cert
            ruta_final   = ruta_cert + "||" + str(archivo)
        else:
            titulo_final = titulo_cert
            ruta_final   = str(archivo)

        escribir_celda(service, fila_sheets, "G", titulo_final)
        escribir_celda(service, fila_sheets, "H", ruta_final)

        print(f"  [OK] fila {fila_sheets} — {domicilio} — {tipo} — {archivo.name}")
        actualizadas += 1

    print("")
    print("=" * 55)
    print(f"Actualizadas : {actualizadas}")
    print(f"Saltadas     : {saltadas}")
    print(f"Sin archivo  : {sin_archivo}")

    if faltantes:
        print("")
        print("FALTANTES:")
        for f in faltantes:
            print(f"  {f['domicilio']:<45} → {f['tipo']}")

    print("=" * 55)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
