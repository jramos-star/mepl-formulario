# -*- coding: utf-8 -*-
"""
mepl_saneamiento.py
-------------------
Lee los PDFs de certificados de saneamiento de la carpeta del mes actual,
extrae la fecha de cada uno, y completa automáticamente en el Sheets:
  - Col E (RESPUESTAS): "se realizo saneamiento el dia DD/MM/YYYY||se realizo saneamiento el dia DD/MM/YYYY"
  - Col G (TITULO CERT): "CERTIFICADO DE SANEAMIENTO"
  - Col H (RUTA CERT):   ruta completa del PDF

Detecta el mes y año automáticamente.
Asocia cada PDF al colegio correcto buscando nombre de calle + número.

Ejecutar antes del launcher cada mes.
"""

import sys
import subprocess
import re
from pathlib import Path
from datetime import datetime

# ========= CONFIGURACION =========
SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"
HOJA = "MAYO 2026"

BASE_CERTIFICADOS = r"C:\Users\usuario\Desktop\jorge ex\CERTIFICADOS"
# Ruta completa: BASE_CERTIFICADOS\{MES AÑO}\SANEAMIENTO\*.pdf

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
# =================================


def asegurar_librerias():
    try:
        import pdfplumber  # noqa
        import google.auth  # noqa
    except Exception:
        print("Instalando librerias...")
        subprocess.check_call([sys.executable, "-m", "pip", "install",
            "pdfplumber", "google-auth", "google-auth-oauthlib",
            "google-api-python-client"])


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


def extraer_fecha_pdf(ruta_pdf):
    """
    Extrae la fecha del PDF buscando el patrón 'el día DD/MM/YY' o 'el día DD/MM/YYYY'.
    Devuelve la fecha como string DD/MM/YYYY o None si no la encuentra.
    """
    try:
        import pdfplumber
        with pdfplumber.open(str(ruta_pdf)) as pdf:
            for page in pdf.pages:
                texto = page.extract_text() or ""
                # Buscar patrón: "el día DD/MM/YY" o "el día DD/MM/YYYY"
                match = re.search(
                    r'el\s+d[ií]a\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})',
                    texto, re.IGNORECASE
                )
                if match:
                    fecha_raw = match.group(1).replace("-", "/")
                    partes = fecha_raw.split("/")
                    if len(partes) == 3:
                        d, m, a = partes
                        if len(a) == 2:
                            a = "20" + a
                        return f"{d.zfill(2)}/{m.zfill(2)}/{a}"
    except Exception as e:
        print(f"    Error leyendo PDF: {e}")
    return None


def normalizar_para_busqueda(texto):
    """
    Normaliza un texto para comparación:
    - Elimina prefijos de calle (AVDA, AV, CALLE, DR, etc.)
    - Convierte a minúsculas
    - Elimina puntos y comas
    """
    texto = texto.upper()
    # Eliminar prefijos comunes
    prefijos = [
        r'\bAVDA\b\.?', r'\bAV\b\.?', r'\bAVENIDA\b',
        r'\bCALLE\b', r'\bCLL\b\.?',
        r'\bDR\b\.?', r'\bDOCTOR\b',
        r'\bGENERAL\b', r'\bGRAL\b\.?',
        r'\bCORNEL\b', r'\bCNL\b\.?',
        r'\bINGENIERO\b', r'\bING\b\.?',
        r'\bPRESIDENTE\b', r'\bPTE\b\.?',
    ]
    for p in prefijos:
        texto = re.sub(p, '', texto)
    # Eliminar puntos, comas, guiones extra
    texto = re.sub(r'[.,\-]', ' ', texto)
    # Colapsar espacios
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto.lower()


def extraer_clave_domicilio(domicilio):
    """
    Extrae la parte clave del domicilio para buscar en el nombre del PDF.
    Ejemplo: 'AVDA. LA PLATA 623' → 'la plata 623'
    """
    normalizado = normalizar_para_busqueda(domicilio)
    return normalizado.strip()


def buscar_pdf_para_domicilio(domicilio, pdfs_disponibles):
    """
    Busca entre los PDFs disponibles el que corresponde al domicilio.
    Compara nombre de calle + número ignorando prefijos.
    """
    clave = extraer_clave_domicilio(domicilio)

    # Extraer palabras clave — incluir números aunque sean cortos
    palabras = [p for p in clave.split() if len(p) > 2 or p.isdigit()]

    mejor_pdf = None
    mejor_score = 0

    for pdf_path in pdfs_disponibles:
        nombre_pdf = normalizar_para_busqueda(pdf_path.stem)
        score = sum(1 for p in palabras if p in nombre_pdf)
        if score > mejor_score:
            mejor_score = score
            mejor_pdf = pdf_path

    # Requerir al menos 2 coincidencias (nombre + número)
    if mejor_score >= 2:
        return mejor_pdf
    return None


def contar_tareas_saneamiento(filas, idx):
    """Cuenta cuántas tareas tiene la fila de saneamiento (cantidad de || + 1 en respuestas existentes, o 2 por defecto)."""
    row = filas[idx]
    respuestas = row[COL_RESPUESTAS].strip() if len(row) > COL_RESPUESTAS else ""
    if respuestas:
        return len(respuestas.split("||"))
    return 2  # default: 2 tareas (desinfección + desinfestación)


def main():
    print("=" * 55)
    print("   MEPL — Autocompletar Saneamiento desde PDFs")
    print("=" * 55)
    print("")

    ahora = datetime.now()
    mes_nombre = MESES_ES[ahora.month]
    anio = str(ahora.year)
    carpeta_mes = f"{mes_nombre} {anio}"

    carpeta_saneamiento = Path(BASE_CERTIFICADOS) / carpeta_mes / "SANEAMIENTO"
    print(f"Carpeta: {carpeta_saneamiento}")

    if not carpeta_saneamiento.exists():
        print(f"\nERROR: No se encontró la carpeta:")
        print(f"  {carpeta_saneamiento}")
        print("Verificá que la carpeta existe y el nombre es correcto.")
        input("\nPresione Enter para salir...")
        return

    pdfs = list(carpeta_saneamiento.glob("*.pdf"))
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
    sin_fecha    = 0

    for idx, row in enumerate(filas):
        if idx == 0:
            continue

        domicilio   = row[COL_DOMICILIO].strip() if len(row) > COL_DOMICILIO else ""
        tipo        = row[COL_TIPO].strip()       if len(row) > COL_TIPO       else ""
        resultado   = row[COL_RESULTADO].strip()  if len(row) > COL_RESULTADO  else ""
        respuestas  = row[COL_RESPUESTAS].strip() if len(row) > COL_RESPUESTAS else ""
        ruta_cert   = row[COL_RUTA_CERT].strip()  if len(row) > COL_RUTA_CERT  else ""

        # Solo filas de SANEAMIENTO
        if "SANEAMIENTO" not in tipo.upper():
            continue

        titulo_actual = row[COL_TITULO_CERT].strip() if len(row) > COL_TITULO_CERT else ""

        # Saltar solo si ya tiene tanto el certificado como la respuesta de saneamiento
        if "CERTIFICADO DE SANEAMIENTO" in titulo_actual and "se realizo saneamiento" in respuestas.lower():
            saltadas += 1
            continue

        # Buscar PDF correspondiente
        pdf_path = buscar_pdf_para_domicilio(domicilio, pdfs)
        if not pdf_path:
            print(f"  [SIN PDF]   {domicilio}")
            sin_pdf += 1
            continue

        # Extraer fecha del PDF
        fecha = extraer_fecha_pdf(pdf_path)
        if not fecha:
            print(f"  [SIN FECHA] {domicilio} — {pdf_path.name}")
            sin_fecha += 1
            continue

        # Armar respuesta: una por cada tarea de saneamiento
        cant_tareas = contar_tareas_saneamiento(filas, idx)
        texto_respuesta = f"se realizo saneamiento el dia {fecha}"
        respuesta_final = "||".join([texto_respuesta] * cant_tareas)

        fila_sheets = idx + 1

        # Solo escribir respuesta si está vacía o no tiene saneamiento
        if not respuestas or "se realizo saneamiento" not in respuestas.lower():
            cant_tareas = contar_tareas_saneamiento(filas, idx)
            texto_respuesta = f"se realizo saneamiento el dia {fecha}"
            respuesta_final = "||".join([texto_respuesta] * cant_tareas)
            escribir_celda(service, fila_sheets, "E", respuesta_final)

        # Agregar certificado de saneamiento a lo existente
        if titulo_actual:
            titulo_final = titulo_actual + "||CERTIFICADO DE SANEAMIENTO"
            ruta_final   = ruta_cert + "||" + str(pdf_path)
        else:
            titulo_final = "CERTIFICADO DE SANEAMIENTO"
            ruta_final   = str(pdf_path)

        escribir_celda(service, fila_sheets, "G", titulo_final)
        escribir_celda(service, fila_sheets, "H", ruta_final)

        print(f"  [OK] fila {fila_sheets} — {domicilio} — {fecha} — {pdf_path.name}")
        actualizadas += 1

    print("")
    print("=" * 55)
    print(f"Actualizadas : {actualizadas}")
    print(f"Ya tenían ruta (saltadas): {saltadas}")
    print(f"Sin PDF encontrado: {sin_pdf}")
    print(f"Sin fecha en PDF  : {sin_fecha}")
    print("=" * 55)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
