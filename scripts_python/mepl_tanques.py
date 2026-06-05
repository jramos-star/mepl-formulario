# -*- coding: utf-8 -*-
"""
mepl_tanques.py
---------------
Lee los PDFs de certificados de tanques de la carpeta TANQUES\2026,
extrae la fecha y domicilio de cada uno, y completa automáticamente en el Sheets:
  - Col E (RESPUESTAS): "se realizo limpieza el dia DD/MM/YYYY"
  - Col G (TITULO CERT): "CERTIFICADO DE LDT||CERTIFICADO FQ||CERTIFICADO DE BAC"
  - Col H (RUTA CERT):   ruta_cert||ruta_fq||ruta_bac

Carpetas:
  CERT → Certificado de Limpieza y Desinfección de Tanques (LDT)
  FQ   → Físico-Químico
  BAC  → Bacteriológico
"""

import sys
import subprocess
import re
from pathlib import Path
from datetime import datetime

# ========= CONFIGURACION =========
SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"
HOJA           = "MAYO 2026"

BASE_TANQUES = r"C:\Users\usuario\Desktop\jorge ex\TANQUES\2026"
# Subcarpetas: CERT, FQ, BAC

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


def extraer_fecha_y_domicilio_pdf(ruta_pdf):
    """
    Extrae la fecha (patrón FECHA: DD/MM/YYYY) y el domicilio del PDF.
    Devuelve (fecha_str, domicilio_str) o (None, None).
    """
    try:
        import pdfplumber
        with pdfplumber.open(str(ruta_pdf)) as pdf:
            for page in pdf.pages:
                texto = page.extract_text() or ""

                # Buscar fecha: FECHA: DD/MM/YYYY
                match_fecha = re.search(
                    r'FECHA[:\s]+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})',
                    texto, re.IGNORECASE
                )
                fecha = None
                if match_fecha:
                    fecha_raw = match_fecha.group(1).replace("-", "/")
                    partes = fecha_raw.split("/")
                    if len(partes) == 3:
                        d, m, a = partes
                        if len(a) == 2:
                            a = "20" + a
                        fecha = f"{d.zfill(2)}/{m.zfill(2)}/{a}"

                # Buscar domicilio: DOMICILIO: texto, CABA
                match_dom = re.search(
                    r'DOMICILIO[:\s]+(.+?)(?:,\s*CABA|$)',
                    texto, re.IGNORECASE
                )
                domicilio = None
                if match_dom:
                    domicilio = match_dom.group(1).strip()

                if fecha:
                    return fecha, domicilio
    except Exception as e:
        print(f"    Error leyendo PDF: {e}")
    return None, None


def normalizar(texto):
    texto = texto.upper()
    prefijos = [
        r'\bAVDA\b\.?', r'\bAV\b\.?', r'\bAVENIDA\b',
        r'\bCALLE\b', r'\bDR\b\.?', r'\bDOCTOR\b',
        r'\bGENERAL\b', r'\bGRAL\b\.?',
    ]
    for p in prefijos:
        texto = re.sub(p, '', texto)
    texto = re.sub(r'[.,\-]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto.lower()


def buscar_pdf_por_domicilio(domicilio_sheet, pdfs):
    """Busca el PDF que mejor coincide con el domicilio del Sheet."""
    clave = normalizar(domicilio_sheet)
    palabras = [p for p in clave.split() if len(p) > 2 or p.isdigit()]
    # Separar números y palabras
    numeros  = [p for p in palabras if p.isdigit()]
    palabras_calle = [p for p in palabras if not p.isdigit() and len(p) > 3]

    mejor_pdf = None
    mejor_score = 0

    for pdf_path in pdfs:
        nombre = normalizar(pdf_path.stem)
        # Score por número de calle
        score_num  = sum(1 for n in numeros if n in nombre.split())
        # Score por palabras de la calle
        score_cal  = sum(1 for p in palabras_calle if p in nombre)
        score = score_num * 2 + score_cal  # número tiene más peso
        if score > mejor_score:
            mejor_score = score
            mejor_pdf = pdf_path

    # Necesita al menos el número O dos palabras de la calle
    if mejor_score >= 2:
        return mejor_pdf
    return None


def buscar_pdf_por_contenido(domicilio_sheet, pdfs):
    """Busca el PDF leyendo el domicilio dentro del contenido del PDF."""
    clave = normalizar(domicilio_sheet)
    palabras = [p for p in clave.split() if len(p) > 2 or p.isdigit()]

    for pdf_path in pdfs:
        _, dom_pdf = extraer_fecha_y_domicilio_pdf(pdf_path)
        if dom_pdf:
            dom_norm = normalizar(dom_pdf)
            score = sum(1 for p in palabras if p in dom_norm)
            if score >= 2:
                return pdf_path
    return None


def main():
    print("=" * 55)
    print("   MEPL — Autocompletar Tanques desde PDFs")
    print("=" * 55)
    print("")

    carpeta_cert = Path(BASE_TANQUES) / "CERT"
    carpeta_fq   = Path(BASE_TANQUES) / "FQ"
    carpeta_bac  = Path(BASE_TANQUES) / "BAC"

    for carpeta in [carpeta_cert, carpeta_fq, carpeta_bac]:
        if not carpeta.exists():
            print(f"ERROR: No se encontró la carpeta: {carpeta}")
            input("\nPresione Enter para salir...")
            return

    pdfs_cert = list(carpeta_cert.glob("*.pdf"))
    pdfs_fq   = list(carpeta_fq.glob("*.pdf"))
    pdfs_bac  = list(carpeta_bac.glob("*.pdf"))

    print(f"PDFs CERT: {len(pdfs_cert)} | FQ: {len(pdfs_fq)} | BAC: {len(pdfs_bac)}")
    print("")

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
    faltantes    = []  # lista de {domicilio, falta: []}

    for idx, row in enumerate(filas):
        if idx == 0:
            continue

        domicilio  = row[COL_DOMICILIO].strip()  if len(row) > COL_DOMICILIO  else ""
        tipo       = row[COL_TIPO].strip()        if len(row) > COL_TIPO        else ""
        respuestas = row[COL_RESPUESTAS].strip()  if len(row) > COL_RESPUESTAS  else ""
        ruta_cert  = row[COL_RUTA_CERT].strip()   if len(row) > COL_RUTA_CERT   else ""

        if "SANEAMIENTO" not in tipo.upper():
            continue

        # Saltar solo si ya tiene los 3 certificados
        titulo_actual_check = row[COL_TITULO_CERT].strip() if len(row) > COL_TITULO_CERT else ""
        if ("CERTIFICADO DE LDT" in titulo_actual_check and
            "CERTIFICADO FQ" in titulo_actual_check and
            "CERTIFICADO DE BAC" in titulo_actual_check):
            saltadas += 1
            continue

        # Buscar los 3 certificados siempre
        pdf_cert = buscar_pdf_por_domicilio(domicilio, pdfs_cert)
        if not pdf_cert:
            pdf_cert = buscar_pdf_por_contenido(domicilio, pdfs_cert)

        pdf_fq = buscar_pdf_por_domicilio(domicilio, pdfs_fq)
        if not pdf_fq:
            pdf_fq = buscar_pdf_por_contenido(domicilio, pdfs_fq)

        pdf_bac = buscar_pdf_por_domicilio(domicilio, pdfs_bac)
        if not pdf_bac:
            pdf_bac = buscar_pdf_por_contenido(domicilio, pdfs_bac)

        # Si no hay ninguno, saltar
        if not pdf_cert and not pdf_fq and not pdf_bac:
            print(f"  [SIN CERT/FQ/BAC] {domicilio}")
            faltantes.append({"domicilio": domicilio, "falta": ["CERT", "FQ", "BAC"]})
            sin_pdf += 1
            continue

        # Extraer fecha del CERT si existe
        fecha = None
        if pdf_cert:
            fecha, _ = extraer_fecha_y_domicilio_pdf(pdf_cert)
            if not fecha:
                print(f"  [SIN FECHA] {domicilio} — {pdf_cert.name}")
                sin_fecha += 1

        fila_sheets = idx + 1

        # Leer lo que ya hay en el Sheet
        titulo_actual = row[COL_TITULO_CERT].strip() if len(row) > COL_TITULO_CERT else ""
        ruta_actual   = row[COL_RUTA_CERT].strip()   if len(row) > COL_RUTA_CERT   else ""

        # Solo agregar los que faltan
        titulos_nuevos = []
        rutas_nuevas   = []
        falta_este     = []

        if pdf_cert and "CERTIFICADO DE LDT" not in titulo_actual:
            titulos_nuevos.append("CERTIFICADO DE LDT")
            rutas_nuevas.append(str(pdf_cert))
        elif not pdf_cert and "CERTIFICADO DE LDT" not in titulo_actual:
            print(f"  [SIN CERT]  {domicilio}")
            falta_este.append("CERT")

        if pdf_fq and "CERTIFICADO FQ" not in titulo_actual:
            titulos_nuevos.append("CERTIFICADO FQ")
            rutas_nuevas.append(str(pdf_fq))
        elif not pdf_fq and "CERTIFICADO FQ" not in titulo_actual:
            print(f"  [SIN FQ]    {domicilio}")
            falta_este.append("FQ")

        if pdf_bac and "CERTIFICADO DE BAC" not in titulo_actual:
            titulos_nuevos.append("CERTIFICADO DE BAC")
            rutas_nuevas.append(str(pdf_bac))
        elif not pdf_bac and "CERTIFICADO DE BAC" not in titulo_actual:
            print(f"  [SIN BAC]   {domicilio}")
            falta_este.append("BAC")

        if falta_este:
            faltantes.append({"domicilio": domicilio, "falta": falta_este})

        # Si no hay nada nuevo que agregar, saltar
        if not titulos_nuevos:
            saltadas += 1
            continue

        # Escribir respuesta solo si hay fecha y no tiene respuesta de tanques
        if fecha and "se realizo limpieza" not in respuestas:
            respuesta_tanques = f"se realizo limpieza el dia {fecha}"
            respuesta_final = (respuestas + "||" + respuesta_tanques) if respuestas else respuesta_tanques
            escribir_celda(service, fila_sheets, "E", respuesta_final)

        # Agregar solo los nuevos
        titulo_final = (titulo_actual + "||" + "||".join(titulos_nuevos)) if titulo_actual else "||".join(titulos_nuevos)
        ruta_final   = (ruta_actual   + "||" + "||".join(rutas_nuevas))   if ruta_actual   else "||".join(rutas_nuevas)

        escribir_celda(service, fila_sheets, "G", titulo_final)
        escribir_celda(service, fila_sheets, "H", ruta_final)

        print(f"  [OK] fila {fila_sheets} — {domicilio} — {fecha or 'sin fecha'} — {len(titulos_nuevos)} cert. nuevo(s)")
        actualizadas += 1

    print("")
    print("=" * 55)
    print(f"Actualizadas : {actualizadas}")
    print(f"Saltadas     : {saltadas}")
    print(f"Sin CERT     : {sin_pdf}")
    print(f"Sin fecha    : {sin_fecha}")

    if faltantes:
        print("")
        print("FALTANTES:")
        for f in faltantes:
            print(f"  {f['domicilio']:<45} → sin {', '.join(f['falta'])}")

    print("=" * 55)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
