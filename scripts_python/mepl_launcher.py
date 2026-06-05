# -*- coding: utf-8 -*-
"""
MEPL Launcher — Google Sheets + SAP GUI Scripting
--------------------------------------------------
Lee las filas del Sheets que tengan RESPUESTAS completas
y ejecuta MEPL_IW32_RESPUESTAS.vbs por cada una.
Escribe OK o ERROR en la columna RESULTADO.

Columnas del Sheets:
  A = MEIN (orden a modificar)
  B = ORDEN_REF (MEPL)
  C = DOMICILIO
  D = TIPO
  E = RESPUESTAS (separar tareas con ||)
  F = RESULTADO  <- escribe el launcher
  G = TITULO CERTIFICADO (titulos separados por ||, opcional)
  H = RUTA CERTIFICADO (rutas completas separadas por ||, opcional)
"""

import sys
import subprocess
import csv
import requests
from pathlib import Path
from datetime import datetime

# ========= CONFIGURACION =========
SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"

# Indices de columna (0-based)
COL_MEIN            = 0   # A
COL_ORDEN_REF       = 1   # B
COL_DOMICILIO       = 2   # C
COL_TIPO            = 3   # D
COL_RESPUESTAS      = 4   # E
COL_RESULTADO       = 5   # F
COL_TITULO_CERT     = 6   # G
COL_RUTA_CERT       = 7   # H

SUPABASE_URL = "https://zynvptlesftdpyijxvpv.supabase.co"
SUPABASE_KEY = "sb_publishable_a5BtY5K2O8w2rPBlIeKRUg_u6z478MX"

HOJA = ""  # Se define al inicio segun el mes elegido

SCRIPT_SAP = r"C:\Users\usuario\AppData\Roaming\SAP\SAP GUI\Scripts\MEPL_IW32_RESPUESTAS.vbs"

BASE_DIR         = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token_mepl.json"
LOG_FILE         = BASE_DIR / "registro_mepl.csv"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]
# =================================



def cargar_tareas_supabase(mes):
    """Carga de Supabase las claves y cantidad de tareas por orden para el mes dado."""
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}"
    }
    res = requests.get(
        f"{SUPABASE_URL}/rest/v1/claves_tareas",
        headers=headers,
        params={"select": "orden,clave", "mes": f"eq.{mes}", "order": "orden_fila.asc"}
    )
    res.raise_for_status()
    data = res.json()
    # Agrupar por orden: conteo y lista de claves
    conteo = {}
    claves = {}
    for r in data:
        orden = str(r["orden"])
        clave = str(r["clave"] or "")
        conteo[orden] = conteo.get(orden, 0) + 1
        if orden not in claves:
            claves[orden] = []
        claves[orden].append(clave)
    return conteo, claves


def seleccionar_mes():
    """Muestra un menu para seleccionar el mes a procesar."""
    meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    print("Selecciona el mes a procesar:")
    for i, m in enumerate(meses, 1):
        print(f"  {i}. {m} 2026")
    print("")
    while True:
        try:
            opcion = int(input("Numero de mes (1-12): ").strip())
            if 1 <= opcion <= 12:
                return meses[opcion - 1]
        except ValueError:
            pass
        print("Opcion invalida, ingresa un numero del 1 al 12.")

def asegurar_librerias():
    try:
        import google.auth  # noqa
        from google_auth_oauthlib.flow import InstalledAppFlow  # noqa
        from googleapiclient.discovery import build  # noqa
    except Exception:
        print("Faltan librerias de Google. Instalando...")
        paquetes = [
            "google-auth",
            "google-auth-oauthlib",
            "google-api-python-client",
        ]
        subprocess.check_call([sys.executable, "-m", "pip", "install", *paquetes])


def crear_servicio_sheets():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    if not CREDENTIALS_FILE.exists():
        raise FileNotFoundError(
            f"No encuentro credentials.json en {CREDENTIALS_FILE}\n"
            "Copialo en la misma carpeta que este archivo."
        )

    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")

    return build("sheets", "v4", credentials=creds)


def leer_filas(service):
    rango = f"'{HOJA}'!A:H"   # hasta columna H
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=rango
    ).execute()
    return result.get("values", [])


def escribir_resultado(service, fila_idx, resultado):
    """fila_idx es 0-based desde el array de values (fila 1 = header)."""
    fila_sheets = fila_idx + 1   # Sheets es 1-based
    rango = f"'{HOJA}'!F{fila_sheets}"
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rango,
        valueInputOption="RAW",
        body={"values": [[resultado]]}
    ).execute()


def registrar_log(mein, orden_ref, tipo, resultado, returncode="", titulo_cert="", ruta_cert=""):
    existe = LOG_FILE.exists()
    with LOG_FILE.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        if not existe:
            writer.writerow([
                "fecha_hora", "mein", "orden_ref", "tipo",
                "resultado", "returncode", "titulo_cert", "ruta_cert"
            ])
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            mein, orden_ref, tipo, resultado, returncode,
            titulo_cert, ruta_cert
        ])


def ejecutar_vbs(mein, respuestas, titulo_cert="", ruta_cert="", claves=""):
    if not Path(SCRIPT_SAP).exists():
        raise FileNotFoundError(f"No encuentro el script SAP: {SCRIPT_SAP}")

    cmd = [
        "cscript.exe",
        "//nologo",
        SCRIPT_SAP,
        mein,
        respuestas,
        titulo_cert,
        ruta_cert,
        claves,   # argumento 5: claves de modelo separadas por ||
    ]

    proceso = subprocess.run(cmd, check=False, capture_output=True, text=True)
    return proceso.returncode, proceso.stdout.strip()


def procesar_filas(service, filas, confirmar_todas, tareas_supabase, claves_supabase):
    procesadas = 0
    errores    = 0

    for idx, row in enumerate(filas):
        # Saltar header (fila 1)
        if idx == 0:
            continue

        mein         = row[COL_MEIN].strip()         if len(row) > COL_MEIN         else ""
        orden_ref    = row[COL_ORDEN_REF].strip()    if len(row) > COL_ORDEN_REF    else ""
        domicilio    = row[COL_DOMICILIO].strip()    if len(row) > COL_DOMICILIO    else ""
        tipo         = row[COL_TIPO].strip()         if len(row) > COL_TIPO         else ""
        respuestas   = row[COL_RESPUESTAS].strip()   if len(row) > COL_RESPUESTAS   else ""
        resultado    = row[COL_RESULTADO].strip()    if len(row) > COL_RESULTADO    else ""
        titulo_cert  = row[COL_TITULO_CERT].strip()  if len(row) > COL_TITULO_CERT  else ""
        ruta_cert    = row[COL_RUTA_CERT].strip()    if len(row) > COL_RUTA_CERT    else ""

        # Saltar si no tiene MEIN
        if not mein:
            continue

        # Saltar si no tiene respuestas
        if not respuestas:
            continue

        # Saltar si ya tiene resultado OK
        if resultado.upper() == "OK":
            print(f"[SALTADA] {mein} — ya tiene resultado OK")
            continue

        print("")
        print("=" * 45)
        print(f"MEIN      : {mein}")
        print(f"ORDEN REF : {orden_ref}")
        print(f"DOMICILIO : {domicilio}")
        print(f"TIPO      : {tipo}")
        print(f"RESPUESTAS: {respuestas[:80]}{'...' if len(respuestas) > 80 else ''}")
        if titulo_cert and ruta_cert:
            certs = titulo_cert.split("||")
            rutas = ruta_cert.split("||")
            for i, t in enumerate(certs):
                r = rutas[i].strip() if i < len(rutas) else ""
                print(f"CERT {i+1}    : {t.strip()} → {r}")
        else:
            print(f"CERT.     : (sin certificado)")
        print("=" * 45)

        # Validar que los archivos de certificado existen antes de ejecutar
        if ruta_cert:
            rutas = [r.strip() for r in ruta_cert.split("||") if r.strip()]
            archivos_faltantes = [r for r in rutas if not Path(r).exists()]
            if archivos_faltantes:
                print("")
                print("  ADVERTENCIA — Archivos no encontrados:")
                for af in archivos_faltantes:
                    print(f"     X {af}")
                resp_arch = input("  Continuar igual sin esos archivos? (S/N): ").strip().upper()
                if resp_arch not in ("S", "SI"):
                    print("Saltada por archivos faltantes.")
                    registrar_log(mein, orden_ref, tipo, "SALTADA - archivos faltantes", titulo_cert=titulo_cert, ruta_cert=ruta_cert)
                    continue

        # Validar cantidad de respuestas vs tareas en Supabase
        cant_resp   = len([r for r in respuestas.split("||") if r.strip()])
        cant_tareas = tareas_supabase.get(str(orden_ref), None)

        if cant_tareas is not None and cant_resp != cant_tareas:
            print(f"  [INCOMPLETA] {mein} — tiene {cant_resp} respuesta(s) pero necesita {cant_tareas}")
            print(f"  Saltando esta MEIN.")
            registrar_log(mein, orden_ref, tipo, f"INCOMPLETA ({cant_resp}/{cant_tareas})", titulo_cert=titulo_cert, ruta_cert=ruta_cert)
            continue

        # Validar certificados para AGUA, ELECTRICA y GAS
        tipo_upper = tipo.upper()
        requiere_cert = any(t in tipo_upper for t in ["AGUA", "ELECTRICA", "GAS"])
        if requiere_cert:
            titulos_cert = [t.strip() for t in titulo_cert.split("||") if t.strip()]
            rutas_cert   = [r.strip() for r in ruta_cert.split("||") if r.strip()]
            es_gas       = "GAS" in tipo_upper
            cant_req     = 2 if es_gas else 1

            if len(titulos_cert) < cant_req or len(rutas_cert) < cant_req:
                print(f"  [SIN CERT] {mein} — {tipo} requiere {cant_req} titulo(s) y ruta(s) de certificado")
                print(f"    Tiene: {len(titulos_cert)} titulo(s) y {len(rutas_cert)} ruta(s)")
                print(f"  Saltando esta MEIN.")
                registrar_log(mein, orden_ref, tipo, f"SIN CERTIFICADO ({len(titulos_cert)}/{cant_req})", titulo_cert=titulo_cert, ruta_cert=ruta_cert)
                continue

        if not confirmar_todas:
            resp = input("Ejecutar esta orden? (S/N): ").strip().upper()
            if resp not in ("S", "SI"):
                print("Saltada.")
                registrar_log(mein, orden_ref, tipo, "SALTADA", titulo_cert=titulo_cert, ruta_cert=ruta_cert)
                continue

        # Obtener claves de modelo de Supabase para esta orden
        claves_orden = claves_supabase.get(str(orden_ref), [])
        claves_str = "||".join(claves_orden)

        try:
            rc, salida = ejecutar_vbs(mein, respuestas, titulo_cert, ruta_cert, claves_str)

            if rc == 0:
                resultado_final = "OK"
                procesadas += 1
                print(f"OK — {mein}")
            else:
                resultado_final = f"ERROR (rc={rc}): {salida}" if salida else f"ERROR (rc={rc})"
                errores += 1
                print(f"ERROR rc={rc} — {mein}")
                print(f"  Detalle: {salida if salida else 'sin mensaje'}")

            escribir_resultado(service, idx, resultado_final)
            registrar_log(mein, orden_ref, tipo, resultado_final, rc, titulo_cert, ruta_cert)

        except Exception as e:
            msg = f"ERROR: {e}"
            errores += 1
            print(msg)
            escribir_resultado(service, idx, msg)
            registrar_log(mein, orden_ref, tipo, msg, titulo_cert=titulo_cert, ruta_cert=ruta_cert)

    return procesadas, errores


def main():
    print("=" * 45)
    print("   MEPL — Google Sheets + SAP IW32")
    print("=" * 45)
    print("")

    try:
        asegurar_librerias()

        # Seleccionar mes
        mes = seleccionar_mes()
        global HOJA
        HOJA = f"{mes} 2026"
        print(f"\nProcesando: {HOJA}")
        print("")

        # Cargar tareas de Supabase para validacion
        print("Cargando tareas de Supabase...")
        tareas_supabase, claves_supabase = cargar_tareas_supabase(mes)
        print(f"  {len(tareas_supabase)} orden(es) con tareas cargadas")
        print("")

        service = crear_servicio_sheets()

        while True:
            filas = leer_filas(service)
            total_con_respuesta = sum(
                1 for i, r in enumerate(filas)
                if i > 0
                and len(r) > COL_RESPUESTAS
                and r[COL_RESPUESTAS].strip()
                and (len(r) <= COL_RESULTADO or r[COL_RESULTADO].strip().upper() != "OK")
            )

            print(f"Filas con respuesta pendiente: {total_con_respuesta}")
            print("")

            if total_con_respuesta == 0:
                print("No hay filas pendientes para procesar.")
                break

            # Mostrar resumen de tareas a cargar
            print("-" * 45)
            print("  RESUMEN DE TAREAS A CARGAR:")
            print("-" * 45)
            for i, r in enumerate(filas):
                if i == 0:
                    continue
                mein_r      = r[COL_MEIN].strip()       if len(r) > COL_MEIN       else ""
                domicilio_r = r[COL_DOMICILIO].strip()  if len(r) > COL_DOMICILIO  else ""
                tipo_r      = r[COL_TIPO].strip()       if len(r) > COL_TIPO       else ""
                respuesta_r = r[COL_RESPUESTAS].strip() if len(r) > COL_RESPUESTAS else ""
                resultado_r = r[COL_RESULTADO].strip()  if len(r) > COL_RESULTADO  else ""
                titulo_r    = r[COL_TITULO_CERT].strip() if len(r) > COL_TITULO_CERT else ""

                if not mein_r or not respuesta_r:
                    continue
                if resultado_r.upper() == "OK":
                    continue

                cant_resp = len(respuesta_r.split("||"))
                cert_info = f" + {len(titulo_r.split('||'))} cert." if titulo_r else ""
                print(f"  MEIN {mein_r} | {domicilio_r[:30]:<30} | {tipo_r[:22]:<22} | {cant_resp} resp{cert_info}")
            print("-" * 45)
            print("")

            confirmar_todas = False
            if total_con_respuesta > 1:
                resp = input("Ejecutar todas sin preguntar una por una? (S/N): ").strip().upper()
                confirmar_todas = resp in ("S", "SI")

            procesadas, errores = procesar_filas(service, filas, confirmar_todas, tareas_supabase, claves_supabase)

            print("")
            print("=" * 45)
            print(f"Procesadas : {procesadas}")
            print(f"Errores    : {errores}")
            print("=" * 45)

            resp = input("Procesar otra tanda? (S/N): ").strip().upper()
            if resp not in ("S", "SI"):
                print("Fin del proceso.")
                break

    except Exception as e:
        print("")
        print("ERROR:")
        print(str(e))

    input("Presione Enter para salir...")


if __name__ == "__main__":
    main()
