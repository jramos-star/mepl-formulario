# -*- coding: utf-8 -*-
"""
supabase_cargar_datos.py
------------------------
Carga los datos del MEPL MAYO y JUNIO a Supabase.
Ejecutar una vez al mes cuando se actualizan las ordenes.
"""

import sys
import subprocess
import requests
from pathlib import Path

# ========= CONFIGURACION =========
SUPABASE_URL = "https://zynvptlesftdpyijxvpv.supabase.co"
SUPABASE_KEY = "sb_publishable_a5BtY5K2O8w2rPBlIeKRUg_u6z478MX"

ID_CLAVES    = "14TEgiwQCOsRNF3cdPosbpzvauErL94131PaVpxEeuhk"    # MEPL CLAVES (Mayo y Junio)
HOJA_MAYO    = "MAYO 2026"
HOJA_JUNIO   = "JUNIO 2026"

ID_TITULOS    = "1erL0LPLtpcjDE9Z0G7oBnRDonbdAbUBhJ7Fy4FukWIY"   # MEPL TITULOS (colegios Mayo y Junio)
HOJA_MAYO_COL  = "MAYO 2026"
HOJA_JUNIO_COL = "JUNIO 2026"

ID_ANEXO     = "105WV_ORWzHvoll_4s08Y8sjAPQ6lZRbUEXv-x9p_DGA"   # ANEXO
HOJA_ANEXO   = "ANEXO5"

BASE_DIR         = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token_mepl.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
# =================================

HEADERS_SB = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}


def asegurar_librerias():
    try:
        import google.auth  # noqa
    except Exception:
        print("Instalando librerias...")
        subprocess.check_call([sys.executable, "-m", "pip", "install",
            "google-auth", "google-auth-oauthlib", "google-api-python-client", "requests"])


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


def leer_hoja(service, spreadsheet_id, rango):
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=rango
    ).execute()
    return result.get("values", [])


def limpiar_tabla(tabla):
    funcion = f"truncate_{tabla}"
    res = requests.post(
        f"{SUPABASE_URL}/rest/v1/rpc/{funcion}",
        headers=HEADERS_SB,
        json={}
    )
    print(f"  Tabla {tabla} limpiada (status {res.status_code})")


def cargar_en_supabase(tabla, filas, batch=200):
    total = 0
    for i in range(0, len(filas), batch):
        batch_data = filas[i:i+batch]
        res = requests.post(
            f"{SUPABASE_URL}/rest/v1/{tabla}",
            headers=HEADERS_SB,
            json=batch_data
        )
        if not res.ok:
            print(f"  ERROR al cargar batch: {res.text}")
        else:
            total += len(batch_data)
    return total


def formatear_fecha(fecha):
    try:
        if not fecha or str(fecha).strip() == "":
            return ""
        s = str(fecha).strip()
        if "/" in s or "-" in s:
            return s
        n = float(s)
        from datetime import datetime, timedelta
        base = datetime(1899, 12, 30)
        d = base + timedelta(days=n)
        return f"{d.day}/{d.month}/{d.year}"
    except Exception:
        return str(fecha).strip()


def procesar_claves(service, hoja_nombre, mes_label, descripciones, idx_offset=0):
    print(f"\n  Leyendo hoja '{hoja_nombre}'...")
    data = leer_hoja(service, ID_CLAVES, f"'{hoja_nombre}'!A:G")
    filas = []
    for idx, row in enumerate(data[1:]):
        if len(row) < 4:
            continue
        orden      = str(row[0]).strip() if len(row) > 0 else ""
        orden_ppal = str(row[2]).strip() if len(row) > 2 else ""
        clave      = str(row[3]).strip() if len(row) > 3 else ""
        texto      = str(row[4]).strip() if len(row) > 4 else ""

        if not orden or not clave:
            continue

        filas.append({
            "orden": orden,
            "orden_principal": orden_ppal,
            "clave": clave,
            "texto_breve": texto,
            "descripcion": descripciones.get(clave, ""),
            "orden_fila": idx + idx_offset,
            "mes": mes_label
        })
    print(f"  {len(filas)} filas leídas de {mes_label}")
    return filas


def main():
    print("=" * 50)
    print("   MEPL — Cargar datos a Supabase")
    print("=" * 50)
    print("")

    asegurar_librerias()

    print("Conectando con Google Sheets...")
    service = crear_servicio_sheets()

    # ── 1. MEPL TITULOS → colegios_hojas (Mayo y Junio) ───────
    print("\n[1/3] Leyendo MEPL TITULOS (colegios Mayo y Junio)...")

    def procesar_colegios(data, mes_label):
        filas = []
        for row in data[1:]:
            if len(row) < 3:
                continue
            orden      = str(row[0]).strip() if len(row) > 0 else ""
            orden_ppal = str(row[1]).strip() if len(row) > 1 else ""
            domicilio  = str(row[2]).strip() if len(row) > 2 else ""
            fecha_ini  = formatear_fecha(row[3]) if len(row) > 3 else ""
            fecha_fin  = formatear_fecha(row[4]) if len(row) > 4 else ""
            texto      = str(row[5]).strip() if len(row) > 5 else ""
            if not orden or not domicilio:
                continue
            filas.append({
                "orden": orden,
                "orden_principal": orden_ppal,
                "domicilio": domicilio,
                "fecha_inicio": fecha_ini,
                "fecha_fin": fecha_fin,
                "texto_breve": texto,
                "mes": mes_label
            })
        print(f"  {len(filas)} filas leídas de {mes_label}")
        return filas

    data_mayo  = leer_hoja(service, ID_TITULOS, f"'{HOJA_MAYO_COL}'!A:H")
    data_junio = leer_hoja(service, ID_TITULOS, f"'{HOJA_JUNIO_COL}'!A:H")

    filas_colegios = procesar_colegios(data_mayo, "Mayo") + procesar_colegios(data_junio, "Junio")

    print(f"  Total: {len(filas_colegios)} filas")
    print("  Limpiando tabla colegios_hojas...")
    limpiar_tabla("colegios_hojas")
    cargadas = cargar_en_supabase("colegios_hojas", filas_colegios)
    print(f"  {cargadas} filas cargadas en colegios_hojas")

    # ── 2. ANEXO → descripciones ────────────────────────────
    print("\n[2/3] Leyendo Anexo...")
    data_anexo = leer_hoja(service, ID_ANEXO, f"'{HOJA_ANEXO}'!A:B")

    descripciones = {}
    for row in data_anexo:
        if len(row) >= 2:
            clave = str(row[0]).strip()
            desc  = str(row[1]).strip()
            if clave:
                descripciones[clave] = desc

    print(f"  {len(descripciones)} claves con descripción")

    # ── 3. MEPL CLAVES Mayo y Junio → claves_tareas ─────────
    print("\n[3/3] Leyendo MEPL CLAVES (Mayo y Junio)...")
    filas_mayo  = procesar_claves(service, HOJA_MAYO,  "Mayo",  descripciones, idx_offset=0)
    filas_junio = procesar_claves(service, HOJA_JUNIO, "Junio", descripciones, idx_offset=len(filas_mayo))
    filas_claves = filas_mayo + filas_junio

    print(f"\n  Total: {len(filas_claves)} filas")
    print("  Limpiando tabla claves_tareas...")
    limpiar_tabla("claves_tareas")
    cargadas = cargar_en_supabase("claves_tareas", filas_claves)
    print(f"  {cargadas} filas cargadas en claves_tareas")

    print("")
    print("=" * 50)
    print("Carga completada exitosamente.")
    print("El formulario ya puede leer los datos desde Supabase.")
    print("=" * 50)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
