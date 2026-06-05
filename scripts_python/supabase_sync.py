# -*- coding: utf-8 -*-
"""
supabase_sync.py
----------------
Lee las respuestas nuevas de Supabase y las vuelca al Google Sheets
de MEPL_RESPUESTAS. Marca cada fila como procesada en Supabase.

También lee mein_creada.txt y registra las MEINs nuevas en JUNIO 2026.

Lógica de escritura:
  - Si ya existe una fila con el mismo número en col B (ORDEN REF),
    escribe la respuesta en col E de esa fila existente.
  - Si no existe, agrega una fila nueva al final.
"""

import sys
import subprocess
import requests
import re
from pathlib import Path

# ========= CONFIGURACION =========
SUPABASE_URL = "https://zynvptlesftdpyijxvpv.supabase.co"
SUPABASE_KEY = "sb_publishable_a5BtY5K2O8w2rPBlIeKRUg_u6z478MX"

SPREADSHEET_ID = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"
HOJA_MAYO      = "MAYO 2026"
HOJA_JUNIO     = "JUNIO 2026"

BASE_DIR         = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token_mepl.json"
MEIN_FILE        = BASE_DIR / "mein_creada.txt"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

VBS_BUSCAR_MEIN = r"C:\Users\usuario\AppData\Roaming\SAP\SAP GUI\Scripts\MEPL_BUSCAR_MEIN.vbs"
# =================================


def asegurar_librerias():
    try:
        import google.auth  # noqa
    except Exception:
        print("Instalando librerias de Google...")
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


def leer_supabase():
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    res = requests.get(
        f"{SUPABASE_URL}/rest/v1/respuestas",
        headers=headers,
        params={"procesado": "eq.false", "select": "*"}
    )
    res.raise_for_status()
    return res.json()


def marcar_procesado(ids):
    if not ids:
        return
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    ids_str = ",".join(ids)
    res = requests.patch(
        f"{SUPABASE_URL}/rest/v1/respuestas",
        headers=headers,
        params={"id": f"in.({ids_str})"},
        json={"procesado": True}
    )
    print(f"  [PATCH procesado] status={res.status_code} | {res.text[:200] if res.text else 'sin respuesta'}")


def leer_columna_b(service, hoja):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{hoja}'!B:B"
    ).execute()
    filas = result.get("values", [])
    indice = {}
    for i, row in enumerate(filas):
        if i == 0:
            continue
        if row and row[0].strip():
            orden = row[0].strip()
            if orden not in indice:
                indice[orden] = i + 1
    return indice


def escribir_sheets(service, filas):
    if not filas:
        return 0, 0

    # Construir indices por hoja
    indice_mayo  = leer_columna_b(service, HOJA_MAYO)
    indice_junio = leer_columna_b(service, HOJA_JUNIO)

    actualizadas  = 0
    nuevas_mayo   = []
    nuevas_junio  = []

    for f in filas:
        orden_ref = str(f.get("orden_ref", "") or "").strip()
        respuesta = f.get("respuesta", "")
        mes       = str(f.get("mes", "") or "").strip()

        # Determinar hoja segun mes
        if mes.lower() == "junio":
            hoja    = HOJA_JUNIO
            indice  = indice_junio
            nuevas  = nuevas_junio
        else:
            hoja    = HOJA_MAYO
            indice  = indice_mayo
            nuevas  = nuevas_mayo

        if orden_ref and orden_ref in indice:
            fila_sheets = indice[orden_ref]
            rango = f"'{hoja}'!E{fila_sheets}"
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=rango,
                valueInputOption="RAW",
                body={"values": [[respuesta]]}
            ).execute()
            actualizadas += 1
            print(f"  [ACTUALIZADA] fila {fila_sheets} — orden {orden_ref} — {hoja}")
        else:
            nuevas.append([
                f.get("mein", ""),
                orden_ref,
                f.get("domicilio", ""),
                f.get("tipo", ""),
                respuesta,
                "", "", ""
            ])

    if nuevas_mayo:
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{HOJA_MAYO}'!A:H",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": nuevas_mayo}
        ).execute()
        print(f"  [NUEVAS] {len(nuevas_mayo)} fila(s) agregadas en MAYO 2026")

    if nuevas_junio:
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{HOJA_JUNIO}'!A:H",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": nuevas_junio}
        ).execute()
        print(f"  [NUEVAS] {len(nuevas_junio)} fila(s) agregadas en JUNIO 2026")

    return actualizadas, len(nuevas_mayo) + len(nuevas_junio)


def leer_mein_creadas():
    """Lee mein_creada.txt y devuelve lista de dicts con los datos (flujo legacy)."""
    if not MEIN_FILE.exists():
        return []
    filas = []
    with open(MEIN_FILE, "r", encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if not linea:
                continue
            partes = linea.split("|")
            if len(partes) >= 4:
                filas.append({
                    "mein":      partes[0].strip(),
                    "orden":     partes[1].strip(),
                    "domicilio": partes[2].strip(),
                    "tipo":      partes[3].strip()
                })
    return filas


def leer_meins_supabase():
    """Lee MEINs creadas de meins_pendientes que aun no fueron sincronizadas al Sheet."""
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    res = requests.get(
        f"{SUPABASE_URL}/rest/v1/meins_pendientes",
        headers=headers,
        params={"estado": "eq.creada", "sincronizado": "eq.false", "select": "*"}
    )
    res.raise_for_status()
    data = res.json()
    return [
        {
            "mein":      r["mein_creada"],
            "orden":     r["orden"],
            "domicilio": r["domicilio"],
            "tipo":      r["tipo"],
            "id":        r["id"]
        }
        for r in data if r.get("mein_creada")
    ]


def marcar_meins_sincronizadas(ids):
    """Marca las MEINs como sincronizadas en Supabase para no repetirlas."""
    if not ids:
        return
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    ids_str = ",".join(str(i) for i in ids)
    requests.patch(
        f"{SUPABASE_URL}/rest/v1/meins_pendientes",
        headers=headers,
        params={"id": f"in.({ids_str})"},
        json={"sincronizado": True}
    )


def buscar_mein_en_sap(orden):
    """Llama al VBS MEPL_BUSCAR_MEIN para obtener el numero de MEIN real de una orden."""
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
    except Exception:
        pass
    return None


def corregir_meins_iguales(service):
    """
    Recorre el Sheet JUNIO y para las filas donde col A = col B
    (mein = orden), busca la MEIN real en SAP y la reemplaza en col A.
    Despues valida duplicados y elimina filas repetidas.
    """
    print("\nCorrigiendo filas donde MEIN = ORDEN en el Sheet...")

    # Leer todo el Sheet JUNIO
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA_JUNIO}'!A:D"
    ).execute()
    filas = result.get("values", [])[1:]  # saltar header

    corregidas  = 0
    sin_mein    = []

    for idx, row in enumerate(filas):
        fila_num = idx + 2  # +2 por header y base 1
        mein  = row[0].strip() if len(row) > 0 and row[0] else ""
        orden = row[1].strip() if len(row) > 1 and row[1] else ""
        dom   = row[2].strip() if len(row) > 2 and row[2] else ""
        tipo  = row[3].strip() if len(row) > 3 and row[3] else ""

        if not mein or not orden:
            continue

        # Si mein = orden, buscar MEIN real en SAP
        if mein == orden:
            print(f"  [CORRIGIENDO] fila {fila_num} — Orden {orden} | {dom} | {tipo}")
            mein_real = buscar_mein_en_sap(orden)

            if mein_real and mein_real != orden:
                # Reemplazar col A con la MEIN real
                service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"'{HOJA_JUNIO}'!A{fila_num}",
                    valueInputOption="RAW",
                    body={"values": [[mein_real]]}
                ).execute()
                print(f"    → MEIN corregida: {mein_real}")
                corregidas += 1
            else:
                print(f"    → No se encontro MEIN en SAP")
                sin_mein.append({"fila": fila_num, "orden": orden, "dom": dom, "tipo": tipo})

    print(f"  Corregidas: {corregidas} | Sin MEIN en SAP: {len(sin_mein)}")
    if sin_mein:
        print("  PENDIENTES (sin MEIN en SAP):")
        for s in sin_mein:
            print(f"    fila {s['fila']} — Orden {s['orden']} | {s['dom']} | {s['tipo']}")

    # Validar y eliminar duplicados
    eliminar_duplicados(service)


def eliminar_duplicados(service):
    """
    Busca filas duplicadas en el Sheet JUNIO (mismo numero en col A)
    y elimina la segunda ocurrencia, avisando cuales se eliminaron.
    """
    print("\nValidando duplicados en el Sheet JUNIO...")

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA_JUNIO}'!A:D"
    ).execute()
    filas = result.get("values", [])[1:]

    vistos   = {}   # mein → primera fila donde aparece
    a_borrar = []   # filas a eliminar (en orden inverso para no desplazar)

    for idx, row in enumerate(filas):
        fila_num = idx + 2
        mein = row[0].strip() if len(row) > 0 and row[0] else ""
        if not mein:
            continue
        if mein in vistos:
            orden = row[1].strip() if len(row) > 1 else ""
            dom   = row[2].strip() if len(row) > 2 else ""
            tipo  = row[3].strip() if len(row) > 3 else ""
            print(f"  [DUPLICADO] MEIN {mein} en fila {fila_num} (primera vez en fila {vistos[mein]}) — {dom} | {tipo} → ELIMINANDO")
            a_borrar.append(fila_num)
        else:
            vistos[mein] = fila_num

    if not a_borrar:
        print("  Sin duplicados.")
        return

    # Eliminar en orden inverso para no desplazar indices
    for fila_num in sorted(a_borrar, reverse=True):
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{
                "deleteDimension": {
                    "range": {
                        "sheetId": obtener_sheet_id(service, HOJA_JUNIO),
                        "dimension": "ROWS",
                        "startIndex": fila_num - 1,
                        "endIndex": fila_num
                    }
                }
            }]}
        ).execute()

    print(f"  {len(a_borrar)} fila(s) duplicada(s) eliminada(s).")


def obtener_sheet_id(service, nombre_hoja):
    """Obtiene el sheetId de una hoja por nombre."""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == nombre_hoja:
            return s["properties"]["sheetId"]
    return 0


def registrar_meins_junio(service, meins):
    """
    Agrega TODAS las MEINs al Sheet JUNIO (incluso mein=orden).
    Despues llama a corregir_meins_iguales para:
      1. Buscar la MEIN real en SAP para las que tienen mein=orden
      2. Eliminar duplicados
    """
    if not meins:
        return 0

    # Leer col A y combo B+C+D para no duplicar al agregar
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA_JUNIO}'!A:D"
    ).execute()
    filas_existentes = result.get("values", [])[1:]

    meins_existentes  = set()
    combos_existentes = set()

    for row in filas_existentes:
        mein_s  = row[0].strip() if len(row) > 0 and row[0] else ""
        orden_s = row[1].strip() if len(row) > 1 and row[1] else ""
        dom_s   = row[2].strip() if len(row) > 2 and row[2] else ""
        tipo_s  = row[3].strip() if len(row) > 3 and row[3] else ""
        if mein_s:  meins_existentes.add(mein_s)
        if orden_s and dom_s and tipo_s:
            combos_existentes.add(f"{orden_s}|{dom_s}|{tipo_s}")

    nuevas   = []
    saltadas = 0

    for m in meins:
        mein      = str(m.get("mein", "")).strip()
        orden     = str(m.get("orden", "")).strip()
        domicilio = str(m.get("domicilio", "")).strip()
        tipo      = str(m.get("tipo", "")).strip()

        # No agregar si la MEIN ya existe en col A
        if mein in meins_existentes:
            saltadas += 1
            continue

        # No agregar si la combinacion orden+dom+tipo ya existe
        combo = f"{orden}|{domicilio}|{tipo}"
        if combo in combos_existentes:
            saltadas += 1
            continue

        nuevas.append([mein, orden, domicilio, tipo, "", "", "", ""])
        meins_existentes.add(mein)
        combos_existentes.add(combo)

    if nuevas:
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{HOJA_JUNIO}'!A:H",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": nuevas}
        ).execute()
        print(f"  [JUNIO] {len(nuevas)} MEIN(s) agregada(s) | {saltadas} saltada(s) por duplicado")
        print("  Ejecuta mepl_corregir_meins.py para corregir las que tienen MEIN = ORDEN")

    return len(nuevas)


def limpiar_mein_file():
    """Borra el contenido de mein_creada.txt después de procesarlo."""
    if MEIN_FILE.exists():
        MEIN_FILE.write_text("", encoding="utf-8")



def sincronizar_meins_desde_sheet(service):
    """
    Lee el Sheet JUNIO 2026, busca filas donde col A != col B
    (MEIN corregida por mepl_corregir_meins.py), y actualiza
    mein_creada en Supabase en un solo batch.
    """
    print("\nSincronizando MEINs corregidas del Sheet a Supabase...")

    # Leer col A y col B del Sheet
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{HOJA_JUNIO}'!A:B"
    ).execute()
    filas = result.get("values", [])[1:]  # saltar header

    # Construir mapa orden → mein_correcta solo donde A != B
    mapa_correcciones = {}
    for row in filas:
        mein_sheet  = row[0].strip() if len(row) > 0 and row[0] else ""
        orden_sheet = row[1].strip() if len(row) > 1 and row[1] else ""
        if not mein_sheet or not orden_sheet:
            continue
        if mein_sheet != orden_sheet:
            mapa_correcciones[orden_sheet] = mein_sheet

    if not mapa_correcciones:
        print("  No hay MEINs para sincronizar.")
        return

    print(f"  {len(mapa_correcciones)} MEINs corregidas en Sheet, verificando Supabase...")

    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }

    # Traer de Supabase todas las meins_pendientes donde mein_creada = orden
    # en una sola consulta con los órdenes del mapa
    ordenes_lista = list(mapa_correcciones.keys())
    # Consultar en batches de 100 para no superar límites de URL
    actualizadas = 0
    batch_size = 100
    for i in range(0, len(ordenes_lista), batch_size):
        batch = ordenes_lista[i:i+batch_size]
        ordenes_str = ",".join(batch)
        res = requests.get(
            f"{SUPABASE_URL}/rest/v1/meins_pendientes",
            headers=headers,
            params={
                "orden":   f"in.({ordenes_str})",
                "select":  "id,orden,mein_creada"
            }
        )
        if not res.ok:
            print(f"  [ERROR] No se pudieron leer meins_pendientes: {res.text[:100]}")
            continue

        registros = res.json()
        for reg in registros:
            orden = str(reg.get("orden", ""))
            mein_actual = str(reg.get("mein_creada", ""))
            mein_correcta = mapa_correcciones.get(orden, "")

            # Solo actualizar si mein_creada = orden (sin corregir)
            if not mein_correcta or mein_actual != orden:
                continue

            patch = requests.patch(
                f"{SUPABASE_URL}/rest/v1/meins_pendientes",
                headers={**headers, "Prefer": "return=minimal"},
                params={"id": f"eq.{reg['id']}"},
                json={"mein_creada": mein_correcta}
            )
            if patch.ok:
                print(f"  [MEIN ACTUALIZADA] Orden {orden} → MEIN {mein_correcta}")
                actualizadas += 1
            else:
                print(f"  [ERROR] Orden {orden}: {patch.text[:100]}")

    if actualizadas == 0:
        print("  No hay MEINs para sincronizar.")
    else:
        print(f"  {actualizadas} MEIN(s) actualizadas en Supabase.")


def main():
    print("=" * 45)
    print("   MEPL — Sincronizar Supabase → Sheets")
    print("=" * 45)
    print("")

    asegurar_librerias()

    # ── MEINs nuevas de JUNIO (desde Supabase) ────
    meins_nuevas = leer_meins_supabase()
    if meins_nuevas:
        print(f"MEINs nuevas en Supabase (no sincronizadas): {len(meins_nuevas)}")
        for m in meins_nuevas:
            print(f"  - MEIN: {m['mein']} | Orden: {m['orden']} | {m['domicilio']} | {m['tipo']}")
    else:
        print("No hay MEINs nuevas en Supabase")

    # Flujo legacy: leer txt si existe
    meins_txt = leer_mein_creadas()
    if meins_txt:
        print(f"MEINs en mein_creada.txt: {len(meins_txt)}")
        meins_nuevas = meins_nuevas + meins_txt

    # ── Respuestas de Supabase ─────────────────────
    print("\nLeyendo respuestas nuevas de Supabase...")
    filas = leer_supabase()
    print(f"Encontradas: {len(filas)} respuesta(s) nuevas")

    if not filas and not meins_nuevas:
        print("No hay nada para procesar.")
        input("\nPresione Enter para salir...")
        return

    if filas:
        for f in filas:
            print(f"  - MEIN: {f.get('mein')} | orden_ref: {f.get('orden_ref')} | {f.get('domicilio')} | {f.get('tipo')}")

    resp = input(f"\nProcesar todo? (S/N): ").strip().upper()
    if resp not in ("S", "SI"):
        print("Cancelado.")
        input("\nPresione Enter para salir...")
        return

    print("Conectando con Google Sheets...")
    service = crear_servicio_sheets()

    # Sincronizar MEINs corregidas del Sheet a Supabase
    sincronizar_meins_desde_sheet(service)

    # Registrar MEINs nuevas en JUNIO
    if meins_nuevas:
        registrar_meins_junio(service, meins_nuevas)
        # Marcar como sincronizadas en Supabase
        ids_supabase = [m["id"] for m in meins_nuevas if "id" in m]
        if ids_supabase:
            marcar_meins_sincronizadas(ids_supabase)
            print(f"  {len(ids_supabase)} MEIN(s) marcadas como sincronizadas en Supabase")
        limpiar_mein_file()

    # Volcar respuestas de Supabase en MAYO
    actualizadas, nuevas = escribir_sheets(service, filas)
    print(f"Actualizadas: {actualizadas} | Nuevas: {nuevas}")

    if filas:
        ids = [str(f["id"]) for f in filas]
        marcar_procesado(ids)
        print("Marcadas como procesadas en Supabase.")

    print("")
    print("=" * 45)
    print(f"Completado: {actualizadas} actualizada(s), {nuevas} nueva(s)")
    print("=" * 45)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
