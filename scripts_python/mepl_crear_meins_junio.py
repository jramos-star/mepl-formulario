# -*- coding: utf-8 -*-
"""
mepl_crear_meins_junio.py
--------------------------
Lee las MEINs pendientes de Supabase (tabla meins_pendientes),
ejecuta el VBS para crearlas en SAP, y registra el resultado
de vuelta en Supabase.

El panel admin del formulario web es quien encola las ordenes
en meins_pendientes. Este script solo las procesa.

Flujo:
  1. Leer meins_pendientes donde estado = 'pendiente'
  2. Marcar cada una como 'procesando'
  3. Ejecutar VBS → capturar MEIN creada del stdout
  4. Actualizar estado a 'creada' (con mein_creada) o 'error'
  5. supabase_sync.py lee meins_pendientes y vuelca al Sheet
"""

import subprocess
import requests
import re
from pathlib import Path
from datetime import datetime, timezone

# ========= CONFIGURACION =========
SUPABASE_URL = "https://zynvptlesftdpyijxvpv.supabase.co"
SUPABASE_KEY = "sb_publishable_a5BtY5K2O8w2rPBlIeKRUg_u6z478MX"

VBS_PATH = r"C:\Users\usuario\AppData\Roaming\SAP\SAP GUI\Scripts\MEPL_01_CREAR_MEIN.vbs"

# Filtrar solo un mes, o None para procesar todos los pendientes
MES_FILTRO = None   # Ej: "Junio" para procesar solo Junio
# =================================

HEADERS = {
    "apikey":        SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type":  "application/json",
    "Prefer":        "return=representation"
}


def sb_get(tabla, params):
    url = f"{SUPABASE_URL}/rest/v1/{tabla}"
    res = requests.get(url, headers=HEADERS, params=params)
    res.raise_for_status()
    return res.json()


def sb_patch(tabla, params, body):
    url = f"{SUPABASE_URL}/rest/v1/{tabla}"
    res = requests.patch(url, headers=HEADERS, params=params, json=body)
    res.raise_for_status()
    return res.json()


def sb_post(tabla, body):
    url = f"{SUPABASE_URL}/rest/v1/{tabla}"
    res = requests.post(url, headers=HEADERS, json=body)
    res.raise_for_status()
    return res.json()


def leer_pendientes():
    params = {"estado": "eq.pendiente", "select": "*"}
    if MES_FILTRO:
        params["mes"] = f"eq.{MES_FILTRO}"
    return sb_get("meins_pendientes", params)


def marcar_procesando(id_):
    sb_patch("meins_pendientes",
             {"id": f"eq.{id_}"},
             {"estado": "procesando"})


def marcar_creada(id_, mein):
    sb_patch("meins_pendientes",
             {"id": f"eq.{id_}"},
             {"estado": "creada",
              "mein_creada": mein,
              "procesado_en": datetime.now(timezone.utc).isoformat()})


def marcar_error(id_, msg):
    sb_patch("meins_pendientes",
             {"id": f"eq.{id_}"},
             {"estado": "error",
              "error_msg": msg,
              "procesado_en": datetime.now(timezone.utc).isoformat()})





def extraer_mein_del_stdout(stdout):
    """
    El VBS de creacion de MEIN debe imprimir algo como:
      OK | MEIN: 421503210
    o simplemente el numero de MEIN en el stdout.
    Intentamos extraer el numero.
    """
    # Buscar patron "MEIN: XXXXXXXXX"
    match = re.search(r'MEIN[:\s]+(\d{6,12})', stdout, re.IGNORECASE)
    if match:
        return match.group(1)
    # Buscar patron "OK | XXXXXXXXX"
    match = re.search(r'OK\s*\|\s*(\d{6,12})', stdout, re.IGNORECASE)
    if match:
        return match.group(1)
    # Buscar cualquier numero largo en el stdout
    match = re.search(r'\b(\d{9,12})\b', stdout)
    if match:
        return match.group(1)
    return None


def main():
    print("=" * 55)
    print("   MEPL — Crear MEINs desde Supabase")
    print("=" * 55)
    print("")

    print("Leyendo pendientes de Supabase...")
    try:
        pendientes = leer_pendientes()
    except Exception as e:
        print(f"ERROR al conectar con Supabase: {e}")
        input("\nPresione Enter para salir...")
        return

    if not pendientes:
        print("No hay MEINs pendientes de crear.")
        input("\nPresione Enter para salir...")
        return

    print(f"  {len(pendientes)} MEIN(s) pendiente(s):")
    for o in pendientes:
        print(f"  - [{o['mes']}] Orden {o['orden']} | {o['domicilio']} | {o['tipo']}")

    resp = input(f"\nCrear {len(pendientes)} MEIN(s) en SAP? (S/N): ").strip().upper()
    if resp not in ("S", "SI"):
        print("Cancelado.")
        input("\nPresione Enter para salir...")
        return

    print("\nCreando MEINs en SAP...")
    creadas = 0
    errores = 0

    for i, o in enumerate(pendientes, 1):
        id_   = o["id"]
        orden = o["orden"]
        tipo  = o["tipo"]
        dom   = o["domicilio"]
        mes   = o["mes"]

        print(f"\n  [{i}/{len(pendientes)}] Orden {orden} — {dom} — {tipo}")

        marcar_procesando(id_)

        try:
            resultado = subprocess.run(
                ["cscript", "//NoLogo", VBS_PATH, orden, tipo],
                capture_output=True, text=True, timeout=60
            )

            stdout = resultado.stdout.strip()
            stderr = resultado.stderr.strip()

            if stdout:
                print(f"    SAP: {stdout}")
            if stderr:
                print(f"    ERR: {stderr}")

            if resultado.returncode != 0:
                msg = stderr or stdout or f"returncode {resultado.returncode}"
                marcar_error(id_, msg)
                print(f"    → ERROR")
                errores += 1
                continue

            # Extraer numero de MEIN del stdout
            mein = extraer_mein_del_stdout(stdout)
            if not mein:
                # Si el VBS no imprime el numero, usar la orden como fallback
                print(f"    ADVERTENCIA: no se pudo extraer numero de MEIN del stdout")
                print(f"    Revisá el VBS MEPL_01_CREAR_MEIN.vbs — debe imprimir 'MEIN: XXXXXXXXX'")
                mein = orden  # fallback

            marcar_creada(id_, mein)
            print(f"    → MEIN creada: {mein}")
            creadas += 1

        except subprocess.TimeoutExpired:
            marcar_error(id_, "timeout al ejecutar VBS")
            print(f"    → TIMEOUT")
            errores += 1
        except Exception as e:
            marcar_error(id_, str(e))
            print(f"    → ERROR: {e}")
            errores += 1

    print("")
    print("=" * 55)
    print(f"Completado: {creadas} creada(s) | {errores} error(es)")
    print("Las MEINs fueron registradas en Supabase.")
    print("Ejecuta supabase_sync.py para volcarlas al Sheet.")
    print("=" * 55)

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
