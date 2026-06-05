# -*- coding: utf-8 -*-
"""
MEPL - Crear Google Forms por colegio
--------------------------------------
Lee el Sheets MEPL_MAYO_2026_FORMULARIO y crea un
Google Form por cada colegio con las tareas del mes.
Las respuestas se vinculan al Sheets MEPL_RESPUESTAS_MAYO.
"""

import sys
import subprocess
import json
from pathlib import Path

# ========= CONFIGURACION =========
FORMULARIO_ID   = "1WTVIO58edDpTumUcmhYhxcMBhYFcG77LjRKRCAOkUIA"  # MEPL_MAYO_2026_FORMULARIO
RESPUESTAS_ID   = "1J15vkpPBd3T3ud1b5CFcVx2YBjMcU0hP68CIpOW_mgQ"  # MEPL_RESPUESTAS_MAYO
HOJA_FORMULARIO = "Sheet1"
HOJA_RESPUESTAS = "MEPL_RESPUESTAS"

BASE_DIR         = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token_forms.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/forms.body",
    "https://www.googleapis.com/auth/drive",
]

OPCIONES_RAPIDAS = ["No aplica", "En obra"]
# =================================


def asegurar_librerias():
    try:
        import google.auth  # noqa
        from googleapiclient.discovery import build  # noqa
    except Exception:
        print("Instalando librerias...")
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "google-auth", "google-auth-oauthlib",
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

    sheets = build("sheets", "v4", credentials=creds)
    forms  = build("forms",  "v1", credentials=creds)
    drive  = build("drive",  "v3", credentials=creds)
    return sheets, forms, drive, creds


def leer_formulario(sheets):
    rango  = f"'{HOJA_FORMULARIO}'!A:J"
    result = sheets.spreadsheets().values().get(
        spreadsheetId=FORMULARIO_ID,
        range=rango
    ).execute()
    filas = result.get("values", [])

    registros = []
    for i, row in enumerate(filas[1:], start=2):  # saltar header
        if len(row) < 5:
            continue
        registros.append({
            "fila":        i,
            "sap_nro":     row[0] if len(row) > 0 else "",
            "domicilio":   row[1] if len(row) > 1 else "",
            "fecha_inicio":row[2] if len(row) > 2 else "",
            "fecha_fin":   row[3] if len(row) > 3 else "",
            "clave":       row[4] if len(row) > 4 else "",
            "nombre":      row[5] if len(row) > 5 else "",
            "descripcion": row[6] if len(row) > 6 else "",
        })
    return registros


def agrupar_por_domicilio(registros):
    grupos = {}
    for r in registros:
        dom = r["domicilio"]
        if dom not in grupos:
            grupos[dom] = []
        grupos[dom].append(r)
    return grupos


def crear_form_colegio(forms, drive, domicilio, tareas):
    """Crea un Google Form para un colegio con sus tareas."""

    # Crear el form base
    form_body = {
        "info": {
            "title": f"MEPL — {domicilio}",
            "documentTitle": f"MEPL_{domicilio[:30]}"
        }
    }
    form = forms.forms().create(body=form_body).execute()
    form_id = form["formId"]

    # Construir los items del formulario
    requests = []
    index = 0

    # Titulo y descripcion del formulario
    requests.append({
        "updateFormInfo": {
            "info": {
                "description": f"Completá las respuestas de tus tareas mensuales.\nColegio: {domicilio}"
            },
            "updateMask": "description"
        }
    })

    # Agrupar tareas por orden SAP
    ordenes = {}
    for t in tareas:
        sap = t["sap_nro"]
        if sap not in ordenes:
            ordenes[sap] = []
        ordenes[sap].append(t)

    for sap_nro, items in ordenes.items():
        # Titulo de sección por orden
        fecha = items[0]["fecha_inicio"] + " → " + items[0]["fecha_fin"] if items else ""
        requests.append({
            "createItem": {
                "item": {
                    "title": f"Orden SAP: {sap_nro}",
                    "description": fecha,
                    "pageBreakItem": {}
                },
                "location": {"index": index}
            }
        })
        index += 1

        # Una pregunta por tarea
        for t in items:
            titulo_pregunta = f"{t['clave']} — {t['nombre']}"
            desc_pregunta   = t["descripcion"][:200] if t["descripcion"] else ""

            requests.append({
                "createItem": {
                    "item": {
                        "title": titulo_pregunta,
                        "description": desc_pregunta,
                        "questionItem": {
                            "question": {
                                "required": True,
                                "textQuestion": {
                                    "paragraph": True
                                }
                            }
                        }
                    },
                    "location": {"index": index}
                }
            })
            index += 1

    # Agregar campo de nombre del operario al final
    requests.append({
        "createItem": {
            "item": {
                "title": "Nombre del operario",
                "questionItem": {
                    "question": {
                        "required": True,
                        "textQuestion": {
                            "paragraph": False
                        }
                    }
                }
            },
            "location": {"index": index}
        }
    })
    index += 1

    # Aplicar todos los cambios
    forms.forms().batchUpdate(
        formId=form_id,
        body={"requests": requests}
    ).execute()

    # Hacer el form público (cualquier persona puede acceder)
    drive.permissions().create(
        fileId=form_id,
        body={"role": "reader", "type": "anyone"}
    ).execute()

    form_url = f"https://docs.google.com/forms/d/{form_id}/viewform"
    return form_id, form_url


def guardar_links(links):
    """Guarda los links de los formularios en un archivo de texto."""
    archivo = BASE_DIR / "MEPL_FORMS_LINKS.txt"
    with archivo.open("w", encoding="utf-8") as f:
        f.write("MEPL — Links de Formularios por Colegio\n")
        f.write("=" * 50 + "\n\n")
        for domicilio, url in links.items():
            f.write(f"{domicilio}\n{url}\n\n")
    print(f"\nLinks guardados en: {archivo}")


def main():
    print("=" * 50)
    print("   MEPL — Crear Google Forms por Colegio")
    print("=" * 50)
    print("")

    try:
        asegurar_librerias()
        sheets, forms, drive, creds = crear_servicio()

        print("Leyendo datos del Sheets...")
        registros = leer_formulario(sheets)
        grupos    = agrupar_por_domicilio(registros)

        print(f"Colegios encontrados: {len(grupos)}")
        for dom in sorted(grupos.keys()):
            print(f"  - {dom} ({len(grupos[dom])} tareas)")

        print("")
        confirmar = input("Crear formularios? (S/N): ").strip().upper()
        if confirmar not in ("S", "SI"):
            print("Cancelado.")
            return

        links = {}
        for domicilio, tareas in sorted(grupos.items()):
            print(f"Creando form para: {domicilio}...")
            form_id, form_url = crear_form_colegio(forms, drive, domicilio, tareas)
            links[domicilio] = form_url
            print(f"  OK — {form_url}")

        guardar_links(links)

        print("")
        print("=" * 50)
        print(f"Formularios creados: {len(links)}")
        print("=" * 50)

    except Exception as e:
        print("")
        print("ERROR:")
        print(str(e))

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
