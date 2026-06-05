# -*- coding: utf-8 -*-
"""
MEPL - Descarga y analisis de Excel IW39 con lenguaje natural
--------------------------------------------------------------
1. Ejecuta MEPL_IW39_DESCARGA.vbs y guarda el Excel
2. Carga el Excel en memoria
3. Permite consultas en lenguaje natural usando Claude API
4. Muestra resultados en consola y exporta a Excel
"""

import sys
import subprocess
import time
import json
from pathlib import Path
from datetime import datetime

# ========= CONFIGURACION =========
EXCEL_PATH    = Path(r"C:\Users\usuario\Desktop\FILTRO IW39\export.XLSX")
SCRIPT_IW39  = Path(r"C:\Users\usuario\AppData\Roaming\SAP\SAP GUI\Scripts\MEPL_IW39_DESCARGA.vbs")
RESULTADO_DIR = Path(r"C:\Users\usuario\Desktop\FILTRO IW39")
TIMEOUT_SEG  = 30
# =================================


def asegurar_librerias():
    pkgs = []
    try:
        import pandas as pd  # noqa
    except Exception:
        pkgs.append("pandas")
    try:
        import openpyxl  # noqa
    except Exception:
        pkgs.append("openpyxl")
    try:
        import google.generativeai  # noqa
    except Exception:
        pkgs.append("google-generativeai")
    if pkgs:
        print(f"Instalando: {', '.join(pkgs)}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *pkgs])


def ejecutar_vbs():
    if not SCRIPT_IW39.exists():
        raise FileNotFoundError(f"No encuentro el script:\n{SCRIPT_IW39}")
    print("Ejecutando IW39 en SAP...")
    subprocess.Popen(["cscript.exe", "//nologo", str(SCRIPT_IW39)])


def esperar_descarga():
    """El VBS guarda directamente en la carpeta - solo esperamos que aparezca."""
    print("SAP esta exportando el Excel...")
    print("Si aparece el popup 'Application not executable', hace click en OK manualmente.")
    time.sleep(3)



def esperar_archivo():
    ts_anterior = EXCEL_PATH.stat().st_mtime if EXCEL_PATH.exists() else 0
    print("Esperando archivo...")
    for _ in range(TIMEOUT_SEG * 2):
        time.sleep(0.5)
        if EXCEL_PATH.exists() and EXCEL_PATH.stat().st_mtime > ts_anterior:
            print("Archivo descargado.")
            return True
    print("ADVERTENCIA: no se detecto el archivo nuevo. Continuando con el existente.")
    return EXCEL_PATH.exists()


def cargar_excel():
    import pandas as pd
    print(f"Leyendo {EXCEL_PATH.name}...")
    df = pd.read_excel(EXCEL_PATH, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].fillna("").str.strip()
    print(f"Total registros: {len(df)}")
    return df


def construir_contexto(df):
    """Genera un resumen del DataFrame para pasarle a Claude."""
    import pandas as pd

    clases   = df["Clase de orden"].value_counts().to_dict()
    statuses = df["Status de usuario"].value_counts().to_dict()
    grupos   = df["Grupo planificación"].value_counts().to_dict() if "Grupo planificación" in df.columns else {}

    domicilios = df["Denominación de la ubicación técnica"].dropna().unique().tolist()
    domicilios = [d for d in domicilios if d][:50]  # maximo 50

    # Obtener muestra de fechas para mostrarle el formato real a la IA
    fechas_muestra = df["Fecha de inicio extrema"].dropna().unique()[:5].tolist()

    contexto = f"""
El DataFrame tiene {len(df)} filas con estas columnas:
{list(df.columns)}

Clases de orden: {json.dumps(clases, ensure_ascii=False)}
Status de usuario: {json.dumps(statuses, ensure_ascii=False)}
Grupos de planificacion: {json.dumps(grupos, ensure_ascii=False)}
Domicilios disponibles (muestra): {domicilios}

FECHAS IMPORTANTES:
- Las columnas "Fecha de inicio extrema" y "Fecha fin extrema" son strings con formato: {fechas_muestra}
- Para filtrar por mes/año USA str.contains() o str.startswith(), NO uses pd.to_datetime()
- Cuando el usuario dice "de mayo", "de abril", etc. se refiere a la columna "Fecha fin extrema" (fecha de vencimiento)
- El formato de fecha es D/M/YYYY sin ceros adelante. Ejemplos reales: "15/5/2026", "1/2/2026", "25/12/2025"
- Para filtrar por mes USA siempre este patron: str.contains("/5/2026") para mayo 2026, str.contains("/4/2026") para abril 2026
- NUNCA uses "/05/" ni "/04/" porque las fechas no tienen cero adelante
- NUNCA filtres por "Fecha de inicio extrema" cuando el usuario mencione un mes, siempre usa "Fecha fin extrema"
- Para buscar domicilios usa str.contains(texto, case=False, na=False)

IMPORTANTE: 
- Devuelve SOLO codigo Python valido, sin markdown, sin ```, sin explicaciones.
- La variable resultado debe ser un DataFrame.
- Para buscar texto en columnas usa str.contains(texto, case=False, na=False).
- Columnas disponibles: "Orden", "Orden principal", "Denominación de la ubicación técnica", "Texto breve", "Clase de orden", "Status de usuario", "Grupo planificación", "Fecha de inicio extrema", "Fecha fin extrema".
"""
    return contexto


def consultar_claude(consulta, contexto):
    import google.generativeai as genai
    import os

    genai.configure(api_key=os.environ["GEMINI_API_KEY"])
    model = genai.GenerativeModel("gemini-2.0-flash")

    prompt = f"""
{contexto}

El usuario quiere: "{consulta}"

REGLAS ESTRICTAS:
- Genera SOLO codigo Python que filtre el DataFrame `df` ya cargado en memoria
- Guarda el resultado en una variable llamada `resultado`
- NO uses pd.read_excel(), NO leas archivos, NO exportes nada
- NO uses to_excel(), NO guardes archivos
- Solo filtros sobre `df` con condiciones booleanas
- Sin markdown, sin ```, sin explicaciones, solo codigo Python puro
"""

    response = model.generate_content(prompt)
    codigo = response.text.strip()
    codigo = codigo.replace("```python", "").replace("```", "").strip()
    return codigo


def ejecutar_filtro(df, codigo):
    """Ejecuta el codigo generado por Claude sobre el DataFrame."""
    namespace = {"df": df, "resultado": None}
    try:
        exec(codigo, namespace)
        return namespace.get("resultado", None)
    except Exception as e:
        print(f"Error ejecutando filtro: {e}")
        print(f"Codigo generado:\n{codigo}")
        return None


def exportar_resultado(resultado, consulta):
    import pandas as pd
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"filtro_{ts}.xlsx"
    ruta = RESULTADO_DIR / nombre
    resultado.to_excel(ruta, index=False)
    print(f"Exportado: {ruta}")


def menu_consultas(df):
    contexto = construir_contexto(df)

    print("")
    print("=" * 55)
    print("Consultas en lenguaje natural")
    print("Escribi lo que queres buscar.")
    print("Ejemplo: MEIN de Gainza de abril")
    print("Ejemplo: MEPL sin MEIN creada status INIC")
    print("Escribi 0 para salir.")
    print("=" * 55)

    while True:
        print("")
        consulta = input("Consulta: ").strip()

        if consulta == "0" or consulta.upper() in ("FIN", "SALIR", "EXIT"):
            print("Saliendo...")
            break

        if not consulta:
            continue

        print("Interpretando consulta...")
        codigo = consultar_claude(consulta, contexto)

        resultado = ejecutar_filtro(df, codigo)

        if resultado is None or len(resultado) == 0:
            print("Sin resultados para esa consulta.")
            continue

        print(f"\n{len(resultado)} registros encontrados:\n")

        # Mostrar en consola
        import pandas as pd
        pd.set_option("display.max_rows", 200)
        pd.set_option("display.width", 180)
        pd.set_option("display.max_colwidth", 40)

        cols_mostrar = [c for c in [
            "Orden", "Orden principal",
            "Denominación de la ubicación técnica",
            "Texto breve", "Clase de orden",
            "Status de usuario", "Grupo planificación"
        ] if c in resultado.columns]

        print(resultado[cols_mostrar].to_string(index=False))

        # Exportar
        exportar = input("\nExportar a Excel? (S/N): ").strip().upper()
        if exportar in ("S", "SI"):
            exportar_resultado(resultado, consulta)


def main():
    print("=" * 55)
    print("   MEPL - IW39 Descarga y Analisis")
    print("=" * 55)
    print("")

    try:
        asegurar_librerias()

        descargar = input("Descargar Excel de SAP ahora? (S/N): ").strip().upper()
        if descargar in ("S", "SI"):
            ejecutar_vbs()
            time.sleep(4)
            esperar_descarga()
            esperar_archivo()
        else:
            if not EXCEL_PATH.exists():
                raise FileNotFoundError(f"No existe el archivo:\n{EXCEL_PATH}")
            print(f"Usando archivo existente: {EXCEL_PATH.name}")

        df = cargar_excel()
        menu_consultas(df)

    except Exception as e:
        print("")
        print("ERROR:")
        print(str(e))

    input("\nPresione Enter para salir...")


if __name__ == "__main__":
    main()
