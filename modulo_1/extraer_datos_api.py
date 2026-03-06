"""
=============================================================
  EXTRACTOR DE PROYECTOS - PLATAFORMA SOLENIUM
=============================================================
Este script se conecta a la API de Solenium y extrae la
información de cada proyecto solar: nombre, ubicación,
fechas, potencia, inversionista, NIT y operador de red.

REQUISITOS (instalar una sola vez desde la terminal):
    pip install requests

CÓMO USAR:
    1. Rellena las variables en la sección CONFIGURACIÓN
    2. Ejecuta el script: python extraer_proyectos_solenium.py
    3. Los datos quedarán en la variable `proyectos` lista
       para que los proceses como necesites.
=============================================================
"""

import requests  # Para hacer llamadas a la API (páginas web)
import json      # Para manejar el formato de respuesta (JSON)
from pathlib import Path


# =============================================================
#  CONFIGURACIÓN  ← Solo debes editar esta sección
# =============================================================

# Tu token de acceso (el que ya tienes)
TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0b2tlbl90eXBlIjoiYWNjZXNzIiwiZXhwIjoxNzcyODk4Nzc4LCJpYXQiOjE3NzI4MTIzNzgsImp0aSI6IjBhOGVhNDI2ZTNlNTQ3MmRhODUyNzZlZDBlZGZmYTIyIiwidXNlcl9pZCI6IjkifQ.f0rbFSTYyBqozyHzUMKcfoW21YTeCfnwIrHEM_qviwE"

# URL base de la plataforma (sin barra al final)
# Ej: "https://app.solenium.co" o la URL que uses en Postman
BASE_URL = "https://dev-api-sunfactory.solenium.co/"

# =============================================================


def crear_headers():
    """
    Crea los encabezados de autenticación que necesita la API.
    Es como mostrar tu credencial antes de entrar a un edificio.
    """
    return {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json"
    }


def obtener_lista_proyectos():
    """
    Obtiene todos los proyectos recorriendo todas las páginas de la API.
    La API devuelve un objeto paginado con 'results', 'next' y 'count'.
    """
    print("📡 Conectando a la API para obtener la lista de proyectos...")
    todos = []
    url = f"{BASE_URL}/api/project/"

    while url:
        respuesta = requests.get(url, headers=crear_headers())
        if respuesta.status_code != 200:
            print(f"❌ Error al obtener proyectos: {respuesta.status_code}")
            print(f"   Detalle: {respuesta.text}")
            break
        data = respuesta.json()
        resultados = data.get("results") or []
        todos.extend(resultados)
        url = data.get("next")  # None cuando no hay más páginas

    print(f"✅ Se encontraron {len(todos)} proyectos.")
    return todos


def extraer_campos(item):
    """
    Extrae los campos disponibles de cada proyecto.
    Los campos provienen directamente del listado de la API
    (el endpoint de detalle solo devuelve URLs de actividades).
    """
    if not item:
        return None

    return {
        "id":                       item.get("id"),
        "nombre":                   item.get("name"),
        "nombre_base":              item.get("base_name"),
        "descripcion":              item.get("description"),
        "ciudad":                   item.get("city"),
        "departamento":             item.get("department"),
        "direccion":                item.get("address"),
        "latitud":                  item.get("lat"),
        "longitud":                 item.get("lon"),
        "codigo_planta":            item.get("plant_code"),
        "cuenta_analitica":         item.get("analytical_account"),
        "cuenta_analitica_odoo":    item.get("odoo_analytical_account"),
        "despliegue_interno":       item.get("internal_deployment"),
        "fase_ejecucion":           item.get("execution_phase"),
        "estado_codigo":            item.get("state"),
        "estado":                   item.get("state_description"),
        "potencia_dc_kw":           item.get("total_dc_power"),
        "potencia_ac_kw":           item.get("total_ac_power"),
        "es_minifarm":              item.get("is_minifarm"),
        "es_tracker":               item.get("is_tracker"),
    }


def main():
    """
    Función principal: orquesta todos los pasos.
    """
    print("\n" + "="*55)
    print("  EXTRACTOR DE PROYECTOS SOLENIUM")
    print("="*55 + "\n")

    # --- PASO 1: Obtener lista de proyectos ---
    lista = obtener_lista_proyectos()

    if not lista:
        print("\n🚫 No se obtuvieron proyectos. Verifica tu TOKEN y BASE_URL.")
        return []

    # --- PASO 2: Extraer campos de cada item de la lista ---
    print(f"\n📋 Procesando {len(lista)} proyectos...\n")
    proyectos = [extraer_campos(item) for item in lista if item]

    # --- RESULTADO ---
    print(f"\n{'='*55}")
    print(f"✅ Extracción completa: {len(proyectos)} proyectos procesados")
    print(f"{'='*55}\n")

    # Muestra un resumen en pantalla
    for i, p in enumerate(proyectos, 1):
        print(f"  {i}. {p['nombre']} (ID: {p['id']})")
        print(f"     📍 {p['ciudad']}, {p['departamento']} | Dir: {p['direccion']}")
        print(f"     ⚡ DC: {p['potencia_dc_kw']} kW | AC: {p['potencia_ac_kw']} kW")
        print(f"     🏭 Código planta  : {p['codigo_planta']}")
        print(f"     🔌 Estado         : {p['estado']} (cod: {p['estado_codigo']})")
        print(f"     🔧 Fase ejecución : {p['fase_ejecucion']}")
        print(f"     📌 Coords         : {p['latitud']}, {p['longitud']}")
        print(f"     🌱 Minifarm: {p['es_minifarm']} | Tracker: {p['es_tracker']}")
        print()

    return proyectos



# Punto de entrada del script
if __name__ == "__main__":
    proyectos = main()

    if proyectos:
        ruta_json = Path(__file__).parent.parent / "data" / "proyectos_solenium.json"
        ruta_json.parent.mkdir(parents=True, exist_ok=True)

        with open(ruta_json, "w", encoding="utf-8") as f:
            json.dump(proyectos, f, indent=2, ensure_ascii=False)

        print(f"💾 JSON guardado en: {ruta_json}")