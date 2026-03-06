"""
=============================================================
  EXTRACTOR DE DATOS - API SOLENIUM (MÓDULO 1)
=============================================================
Extrae datos de 4 endpoints de la API de Solenium:
  - /api/project/                        → Proyectos Sunfactory (listado base)
  - /api/solarverse/project_management/  → Proyectos con gestión y CREG
  - /api/attributes/company/             → Inversionistas
  - /api/solarverse/sunfactory/          → Proyectos Sunfactory importables

Al finalizar guarda los resultados en:
  ../../data/proyectos_solenium.json

REQUISITOS:
    pip install requests python-dotenv

CONFIGURACIÓN:
    Crea un archivo .env en la misma carpeta con:
        SOLENIUM_TOKEN=tu_token_aqui
        SOLENIUM_BASE_URL=https://dev-api-sunfactory.solenium.co

CÓMO USAR:
    python extraer_datos_api.py
=============================================================
"""

import json
import logging
import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv


# =============================================================
#  LOGGING
# =============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# =============================================================
#  CONFIGURACIÓN DESDE .env
# =============================================================
load_dotenv()

TOKEN    = os.getenv("SOLENIUM_TOKEN", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0b2tlbl90eXBlIjoiYWNjZXNzIiwiZXhwIjoxNzcyODk4Nzc4LCJpYXQiOjE3NzI4MTIzNzgsImp0aSI6IjBhOGVhNDI2ZTNlNTQ3MmRhODUyNzZlZDBlZGZmYTIyIiwidXNlcl9pZCI6IjkifQ.f0rbFSTYyBqozyHzUMKcfoW21YTeCfnwIrHEM_qviwE")
BASE_URL = os.getenv("SOLENIUM_BASE_URL", "https://dev-api-sunfactory.solenium.co/").rstrip("/")

# Comportamiento de las peticiones
TIMEOUT_SEGUNDOS = 15
MAX_REINTENTOS   = 3
PAUSA_REINTENTOS = 2
PAGE_SIZE        = 50

# =============================================================
#  EMPRESAS HARDCODEADAS
#  La API de dev no expone estas empresas. Sincronizar contra
#  producción (sunfactory.solenium.co) cuando sea posible.
# =============================================================
EMPRESAS_HARDCODEADAS = {
    24: {"id": 24, "nombre": "FMO",    "nit": None, "codigo_categoria": None, "logo_url": None},
    90: {"id": 90, "nombre": "Ayurá",  "nit": None, "codigo_categoria": None, "logo_url": None},
}

# Correcciones puntuales de campos incorrectos en la API de dev.
# Formato: { project_id: { campo: valor_correcto } }
CORRECCIONES_PROYECTOS = {
    34: {"operador_red_nombre": "Afinia", "operador_red_id": None, "operador_red_nit": None},
}


# =============================================================
#  CAPA DE TRANSPORTE
# =============================================================

def crear_headers() -> dict:
    """Construye los headers de autenticación con el Bearer token."""
    if not TOKEN:
        raise ValueError(
            "SOLENIUM_TOKEN no encontrado. "
            "Agrégalo al archivo .env: SOLENIUM_TOKEN=tu_token_aqui"
        )
    return {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json",
    }


def hacer_get(url: str, params: dict | None = None) -> dict | list | None:
    """
    Realiza una petición GET con timeout, manejo de errores y reintentos.
    Retorna el JSON parseado, o None si todos los intentos fallaron.
    """
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            respuesta = requests.get(
                url,
                headers=crear_headers(),
                params=params,
                timeout=TIMEOUT_SEGUNDOS,
            )
            respuesta.raise_for_status()
            return respuesta.json()

        except requests.exceptions.Timeout:
            logger.warning("Timeout — intento %d/%d: %s", intento, MAX_REINTENTOS, url)

        except requests.exceptions.ConnectionError:
            logger.warning("Sin conexión — intento %d/%d: %s", intento, MAX_REINTENTOS, url)

        except requests.exceptions.HTTPError as e:
            logger.error("Error HTTP %s en: %s", e.response.status_code, url)
            logger.error("Detalle: %s", e.response.text[:300])
            return None  # Errores 4xx no tiene sentido reintentar

        except requests.exceptions.JSONDecodeError:
            logger.error("Respuesta no es JSON válido: %s", url)
            return None

        if intento < MAX_REINTENTOS:
            time.sleep(PAUSA_REINTENTOS)

    logger.error("Se agotaron %d intentos para: %s", MAX_REINTENTOS, url)
    return None


def obtener_todas_las_paginas(endpoint: str, params: dict | None = None) -> list:
    """
    Recorre automáticamente todas las páginas de un endpoint paginado.
    La API devuelve: { count, next, previous, results: [...] }

    endpoint → ruta relativa, ej: "/api/project/"
    params   → parámetros adicionales de filtro (opcionales)
    Retorna la lista completa de resultados de todas las páginas.
    """
    url = f"{BASE_URL}{endpoint}"
    parametros = {"page_size": PAGE_SIZE, **(params or {})}
    todos = []
    pagina = 1

    while url:
        logger.info("  Página %d: %s", pagina, url)
        # Página 1: enviamos params explícitos
        # Páginas siguientes: la URL de "next" ya los trae incluidos
        datos = hacer_get(url, params=parametros if pagina == 1 else None)

        if datos is None:
            logger.error("  Fallo al obtener página %d de %s", pagina, endpoint)
            break

        if isinstance(datos, list):
            todos.extend(datos)
            break

        if isinstance(datos, dict):
            todos.extend(datos.get("results", []))
            url = datos.get("next")  # None = última página
            pagina += 1
        else:
            logger.error("  Formato inesperado: %s", type(datos))
            break

    return todos


# =============================================================
#  EXTRACCIÓN POR ENDPOINT
# =============================================================

def extraer_proyectos_base() -> list[dict]:
    """
    Extrae el listado base de proyectos desde /api/project/.
    Contiene campos técnicos: planta, cuenta analítica, minifarm, tracker, etc.
    """
    logger.info("── Extrayendo proyectos base (/api/project/)...")
    raw = obtener_todas_las_paginas("/api/project/")
    logger.info("   %d proyectos base obtenidos.", len(raw))

    return [
        {
            "id":                     item.get("id"),
            "nombre":                 item.get("name"),
            "nombre_base":            item.get("base_name"),
            "descripcion":            item.get("description"),
            "ciudad":                 item.get("city"),
            "departamento":           item.get("department"),
            "direccion":              item.get("address"),
            "latitud":                item.get("lat"),
            "longitud":               item.get("lon"),
            "codigo_planta":          item.get("plant_code"),
            "cuenta_analitica":       item.get("analytical_account"),
            "cuenta_analitica_odoo":  item.get("odoo_analytical_account"),
            "despliegue_interno":     item.get("internal_deployment"),
            "fase_ejecucion":         item.get("execution_phase"),
            "estado_codigo":          item.get("state"),
            "estado":                 item.get("state_description"),
            "potencia_dc_kw":         item.get("total_dc_power"),
            "potencia_ac_kw":         item.get("total_ac_power"),
            "es_minifarm":            item.get("is_minifarm"),
            "es_tracker":             item.get("is_tracker"),
        }
        for item in raw
        if item
    ]


def extraer_proyectos_gestion() -> list[dict]:
    """
    Extrae proyectos con datos de gestión desde /api/solarverse/project_management/.
    Contiene: CREG 174, CREG 075, FPO, operador de red, inversionistas, avance.
    """
    logger.info("── Extrayendo proyectos gestión (/api/solarverse/project_management/)...")
    raw = obtener_todas_las_paginas("/api/solarverse/project_management/")
    logger.info("   %d proyectos de gestión obtenidos.", len(raw))

    proyectos = []
    for item in raw:
        or_obj      = item.get("grid_operator") or {}
        creg075_obj = item.get("creg_075_status") or {}

        proyectos.append({
            "id":                            item.get("id"),
            "nombre":                        item.get("name"),
            "ciudad":                        item.get("city"),
            "departamento":                  item.get("department"),
            "latitud":                       item.get("lat"),
            "longitud":                      item.get("lon"),
            "potencia_dc_kw":                item.get("total_dc_power"),
            "potencia_ac_kw":                item.get("total_ac_power"),
            "estado":                        item.get("state_description"),
            "fase_ejecucion":                item.get("execution_phase_description"),
            "porcentaje_avance":             item.get("project_percentage"),
            "fecha_inicio":                  item.get("start_date"),
            "fpo_unergy":                    item.get("fpo_unergy"),
            "fpo_contrato":                  item.get("fpo_contract"),
            "fpo_cronograma":                item.get("fpo_schedule"),
            "creg_174_estado":               item.get("creg_174_status"),
            "creg_174_retraso":              item.get("creg_174_delay"),
            "creg_174_fecha_vencimiento":    item.get("creg_174_expiration_date"),
            "creg_075_estado":               creg075_obj.get("state"),
            "creg_075_descripcion":          creg075_obj.get("description"),
            "creg_075_ultima_actualizacion": item.get("creg_075_last_update"),
            "fecha_visita_punto_conexion":   item.get("connection_point_visit_date"),
            "procedimientos":                item.get("procedures"),
            "operador_red_id":               or_obj.get("id"),
            "operador_red_nombre":           or_obj.get("name"),
            "operador_red_nit":              or_obj.get("nit"),
            # IDs temporales — enriquecidos luego con datos completos
            "_inversionistas_ids":           item.get("investors") or [],
        })

    return proyectos



def extraer_inversionistas() -> list[dict]:
    """
    Extrae empresas inversionistas desde /api/attributes/company/.
    """
    logger.info("── Extrayendo inversionistas (/api/attributes/company/)...")
    raw = obtener_todas_las_paginas(
        "/api/attributes/company/",
        params={"type": "investor", "include_images": "true"},
    )
    logger.info("   %d empresas obtenidas del catálogo.", len(raw))

    return [
        {
            "id":               item.get("id"),
            "nombre":           item.get("name"),
            "nit":              item.get("nit"),
            "codigo_categoria": item.get("category_code"),
            "logo_url":         item.get("logo"),
        }
        for item in raw
    ]


def extraer_sunfactory() -> list[dict]:
    """
    Extrae proyectos importables desde /api/solarverse/sunfactory/.
    """
    logger.info("── Extrayendo proyectos Sunfactory (/api/solarverse/sunfactory/)...")
    raw = obtener_todas_las_paginas("/api/solarverse/sunfactory/")
    logger.info("   %d proyectos Sunfactory obtenidos.", len(raw))

    proyectos = []
    for item in raw:
        or_obj = item.get("grid_operator") or {}
        inversores = [
            {"id": inv.get("id"), "nombre": inv.get("name"), "nit": inv.get("nit")}
            for inv in (item.get("investors") or [])
        ]
        proyectos.append({
            "id":                         item.get("id"),
            "nombre":                     item.get("name"),
            "codigo_proyecto":            item.get("project_code"),
            "ciudad":                     item.get("city"),
            "departamento":               item.get("department"),
            "latitud":                    item.get("lat"),
            "longitud":                   item.get("lon"),
            "fpo_unergy":                 item.get("fpo_unergy"),
            "creg_174_estado":            item.get("creg_174_status"),
            "creg_174_retraso":           item.get("creg_174_delay"),
            "creg_174_fecha_vencimiento": item.get("creg_174_expiration_date"),
            "operador_red_id":            or_obj.get("id"),
            "operador_red_nombre":        or_obj.get("name"),
            "operador_red_nit":           or_obj.get("nit"),
            "inversionistas":             inversores,
        })

    return proyectos


# =============================================================
#  COMBINACIÓN Y ENRIQUECIMIENTO
# =============================================================

def enriquecer_con_inversionistas(
    proyectos: list[dict],
    inversionistas: list[dict],
) -> list[dict]:
    """
    Reemplaza los IDs de inversionistas en cada proyecto con sus datos completos.
    """
    indice = {inv["id"]: inv for inv in inversionistas}
    for proyecto in proyectos:
        ids = proyecto.pop("_inversionistas_ids", [])
        proyecto["inversionistas"] = [indice[i] for i in ids if i in indice]
    return proyectos


def combinar_proyectos(
    proyectos_base: list[dict],
    proyectos_gestion: list[dict],
) -> list[dict]:
    """
    Une los datos de /api/project/ y /api/solarverse/project_management/
    usando el ID como clave. Los campos de gestión enriquecen los datos base.
    Si un proyecto solo existe en una fuente, se incluye igualmente.
    """
    indice_gestion = {p["id"]: p for p in proyectos_gestion}
    combinados = []

    for base in proyectos_base:
        pid = base["id"]
        gestion = indice_gestion.pop(pid, {})
        # La base aporta campos técnicos; gestión aporta CREG, FPO, operador, etc.
        # En caso de conflicto de campo, gestión tiene precedencia (datos más ricos)
        combinado = {**base, **{k: v for k, v in gestion.items() if v is not None}}
        if pid in CORRECCIONES_PROYECTOS:
            combinado.update(CORRECCIONES_PROYECTOS[pid])
        combinados.append(combinado)

    # Proyectos que solo están en gestión (no en base)
    for pid, gestion in indice_gestion.items():
        logger.warning("Proyecto ID %s solo existe en project_management, no en /api/project/", pid)
        combinados.append(gestion)

    return combinados


# =============================================================
#  FUNCIÓN PRINCIPAL
# =============================================================

def main() -> dict:
    """
    Orquesta la extracción completa de todas las fuentes.
    Retorna un diccionario unificado listo para el Módulo 2
    y guarda los resultados en un archivo JSON.
    """
    if not BASE_URL:
        logger.error("SOLENIUM_BASE_URL no encontrado en .env. Abortando.")
        return {}

    logger.info("=" * 55)
    logger.info("  EXTRACTOR DE DATOS SOLENIUM — MÓDULO 1")
    logger.info("=" * 55)

    errores = []

    # 1. Catálogo de empresas desde /api/attributes/company/
    try:
        inversionistas = extraer_inversionistas()
    except Exception as e:
        logger.error("Error extrayendo inversionistas: %s", e)
        inversionistas = []
        errores.append(f"inversionistas: {e}")

    # 2. Proyectos base (campos técnicos)
    try:
        proyectos_base = extraer_proyectos_base()
    except Exception as e:
        logger.error("Error extrayendo proyectos base: %s", e)
        proyectos_base = []
        errores.append(f"proyectos_base: {e}")

    # 3. Proyectos Sunfactory — se extrae antes del enriquecimiento
    #    porque sus objetos de inversionistas completan el catálogo
    try:
        sunfactory = extraer_sunfactory()
    except Exception as e:
        logger.error("Error extrayendo Sunfactory: %s", e)
        sunfactory = []
        errores.append(f"sunfactory: {e}")

    # Catálogo: companies de la API + hardcodeadas + las de sunfactory
    catalogo = {inv["id"]: inv for inv in inversionistas}
    for cid, empresa in EMPRESAS_HARDCODEADAS.items():
        if cid not in catalogo:
            catalogo[cid] = empresa
    for sf in sunfactory:
        for inv in sf.get("inversionistas") or []:
            if inv.get("id") and inv["id"] not in catalogo:
                catalogo[inv["id"]] = {
                    "id":               inv["id"],
                    "nombre":           inv.get("nombre"),
                    "nit":              inv.get("nit"),
                    "codigo_categoria": None,
                    "logo_url":         None,
                }
    # 4. Proyectos gestión (CREG, FPO, operador de red)
    try:
        proyectos_gestion = extraer_proyectos_gestion()

        inversionistas_completos = list(catalogo.values())
        logger.info("Catálogo de empresas: %d", len(inversionistas_completos))

        proyectos_gestion = enriquecer_con_inversionistas(proyectos_gestion, inversionistas_completos)
    except Exception as e:
        logger.error("Error extrayendo proyectos gestión: %s", e)
        proyectos_gestion = []
        errores.append(f"proyectos_gestion: {e}")

    # 5. Combinar base + gestión en un solo objeto por proyecto
    try:
        proyectos = combinar_proyectos(proyectos_base, proyectos_gestion)
    except Exception as e:
        logger.error("Error combinando proyectos: %s", e)
        proyectos = proyectos_base or proyectos_gestion
        errores.append(f"combinacion: {e}")

    # --- Resultado unificado ---
    resultado = {
        "proyectos":      proyectos,
        "inversionistas": inversionistas_completos,
        "sunfactory":     sunfactory,
        "errores":        errores,
        "resumen": {
            "total_proyectos":      len(proyectos),
            "total_inversionistas": len(inversionistas),
            "total_sunfactory":     len(sunfactory),
            "total_errores":        len(errores),
        },
    }

    # --- Resumen en consola ---
    logger.info("=" * 55)
    logger.info("EXTRACCIÓN COMPLETA")
    logger.info("  Proyectos      : %d", len(proyectos))
    logger.info("  Inversionistas : %d", len(inversionistas_completos))
    logger.info("  Sunfactory     : %d", len(sunfactory))
    if errores:
        logger.warning("  Errores        : %d", len(errores))
        for e in errores:
            logger.warning("    - %s", e)
    logger.info("=" * 55)

    # Vista previa — El Copey + primeros 2 con inversionistas
    muestra = [p for p in proyectos if "copey" in (p.get("nombre") or "").lower()][:1]
    muestra += [p for p in proyectos if p.get("inversionistas")][:2]
    logger.info("Vista previa (El Copey + proyectos con inversionistas):")
    for p in muestra:
        invs = ", ".join(i["nombre"] for i in p.get("inversionistas", [])) or "—"
        logger.info(
            "  [%s] %s | %s, %s | DC: %s kW | Estado: %s | Planta: %s | Inversionistas: %s",
            p.get("id"),
            p.get("nombre") or "—",
            p.get("ciudad") or "—",
            p.get("departamento") or "—",
            p.get("potencia_dc_kw") or "—",
            p.get("estado") or "—",
            p.get("codigo_planta") or "—",
            invs,
        )

    return resultado


# =============================================================
#  PUNTO DE ENTRADA
# =============================================================
if __name__ == "__main__":
    datos = main()

    if datos.get("proyectos"):
        ruta_json = Path(__file__).parent.parent / "data" / "proyectos_solenium.json"
        ruta_json.parent.mkdir(parents=True, exist_ok=True)

        with open(ruta_json, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)

        logger.info("💾 JSON guardado en: %s", ruta_json)

    # datos["proyectos"]      → proyectos combinados y enriquecidos
    # datos["inversionistas"] → lista de empresas inversionistas
    # datos["sunfactory"]     → proyectos disponibles en Sunfactory
    # datos["errores"]        → errores no críticos ocurridos
    # datos["resumen"]        → conteos totales
    #
    # Ejemplo para exportar a Excel:
    #   import pandas as pd
    #   pd.DataFrame(datos["proyectos"]).to_excel("proyectos.xlsx", index=False)