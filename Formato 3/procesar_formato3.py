"""
Script para actualizar el FORMATO 3 de "El Papá de los formatos" cruzando datos
con PEPC y BOM por Odoo ID.

Reglas:
  - Filas de F3 cuyo Odoo ID no aparezca ni en PEPC (Código Odoo) ni en BOM (ID)
    se eliminan.
  - Si coincide con BOM: Cantidad <- CANTIDAD de BOM.
  - Si coincide con PEPC: Valor total en COP (Sin incluir IVA) <- PRECIO TOTAL
    (convertido desde USD a COP usando la TRM tomada de D7 cuando MONEDA == USD).
  - Valor IVA en COP = 19% del Valor total, EXCEPTO para paneles/módulos
    fotovoltaicos e inversores/microinversores donde es 0.
  - Si un Odoo ID está duplicado en PEPC o BOM, se toma el valor MÁS ALTO y se
    advierte.
  - Alerta de coincidencias parciales (ID en F3 que matchea solo con PEPC o
    solo con BOM, no con ambos).
  - Se sobreescribe la hoja FORMATO 3 conservando el encabezado y formato
    original; solo se reemplaza el contenido bajo el header.

Uso:
    python procesar_formato3.py
"""

import re
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from rapidfuzz import fuzz, process

# ---------- Configuración ----------
PATH_PAPA = Path("C:/Users/usuario/Downloads/El_Papá_de_los_formatos_CON_ODOO.xlsx")
PATH_PEPC = Path("C:/Users/usuario/Downloads/PEPC_1P_GENERICO_NOCFM.xlsx")
PATH_BOM  = Path("C:/Users/usuario/Downloads/Bill of Materials Inv 249kW-1P.xlsx")

PATH_SALIDA = Path("C:/Users/usuario/Downloads/El_Papá_de_los_formatos_ACTUALIZADO.xlsx")

# Encabezados (1-indexed como los ve un humano en Excel)
HEADER_ROW_PEPC = 12   # -> pandas header=11
HEADER_ROW_F3   = 11   # -> pandas header=10
HEADER_ROW_BOM  = 1    # -> pandas header=0

COL_PEPC_ID       = "Código Odoo"
COL_PEPC_PRECIO   = "PRECIO TOTAL"
COL_PEPC_MONEDA   = "MONEDA"

COL_BOM_ID        = "ID"
COL_BOM_CANTIDAD  = "CANTIDAD"

# Nombres exentos de IVA (se ponen en 0)
NOMBRES_IVA_CERO = {
    "Paneles/modulos o celdas fotovoltaicas",
    "Inversores o microinversores (Off Gid, Grid Tie o Híbrido)",
}

# Valores que NO son IDs reales y deben ignorarse
IDS_INVALIDOS = {
    "", "0", "no encontrado", "no creado en odoo", "no se encuentra item", "nan",
}

IVA_RATE = 0.19


# ---------- Helpers ----------
def encontrar_columna(df, fragmento: str) -> str:
    """Busca la columna cuyo nombre contenga el fragmento dado (ignora mayúsculas y saltos de línea)."""
    fragmento_norm = fragmento.lower().replace("\n", "").replace(" ", "")
    for col in df.columns:
        col_norm = str(col).lower().replace("\n", "").replace(" ", "")
        if fragmento_norm in col_norm:
            return col
    raise KeyError(f"No se encontró columna con fragmento: {fragmento!r}. Columnas disponibles: {list(df.columns)}")


def limpiar_id(valor) -> str | None:
    """Normaliza un valor de ID: strip + lowercase para comparación de invalidez.
    Devuelve el ID limpio (string en mayúsculas) o None si no es válido."""
    if pd.isna(valor):
        return None
    s = str(valor).strip()
    # IDs con comas (ej. 'P00684, P00684') o múltiples tokens los descartamos
    # porque no son un Odoo ID limpio.
    if "," in s:
        return None
    if s.lower() in IDS_INVALIDOS:
        return None
    # Algunos IDs vienen con tabs/espacios internos — limpiarlos
    s = re.sub(r"/s+", "", s)
    if s.lower() in IDS_INVALIDOS:
        return None
    return s.upper()


def extraer_trm(valor_d7) -> float:
    """Extrae el número de TRM desde el string en D7.
    Soporta formatos: 'TRM: $3700', 'TRM: $3.700', 'TRM: 3,700', 'TRM: $4.200,50', '3700', etc."""
    if valor_d7 is None:
        raise ValueError("Celda D7 vacía: no se puede determinar la TRM.")
    if isinstance(valor_d7, (int, float)):
        return float(valor_d7)
    s = str(valor_d7).replace("$", "").replace(" ", "")
    m = re.search(r"([\d][0-9.,]*)", s)
    if not m:
        raise ValueError(f"No se pudo extraer la TRM de D7: {valor_d7!r}")
    raw = m.group(1)
    raw_clean = re.sub(r"[.,](?=\d{3}(?:[.,]|$))", "", raw)
    raw_clean = raw_clean.replace(",", ".")
    return float(raw_clean)


def reducir_max(df, col_id, col_valor, label):
    """Agrupa por ID tomando el valor máximo. Devuelve dict {id: max_value} y
    una lista de tuplas (id, [valores]) para los IDs con duplicados, para alertar."""
    sub = df[[col_id, col_valor]].dropna(subset=[col_id])
    sub = sub[sub[col_valor].notna()]
    # Convertir col_valor a numérico (puede haber strings)
    sub = sub.copy()
    sub[col_valor] = pd.to_numeric(sub[col_valor], errors="coerce")
    sub = sub.dropna(subset=[col_valor])

    duplicados = []
    grupos = sub.groupby(col_id)[col_valor]
    for id_, valores in grupos:
        vals = valores.tolist()
        if len(vals) > 1 and len(set(vals)) > 1:
            # Solo alertamos si los valores no son todos iguales
            duplicados.append((id_, vals))
    return grupos.max().to_dict(), duplicados


# ---------- 1. Cargar dataframes ----------
def cargar_dataframes():
    df_pepc = pd.read_excel(PATH_PEPC, sheet_name="PEPC", header=HEADER_ROW_PEPC - 1)
    df_f3   = pd.read_excel(PATH_PAPA, sheet_name="FORMATO 3", header=HEADER_ROW_F3 - 1)
    df_bom  = pd.read_excel(PATH_BOM,  sheet_name="BOM", header=HEADER_ROW_BOM - 1)
    return df_pepc, df_f3, df_bom


# ---------- 2. Obtener TRM desde D7 de PEPC ----------
def obtener_trm() -> float:
    wb = load_workbook(PATH_PEPC, data_only=True)
    try:
        return extraer_trm(wb["PEPC"]["D7"].value)
    finally:
        wb.close()


# ---------- 3. Procesar el dataframe Formato 3 ----------
def procesar(df_pepc, df_f3, df_bom, trm):
    # Normalizar IDs
    df_f3 = df_f3.copy()
    COL_F3_VALOR_TOT = encontrar_columna(df_f3, "Valor total en COP")
    COL_F3_IVA       = encontrar_columna(df_f3, "Valor IVA en COP")
    COL_F3_CANTIDAD  = encontrar_columna(df_f3, "Cantidad")
    COL_F3_NOMBRE    = encontrar_columna(df_f3, "Nombre del Elemento")
    COL_F3_ID        = encontrar_columna(df_f3, "Odoo ID")
    df_f3["_id_norm"] = df_f3[COL_F3_ID].apply(limpiar_id)

    df_pepc = df_pepc.copy()
    df_pepc["_id_norm"] = df_pepc[COL_PEPC_ID].apply(limpiar_id)

    df_bom = df_bom.copy()
    df_bom["_id_norm"] = df_bom[COL_BOM_ID].apply(limpiar_id)

    # Convertir PRECIO TOTAL de PEPC a COP por fila (según MONEDA)
    def convertir_a_cop(row):
        precio = pd.to_numeric(row[COL_PEPC_PRECIO], errors="coerce")
        if pd.isna(precio):
            return None
        moneda = str(row[COL_PEPC_MONEDA]).strip().upper() if pd.notna(row[COL_PEPC_MONEDA]) else ""
        if moneda == "USD":
            return precio * trm
        # COP o cualquier otra cosa -> se deja tal cual
        return precio

    df_pepc["_precio_cop"] = df_pepc.apply(convertir_a_cop, axis=1)

    # Set de IDs que aparecen en F3 (para filtrar PEPC/BOM antes de reducir)
    ids_f3 = set(df_f3["_id_norm"].dropna().unique())

    # Reducir PEPC y BOM tomando el valor MÁS ALTO en caso de duplicados,
    # pero solo entre IDs que aparezcan en F3.
    pepc_filtrado = df_pepc[df_pepc["_id_norm"].isin(ids_f3)]
    bom_filtrado  = df_bom[df_bom["_id_norm"].isin(ids_f3)]

    precios_pepc, dup_pepc = reducir_max(pepc_filtrado, "_id_norm", "_precio_cop", "PEPC")
    cantidades_bom, dup_bom = reducir_max(bom_filtrado,  "_id_norm", COL_BOM_CANTIDAD, "BOM")

    ids_pepc = set(precios_pepc.keys())
    ids_bom  = set(cantidades_bom.keys())

    # Alertas de duplicados con valores distintos
    if dup_pepc:
        print("/n⚠ ALERTA: IDs con múltiples PRECIO TOTAL distintos en PEPC")
        print("  (se toma el valor MÁS ALTO):")
        tabla = pd.DataFrame(
            [(i, vals, max(vals)) for i, vals in dup_pepc],
            columns=["Odoo ID", "Valores encontrados (COP)", "Valor usado (máx)"],
        )
        print(tabla.to_string(index=False))

    if dup_bom:
        print("/n⚠ ALERTA: IDs con múltiples CANTIDAD distintas en BOM")
        print("  (se toma el valor MÁS ALTO):")
        tabla = pd.DataFrame(
            [(i, vals, max(vals)) for i, vals in dup_bom],
            columns=["Odoo ID", "Cantidades encontradas", "Cantidad usada (máx)"],
        )
        print(tabla.to_string(index=False))

    # Filtrar F3: solo filas cuyo ID esté en PEPC o BOM
    en_alguno = df_f3["_id_norm"].isin(ids_pepc | ids_bom) & df_f3["_id_norm"].notna()
    eliminadas = df_f3[~en_alguno]
    df_f3 = df_f3[en_alguno].copy()

    if len(eliminadas):
        print(f"/nℹ Se eliminaron {len(eliminadas)} filas de FORMATO 3 sin coincidencia "
              f"de Odoo ID en PEPC ni BOM.")

    # Aplicar reemplazos
    # Si el ID no aparece en BOM -> cantidad = 0 (no conservar valor original)
    df_f3[COL_F3_CANTIDAD] = df_f3.apply(
        lambda r: cantidades_bom.get(r["_id_norm"], 0),
        axis=1,
    )
    # Si el ID no aparece en PEPC -> valor total = 0 (no conservar valor original)
    df_f3[COL_F3_VALOR_TOT] = df_f3.apply(
        lambda r: precios_pepc.get(r["_id_norm"], 0),
        axis=1,
    )

    # Recalcular IVA (19% del valor total, 0 para paneles e inversores)
    def calcular_iva(row):
        nombre = str(row[COL_F3_NOMBRE]).strip() if pd.notna(row[COL_F3_NOMBRE]) else ""
        if nombre in NOMBRES_IVA_CERO:
            return 0
        valor = pd.to_numeric(row[COL_F3_VALOR_TOT], errors="coerce")
        if pd.isna(valor):
            return None
        return valor * IVA_RATE

    df_f3[COL_F3_IVA] = df_f3.apply(calcular_iva, axis=1)

    # ---------- Alerta de coincidencias parciales ----------
    # IDs de F3 (después del filtrado) que matchean solo con uno de los dos
    ids_f3_final = set(df_f3["_id_norm"].dropna().unique())
    solo_pepc = (ids_f3_final & ids_pepc) - ids_bom
    solo_bom  = (ids_f3_final & ids_bom) - ids_pepc

    if solo_pepc or solo_bom:
        print("/n⚠ ALERTA: Coincidencias parciales (Odoo ID encontrado solo en uno "
              "de los dos archivos de referencia):")
        filas = []
        for id_ in sorted(solo_pepc):
            filas.append((id_, "✓", "✗"))
        for id_ in sorted(solo_bom):
            filas.append((id_, "✗", "✓"))
        tabla = pd.DataFrame(filas, columns=["Odoo ID", "En PEPC", "En BOM"])
        print(tabla.to_string(index=False))
        print(f"/n  Total parciales: {len(filas)} "
              f"(solo PEPC: {len(solo_pepc)}, solo BOM: {len(solo_bom)})")
    else:
        print("/n✓ Todos los Odoo ID coinciden en ambos archivos (PEPC y BOM).")

    # Limpiar columna auxiliar
    df_f3 = df_f3.drop(columns=["_id_norm"])
    return df_f3


# ---------- 4. Escribir el resultado preservando el formato original ----------
def escribir_resultado(df_f3_final):
    """Sobrescribe la hoja FORMATO 3 conservando el header y formato del original.
    Borra todas las filas debajo del header y escribe las nuevas."""
    PATH_SALIDA.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(PATH_PAPA, PATH_SALIDA)

    wb = load_workbook(PATH_SALIDA)
    ws = wb["FORMATO 3"]

    # Borrar filas debajo del header (header en fila 11, datos desde fila 12)
    primera_fila_datos = HEADER_ROW_F3 + 1
    if ws.max_row >= primera_fila_datos:
        ws.delete_rows(primera_fila_datos, ws.max_row - primera_fila_datos + 1)

    # Escribir nuevas filas. Orden de columnas según los headers en la hoja.
    # Leer la fila del header desde la hoja para saber el orden exacto.
    headers_hoja = [
        ws.cell(row=HEADER_ROW_F3, column=c).value
        for c in range(1, ws.max_column + 1)
    ]

    # Reindexar el dataframe al orden de la hoja (columnas no presentes -> None)
    df_orden = df_f3_final.reindex(columns=headers_hoja)

    for i, (_, row) in enumerate(df_orden.iterrows()):
        excel_row = primera_fila_datos + i
        for j, val in enumerate(row, start=1):
            if pd.isna(val):
                val = None
            ws.cell(row=excel_row, column=j, value=val)

    borde_grueso = Border(
        left=Side(style="medium"),
        right=Side(style="medium"),
        top=Side(style="medium"),
        bottom=Side(style="medium"),
    )

    ultima_fila = primera_fila_datos + len(df_orden) - 1
    for fila in range(HEADER_ROW_F3, ultima_fila + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row=fila, column=col).border = borde_grueso

    wb.save(PATH_SALIDA)
    wb.close()


def _tiene_columna_odoo(df: pd.DataFrame) -> bool:
    """Devuelve True si el dataframe ya tiene una columna 'Código Odoo' con al
    menos un valor válido (no NaN, no inválido)."""
    try:
        col = encontrar_columna(df, "Código Odoo")
    except KeyError:
        return False
    return df[col].apply(limpiar_id).notna().any()


def completar_codigo_odoo(
    path_pepc_objetivo: Path,
    path_pepc_referencia: Path,
    path_bom: Path,
    path_papa: Path,
    path_salida_pepc: Path,
    umbral_similitud: float = 50.0,
):
    """
    Versión flexible: recibe cualquier PEPC (con o sin 'Código Odoo') y lo
    completa fila a fila.

    Lógica por fila:
      - Si la fila ya tiene un Código Odoo válido  →  se conserva sin tocar.
      - Si la fila NO tiene Código Odoo (o está vacía/inválida) → se intenta
        asignar uno en este orden de prioridad:
          1. Coincidencia exacta de DESCRIPCIÓN vs PEPC de referencia.
          2. Fuzzy match de DESCRIPCIÓN vs MATERIAL de BOM.
          3. Fuzzy match de DESCRIPCIÓN vs Modelo/Referencia del Papá (FORMATO 3).
          4. Fuzzy match de DESCRIPCIÓN vs DESCRIPCIÓN del PEPC de referencia.
        Si ninguna fuente supera el umbral, la celda queda vacía y se reporta
        en la tabla de alertas.

    Al terminar imprime:
      - Cuántas filas ya tenían código (conservadas).
      - Cuántas filas se completaron (y desde qué fuente).
      - Tabla de alertas con las filas que no pudieron resolverse.
    """
    print("\n=== completar_codigo_odoo ===")

    # --- Cargar archivos ---
    df_obj = pd.read_excel(path_pepc_objetivo,   sheet_name="PEPC", header=HEADER_ROW_PEPC - 1)
    df_ref = pd.read_excel(path_pepc_referencia, sheet_name="PEPC", header=HEADER_ROW_PEPC - 1)
    df_bom = pd.read_excel(path_bom,             sheet_name="BOM",  header=HEADER_ROW_BOM - 1)
    df_f3  = pd.read_excel(path_papa,            sheet_name="FORMATO 3", header=HEADER_ROW_F3 - 1)

    # --- Detectar columnas ---
    col_desc_obj  = encontrar_columna(df_obj, "DESCRIPCIÓN")
    col_desc_ref  = encontrar_columna(df_ref, "DESCRIPCIÓN")
    col_odoo_ref  = encontrar_columna(df_ref, "Código Odoo")
    col_bom_mat   = encontrar_columna(df_bom, "MATERIAL")
    col_bom_id    = encontrar_columna(df_bom, "ID")
    col_f3_modelo = encontrar_columna(df_f3,  "Modelo / Referencia")
    col_f3_odoo   = encontrar_columna(df_f3,  "Odoo ID")

    # Detectar si el PEPC objetivo ya tiene la columna Código Odoo
    try:
        col_odoo_obj = encontrar_columna(df_obj, "Código Odoo")
        col_existe = True
    except KeyError:
        col_odoo_obj = "Código Odoo"
        df_obj[col_odoo_obj] = None
        col_existe = False

    ya_tenian = df_obj[col_odoo_obj].apply(limpiar_id).notna().sum()
    modo = "parcialmente completado" if (col_existe and ya_tenian > 0) else \
           "sin columna Código Odoo" if not col_existe else "con columna vacía"
    print(f"  PEPC objetivo ({modo}): {len(df_obj)} filas, {ya_tenian} ya con código")
    print(f"  PEPC referencia:        {len(df_ref)} filas")
    print(f"  BOM:                    {len(df_bom)} filas")
    print(f"  FORMATO 3:              {len(df_f3)} filas")

    # --- Construir lookups desde las fuentes de referencia ---
    lookup_ref = {}       # descripción exacta -> (desc_original, código)
    for _, row in df_ref.iterrows():
        desc = str(row[col_desc_ref]).strip() if pd.notna(row[col_desc_ref]) else ""
        cod  = limpiar_id(row[col_odoo_ref])
        if desc and cod:
            lookup_ref[desc.lower()] = (desc, cod)

    lookup_bom = {}       # material -> (material_original, código)
    for _, row in df_bom.iterrows():
        mat = str(row[col_bom_mat]).strip() if pd.notna(row[col_bom_mat]) else ""
        cod = limpiar_id(row[col_bom_id])
        if mat and cod:
            lookup_bom[mat.lower()] = (mat, cod)

    lookup_f3 = {}        # modelo/referencia -> (modelo_original, código)
    for _, row in df_f3.iterrows():
        mod = str(row[col_f3_modelo]).strip() if pd.notna(row[col_f3_modelo]) else ""
        cod = limpiar_id(row[col_f3_odoo])
        if mod and cod:
            lookup_f3[mod.lower()] = (mod, cod)

    claves_ref = list(lookup_ref.keys())
    claves_bom = list(lookup_bom.keys())
    claves_f3  = list(lookup_f3.keys())

    # --- Procesar fila a fila ---
    codigos_finales = []
    alertas = []
    stats = {"conservadas": 0, "exactas": 0, "bom": 0, "f3": 0, "fuzzy_ref": 0, "sin_match": 0}

    for idx, row in df_obj.iterrows():
        # Si ya tiene un código válido, conservarlo
        cod_actual = limpiar_id(row[col_odoo_obj])
        if cod_actual:
            codigos_finales.append(cod_actual)
            stats["conservadas"] += 1
            continue

        desc_raw  = str(row[col_desc_obj]).strip() if pd.notna(row[col_desc_obj]) else ""
        desc_norm = desc_raw.lower()

        if not desc_norm:
            codigos_finales.append(None)
            stats["sin_match"] += 1
            continue

        # Paso 1: coincidencia exacta con PEPC de referencia
        if desc_norm in lookup_ref:
            _, cod = lookup_ref[desc_norm]
            codigos_finales.append(cod)
            stats["exactas"] += 1
            continue

        # Paso 2: fuzzy matching en orden de prioridad
        codigo_asignado = None
        score_max = 0
        fuente_match = ""

        if claves_bom:
            m = process.extractOne(desc_norm, claves_bom,
                                   scorer=fuzz.token_set_ratio,
                                   score_cutoff=umbral_similitud)
            if m and m[1] > score_max:
                score_max = m[1]
                _, codigo_asignado = lookup_bom[m[0]]
                fuente_match = f"BOM/MATERIAL (score={m[1]:.0f}%)"

        if claves_f3:
            m = process.extractOne(desc_norm, claves_f3,
                                   scorer=fuzz.token_set_ratio,
                                   score_cutoff=umbral_similitud)
            if m and m[1] > score_max:
                score_max = m[1]
                _, codigo_asignado = lookup_f3[m[0]]
                fuente_match = f"FORMATO3/Modelo (score={m[1]:.0f}%)"

        if claves_ref:
            m = process.extractOne(desc_norm, claves_ref,
                                   scorer=fuzz.token_set_ratio,
                                   score_cutoff=umbral_similitud)
            if m and m[1] > score_max:
                score_max = m[1]
                _, codigo_asignado = lookup_ref[m[0]]
                fuente_match = f"PEPC-ref/DESCRIPCIÓN-fuzzy (score={m[1]:.0f}%)"

        if codigo_asignado:
            codigos_finales.append(codigo_asignado)
            key = fuente_match.split("/")[0].lower()
            if "bom" in key:
                stats["bom"] += 1
            elif "formato" in key:
                stats["f3"] += 1
            else:
                stats["fuzzy_ref"] += 1
        else:
            codigos_finales.append(None)
            stats["sin_match"] += 1
            alertas.append({
                "Fila Excel": idx + HEADER_ROW_PEPC + 1,
                "DESCRIPCIÓN": desc_raw,
                "Mejor score": f"{score_max:.0f}% (umbral: {umbral_similitud:.0f}%)",
            })

    # --- Resumen ---
    print(f"\n  Resumen:")
    print(f"    Conservadas (ya tenían código):  {stats['conservadas']}")
    print(f"    Asignadas por coincidencia exacta: {stats['exactas']}")
    print(f"    Asignadas por BOM/MATERIAL:        {stats['bom']}")
    print(f"    Asignadas por FORMATO3/Modelo:     {stats['f3']}")
    print(f"    Asignadas por fuzzy PEPC-ref:      {stats['fuzzy_ref']}")
    print(f"    Sin match (quedan vacías):         {stats['sin_match']}")

    if alertas:
        print(f"\n⚠ ALERTA: {len(alertas)} filas sin Código Odoo asignado:")
        print(pd.DataFrame(alertas).to_string(index=False))
    else:
        print("\n✓ Todas las filas sin código previo obtuvieron un Código Odoo.")

    # --- Escribir resultado ---
    df_obj[col_odoo_obj] = codigos_finales

    path_salida_pepc.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(path_pepc_objetivo, path_salida_pepc)

    wb = load_workbook(path_salida_pepc)
    ws = wb["PEPC"]

    # Localizar o crear la columna 'Código Odoo' en el header de la hoja
    header_row = HEADER_ROW_PEPC
    col_insertar = None
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=c).value
        if val and "código odoo" in str(val).lower():
            col_insertar = c
            break
    if col_insertar is None:
        # No existe la columna: buscar primera celda vacía en el header
        for c in range(1, ws.max_column + 2):
            if ws.cell(row=header_row, column=c).value is None:
                col_insertar = c
                break
        ws.cell(row=header_row, column=col_insertar, value="Código Odoo")

    primera_fila_datos = header_row + 1
    for i, cod in enumerate(codigos_finales):
        ws.cell(row=primera_fila_datos + i, column=col_insertar, value=cod)

    wb.save(path_salida_pepc)
    wb.close()
    print(f"\n✓ PEPC actualizado guardado en: {path_salida_pepc}")


def main():
    print("Cargando archivos...")
    df_pepc, df_f3, df_bom = cargar_dataframes()
    print(f"  PEPC:      {df_pepc.shape[0]} filas")
    print(f"  FORMATO 3: {df_f3.shape[0]} filas")
    print(f"  BOM:       {df_bom.shape[0]} filas")

    trm = obtener_trm()
    print(f"/nTRM detectada en D7 de PEPC: {trm:,.2f}")

    df_final = procesar(df_pepc, df_f3, df_bom, trm)
    print(f"/nFORMATO 3 final: {df_final.shape[0]} filas")

    escribir_resultado(df_final)
    print(f"/n✓ Archivo guardado en: {PATH_SALIDA}")

    # --- USO OPCIONAL: completar Código Odoo en cualquier PEPC ---
    # Funciona tanto si el PEPC no tiene la columna, como si la tiene parcialmente.
    # Las filas que ya tienen un código válido se conservan intactas.
    # Descomenta y ajusta las rutas para usar esta función.
    #
    # completar_codigo_odoo(
    #     path_pepc_objetivo    = Path("ruta/al/PEPC_a_completar.xlsx"),
    #     path_pepc_referencia  = Path("ruta/al/PEPC_con_odoo.xlsx"),  # fuente de verdad
    #     path_bom              = Path("ruta/al/BOM.xlsx"),
    #     path_papa             = Path("ruta/al/El_Papá_de_los_formatos.xlsx"),
    #     path_salida_pepc      = Path("ruta/salida/PEPC_completado.xlsx"),
    #     umbral_similitud      = 50.0,  # sube si hay falsos positivos, baja si hay pocos matches
    # )


if __name__ == "__main__":
    main()
