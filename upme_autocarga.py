import os
import re
import math
import pandas as pd
from playwright.sync_api import sync_playwright

# ---------------- CONFIGURACIÓN ---------------- #
URL = "https://automatizacion-upme.bizagi.com/"
USUARIO = "1102867750"
PASSWORD = "S0L3N1UM"
RADICADO = "FNCE_202612071"

EXCEL_FILE = r"C:\Users\rober\Desktop\Automatismo\El COPEY.xlsx"
SHEET = "FORMATO 3"
PDF_DIR = r"C:\Users\rober\Desktop\Automatismo\Fichastecnicas"

# IVA fijo (ajústalo aquí si aplica otro)
IVA_FIJO = "19"

# Selector del "+" obtenido del Chrome Recorder
ADD_PLUS_XPATH = '//*[@id="mp_IncentivosFNCE_idmInformacionsolicitud_xEquipos"]/div/div[2]/div[3]/table/tbody/tr/th[1]/div/div/ul/li[1]/div'

# Nuevo selector para el clip de adjunto (Soporte) proporcionado por el usuario
ATTACH_CLIP_XPATH = 'span.ui-icon.upload-file'


# ---------------- UTILIDADES ---------------- #
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    for c in df.columns:
        c2 = str(c).replace("\n", " ").replace("\r", " ")
        c2 = re.sub(r"\s+", " ", c2).strip()
        cols.append(c2)
    df.columns = cols
    return df


def s(x):
    if x is None or (isinstance(x, float) and math.isnan(x)) or pd.isna(x):
        return ""
    return str(x).strip()


def fmt_number(x):
    if x is None or pd.isna(x):
        return ""
    if isinstance(x, int):
        return str(x)
    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        return str(x)
    return str(x).strip()


def fmt_money(x):
    if x is None or pd.isna(x):
        return ""
    try:
        # Si ya es un número, redondear y convertir a entero
        if isinstance(x, (int, float)):
            return str(int(round(x)))
            
        # Si es string, eliminar $, espacios y separadores comunes
        st = str(x).replace("$", "").replace(" ", "").strip()
        
        # Si tiene puntos y comas (formato contable), nos quedamos con la parte entera
        if "," in st and "." in st:
            if st.rfind(",") > st.rfind("."):
                st = st.split(",")[0].replace(".", "")
            else:
                st = st.split(".")[0].replace(",", "")
        elif "," in st and len(st) - st.rfind(",") == 3:
            st = st.split(",")[0].replace(".", "")
        elif "." in st and len(st) - st.rfind(".") == 3:
            st = st.split(".")[0].replace(",", "")
            
        # Limpieza final: dejar solo dígitos
        res = re.sub(r"\D", "", st)
        return res
    except:
        return ""




def set_zoom(page, zoom=0.8):
    page.add_style_tag(content=f"html, body {{ zoom: {zoom}; }}")
    page.wait_for_timeout(250)


def do_login(page):
    page.wait_for_timeout(1200)

    # Password (evita strict mode)
    if page.locator("#password").count() > 0:
        pw = page.locator("#password")
    elif page.locator("input[name='password']").count() > 0:
        pw = page.locator("input[name='password']")
    else:
        pw = page.locator("input[type='password']").first

    # Usuario
    if page.locator("#username").count() > 0:
        us = page.locator("#username")
    elif page.locator("input[name='username']").count() > 0:
        us = page.locator("input[name='username']")
    else:
        us = page.locator("input[type='text']").first

    us.wait_for(state="visible", timeout=20000)
    pw.wait_for(state="visible", timeout=20000)

    us.click()
    us.fill("")
    us.type(USUARIO, delay=35)

    pw.click()
    pw.fill("")
    pw.type(PASSWORD, delay=35)

    # Ingresar
    if page.locator("button[type='submit']").count() > 0:
        page.locator("button[type='submit']").first.click()
    elif page.get_by_role("button", name="Ingresar").count() > 0:
        page.get_by_role("button", name="Ingresar").click()
    else:
        pw.press("Enter")


def open_inbox(page):
    if page.locator("text=Inbox").count() > 0:
        page.locator("text=Inbox").first.click()
    elif page.locator("text=Bandeja de entrada").count() > 0:
        page.locator("text=Bandeja de entrada").first.click()
    page.wait_for_timeout(2000)


def open_case_by_radicado(page, radicado: str):
    if page.locator("input[placeholder*='Buscar']").count() > 0:
        search = page.locator("input[placeholder*='Buscar']").first
    else:
        search = page.locator("input").first

    search.click()
    search.fill(radicado)
    page.wait_for_timeout(1500)

    page.locator(f"text={radicado}").first.click()
    page.wait_for_timeout(5000)


def open_tab_info_equipos(page) -> bool:
    candidates = [
        page.locator("text=Información de equipos"),
        page.locator("text=Informacion de equipos"),
        page.locator("text=Equipos"),
    ]
    for _ in range(15):
        for loc in candidates:
            if loc.count() > 0:
                try:
                    loc.first.scroll_into_view_if_needed(timeout=3000)
                    loc.first.click(timeout=3000)
                    page.wait_for_timeout(2000)
                    return True
                except Exception:
                    pass
        page.mouse.wheel(0, 1200)
        page.wait_for_timeout(250)
    return False


def click_mas_equipos(page):
    plus = page.locator(f"xpath={ADD_PLUS_XPATH}")
    plus.wait_for(state="attached", timeout=20000)
    plus.scroll_into_view_if_needed(timeout=20000)
    plus.click(timeout=20000)


def wait_form_adicionar_equipo(page) -> bool:
    candidates = [
        page.locator("text=Adicionar"),
        page.locator("text=Adicionar equipo"),
        page.locator("text=Agregar equipo"),
        page.locator("label:has-text('Nombre del Elemento')"),
        page.locator("label:has-text('Nombre')"),
    ]
    for _ in range(80):
        for c in candidates:
            if c.count() > 0:
                return True
        page.wait_for_timeout(250)
    return False

def click_guardar(page):
    """
    Busca y hace clic en el botón Guardar o Aceptar.
    """
    print("  -> Guardando equipo...")
    for b in [
        page.get_by_role("button", name="Guardar"),
        page.locator("button:has-text('Guardar')"),
        page.get_by_role("button", name="Aceptar"),
        page.locator("button:has-text('Aceptar')"),
    ]:
        try:
            if b.count() > 0:
                b.first.click(timeout=5000)
                return
        except Exception:
            pass
    page.keyboard.press("Enter")

def fill_field_by_xpath(page, render_xpath: str, value: str, is_dropdown=False):
    """
    Encuentra un campo mediante el atributo 'data-render-xpath' nativo de Bizagi.
    Este es el método MÁS EFECTIVO y robusto.
    """
    if not value:
        return

    print(f"  -> Llenando [{render_xpath}]: {value}")
    try:
        # Buscamos el div contenedor exacto provisto por Bizagi
        container = page.locator(f"div[data-render-xpath='{render_xpath}']").first
        container.wait_for(state="visible", timeout=5000)

        if is_dropdown:
            # Para comboboxes (Nombre del elemento)
            # Primero localizar el input visible (el que dice "Seleccione...")
            dropdown_input = container.locator("input.ui-selectmenu-value, input[role='combobox']").first
            dropdown_input.click(force=True)
            page.wait_for_timeout(500)
            
            # Escribir para filtrar
            dropdown_input.fill("")
            dropdown_input.type(value, delay=30)
            page.wait_for_timeout(1000) # Esperar a que Bizagi traiga las opciones
            
            # Hacer clic en la opción filtrada que coincida (dentro de la lista flotante ui-select-dropdown)
            # Bizagi crea una lista flotante <ul> al final del body
            option = page.locator(f"ul.ui-selectmenu-menu-dropdown li[role='presentation'] a:has-text('{value}')").first
            if option.count() > 0:
                option.click(force=True)
            else:
                # Fallback: presionar Enter si la opción no se detecta claramente
                page.keyboard.press("Enter")
                
            page.wait_for_timeout(500)

        else:
            # Inputs normales (texto/numero)
            field = container.locator("input:not([type='hidden']), textarea").first
            field.click(force=True)
            
            # Método más riguroso para limpiar inputs numéricos con máscaras en Bizagi
            field.press("End")
            for _ in range(25):
                field.press("Backspace")
            
            field.fill("")
            field.press("Control+A")
            field.press("Delete")
            page.wait_for_timeout(150)
            
            # Escribir simulando teclado real, dependiendo del tipo de campo
            # Si es campo numérico/monetario, usamos press carácter por carácter para las máscaras
            is_numeric = render_xpath in ["cValorIVAenCOP", "cValortotalenCOPsinIV", "cValortotalenCOPsinIVA", "iCantidad"]
            
            if is_numeric:
                # Limpieza extra para campos con máscara
                field.click(force=True)
                field.press("Control+A")
                field.press("Backspace")
                page.wait_for_timeout(150)
                
                for char in str(value):
                    page.keyboard.press(char)
                    page.wait_for_timeout(60) # Delay controlado para que la máscara procese
            else:
                # Para campos de texto normal (como "Función"), soporta tildes
                field.type(str(value), delay=20)
            
            # Clicar fuera y tabular para forzar el guardado temporal (OnBlur)
            page.keyboard.press("Tab")
            page.wait_for_timeout(200)
            page.mouse.click(10, 10)
            
        page.wait_for_timeout(300)
    except Exception as e:
        print(f"     ⚠️ Error en campo con xpath '{render_xpath}': {e}")


def find_pdfs_for_item(idx: int, pdf_folder: str) -> list:
    """
    Encuentra todos los PDFs en la carpeta que correspondan al índice (item).
    El índice 0 equivale al item 1 (1., 1 , etc).
    Maneja rangos también (e.g. 42-43).
    """
    item_num = idx + 1
    if not os.path.exists(pdf_folder):
        print(f"  -> La carpeta de PDFs {pdf_folder} no existe.")
        return []
        
    matched_files = []
    files = os.listdir(pdf_folder)
    
    for f in files:
        if not f.lower().endswith(".pdf"):
            continue
            
        # Coincidencia de número base con decimales o espacio:
        # e.g., "10 IINTERFLEX.pdf", "10.1 IINTERFLEX.pdf", "2. EN_Certificate..."
        m = re.match(r"^(\d+)(?:\.\d+)?[\s\.]", f)
        if m:
            base_num = int(m.group(1))
            if base_num == item_num:
                matched_files.append(os.path.join(pdf_folder, f))
                continue
                
        # Coincidencia de rango: e.g., "42-43 Certificado..."
        m_range = re.match(r"^(\d+)\-(\d+)[\s\.]", f)
        if m_range:
            start, end = int(m_range.group(1)), int(m_range.group(2))
            if start <= item_num <= end:
                matched_files.append(os.path.join(pdf_folder, f))
                
    return matched_files


def attach_files_to_equipment(page, file_paths: list):
    """
    Sube múltiples archivos navegando la interfaz de Bizagi según la interacción descrita:
    1. Clic en el icono inicial de subida (<span class="ui-icon upload-file"></span>).
    2. Esperar modal y Clic en 'Seleccionar Archivos'.
    3. Pasar el archivo a la ventana de Windows.
    4. Clic en el botón interno de 'Subir' (<span class="ui-button-text">Subir</span>).
    """
    if not file_paths:
        print("  -> No hay PDFs encontrados para este equipo.")
        return
        
    for path in file_paths:
        print(f"  -> Adjuntando archivo: {os.path.basename(path)}")
        try:
            # 1. Seleccionar el botón de upload principal
            # Buscar explícitamente dentro del contenedor 'fSoporte' o hacer fallback al primero
            clip = page.locator('div[data-render-xpath="fSoporte"] span.ui-icon.upload-file').first
            if clip.count() == 0:
                clip = page.locator('span.ui-icon.upload-file').first
                
            # Darle tiempo a Bizagi a procesar el último dato ingresado
            page.wait_for_timeout(800) 
            clip.scroll_into_view_if_needed()
            clip.wait_for(state="visible", timeout=5000)
            clip.click(force=True)
            
            # MUY IMPORTANTE: Esperar a que Bizagi termine las animaciones y dibuje el modal completo
            page.wait_for_timeout(2500)
            
            try:
                # 2. La manera más directa y probada por Playwright de subir archivos
                # Si el input file existe en el DOM (incluso oculto), set_input_files lo inyecta sin abrir Windows.
                # Buscamos tu elemento exacto: <input id="file" ...>
                upload_input = page.locator('input#file[type="file"]')
                
                # Inyectar la ruta del archivo directamente
                upload_input.set_input_files(path)
                print("     ✅ Archivo inyectado en el sistema.")
                
                # Esperamos a que Bizagi registre el cambio interno
                page.wait_for_timeout(2000)
                
                # Clic en el botón "Subir"
                btn_subir = page.locator("xpath=//span[contains(@class, 'ui-button-text') and text()='Subir']").first
                if btn_subir.count() > 0:
                    btn_subir.click(force=True)
                    print("     ✅ Clic en el botón 'Subir' ejecutado.")
                else:
                    print("     ⚠️ No se encontró el botón 'Subir' con XPath exacto, fallback...")
                    page.locator('button:has-text("Subir"), span.ui-button-text:has-text("Subir")').first.click(force=True)
                    
                # Esperar a que la subida al servidor de Bizagi se procese
                page.wait_for_timeout(3500)
            except Exception as e_inner:
                print(f"     ⚠️ Error inyectando el archivo nativamente: {e_inner}")
                
        except Exception as e:
            print(f"     ⚠️ Error general en la ventana de adjuntos ({os.path.basename(path)}): {e}")


def write_value_then_tab(page, value: str, delay_after_type_ms=140):
    """
    ✅ CORRECCIÓN:
    Escribe y SOLO DESPUÉS hace TAB. Si está vacío, solo TAB.
    Esto evita que el cursor se descuadre por latencia de Bizagi.
    """
    if value:
        page.keyboard.type(value, delay=10)
        page.wait_for_timeout(delay_after_type_ms)
    page.keyboard.press("Tab")
    page.wait_for_timeout(70)


def get_col(row: dict, *names):
    for n in names:
        if n in row:
            return row.get(n)
    return None


def fill_equipo_with_clicks(page, row: dict, idx: int):
    """
    Llenado con CLICS DIRECTOS. 
    Toma los valores del Excel de manera dinámica para cada elemento.
    """
    nombre = s(get_col(row, "Nombre del Elemento", "Nombre del elemento", "Nombre Elemento", "Nombre"))
    marca = s(get_col(row, "Marca"))
    subpartida = s(get_col(row, "Subpartida arancelaria", "Subpartida"))
    unidad = s(get_col(row, "Unidad de Medida", "Unidad", "Unidad medida"))
    fabricante = s(get_col(row, "Fabricante"))
    # Se busca la columna "Función", incluyendo la versión con el caracter corrupto por la codificación del Excel (Funcin / Funci\ufffdn)
    funcion_keys = ["Función", "Funcion", "Funcion ", "Funci\ufffdn"]
    # Agregar iterativamente cualquier columna que empiece por "Funci"
    funcion = s(get_col(row, *funcion_keys))
    if not funcion:
        # Fallback si las llaves no funcionaron, buscar la columna que contenga "Funci"
        for k in row.keys():
            if "Funci" in str(k):
                funcion = s(row[k])
                break
    modelo = s(get_col(row, "Modelo / Referencia", "Modelo", "Referencia"))
    cantidad = fmt_number(get_col(row, "Cantidad"))
    normas = s(get_col(row, "Normas técnicas", "Normas", "Norma tecnica"))
    proveedor = s(get_col(row, "Proveedor"))
    
    # Búsqueda fortalecida de columnas monetarias usando los nombres literales y aproximaciones
    iva_val = fmt_money(get_col(row, "Valor IVA en COP", "Valor IVA", "IVA", "Valor del IVA"))
    if not iva_val:
        iva_val = IVA_FIJO
        
    valor_sin_iva = fmt_money(get_col(row, "Valor total en COP (Sin incluir IVA)", "Valor total en COP", "Valor total (Sin IVA)", "Valor sin IVA", "Valor total sin IVA"))

    # Llenado usando el mapeo exacto de Bizagi (data-render-xpath)
    fill_field_by_xpath(page, "kpElementoFNCE", nombre, is_dropdown=True)
    fill_field_by_xpath(page, "sMarca", marca)
    fill_field_by_xpath(page, "sSubpartidaarancelaria", subpartida)
    fill_field_by_xpath(page, "sUnidaddemedida", unidad)
    fill_field_by_xpath(page, "sFabricante", fabricante)
    fill_field_by_xpath(page, "sFuncion", funcion)
    fill_field_by_xpath(page, "cValorIVAenCOP", iva_val)
    fill_field_by_xpath(page, "sModeloReferencia", modelo)
    fill_field_by_xpath(page, "iCantidad", cantidad)
    fill_field_by_xpath(page, "sNormastecnicas", normas)
    fill_field_by_xpath(page, "sProveedor", proveedor)
    fill_field_by_xpath(page, "cValortotalenCOPsinIV", valor_sin_iva)
    
    # Eliminamos la llamada aquí para evitar duplicidad, se hará en el main loop


# ---------------- MAIN ---------------- #
def load_clean_excel(file_path: str, sheet_name: str) -> pd.DataFrame:
    print("Analizando estructura del archivo Excel para encontrar cabeceras...")
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    header_idx = -1
    for i, row in raw_df.iterrows():
        row_str = " ".join([str(x).lower() for x in row.values])
        if "elemento" in row_str and "marca" in row_str:
            header_idx = i
            break
            
    if header_idx != -1:
        print(f"-> Cabeceras encontradas dinámicamente en la fila {header_idx + 1} del Excel.")
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_idx)
    else:
        print("-> No se detectó preámbulo, leyendo de modo estándar.")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    df = normalize_columns(df)
    
    # Filtrar solo las filas que tengan un Nombre válido de elemento (ignorar footers vacíos)
    valid_rows = []
    for _, row in df.iterrows():
        n = s(get_col(row.to_dict(), "Nombre del Elemento", "Nombre del elemento", "Nombre Elemento", "Nombre"))
        if n and n.lower() != "nan" and n.lower() != "none" and "total" not in n.lower():
            valid_rows.append(row)
            
    if valid_rows:
        df = pd.DataFrame(valid_rows).reset_index(drop=True)
    return df

print("Leyendo archivo Excel...")
df = load_clean_excel(EXCEL_FILE, SHEET)

print("Columnas detectadas en Excel:")
for c in df.columns:
    print(" -", c)

print(f"Equipos detectados emparejados listos para cargar: {len(df)}")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=120)
    page = browser.new_page()

    try:
        print("Abriendo plataforma...")
        page.goto(URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1500)
        set_zoom(page, 0.8)

        print("Login...")
        do_login(page)
        page.wait_for_timeout(6500)

        print("Entrando a Inbox...")
        open_inbox(page)

        print("Buscando radicado:", RADICADO)
        open_case_by_radicado(page, RADICADO)

        print("Abriendo pestaña Información de equipos...")
        if not open_tab_info_equipos(page):
            print("❌ No pude abrir la pestaña 'Información de equipos'.")
            input("ENTER para dejar el navegador abierto...")
            raise SystemExit(1)

        print("✅ Pestaña abierta. Iniciando carga...")

        for i, row in df.iterrows():
            row_dict = row.to_dict()
            nombre_dbg = s(get_col(row_dict, "Nombre del Elemento", "Nombre del elemento", "Nombre Elemento", "Nombre"))
            print(f"[{i+1}/{len(df)}] Cargando: {nombre_dbg}")

            click_mas_equipos(page)

            if not wait_form_adicionar_equipo(page):
                print("❌ No abrió el formulario de 'Adicionar equipo'.")
                input("ENTER para dejar el navegador abierto...")
                raise SystemExit(1)

            # Carga de datos
            fill_equipo_with_clicks(page, row_dict, i)

            # --- PASO DE ADJUNTOS (Soporte Técnico) ---
            pdfs_to_upload = find_pdfs_for_item(i, PDF_DIR)
            if pdfs_to_upload:
                attach_files_to_equipment(page, pdfs_to_upload)
            else:
                print("     ℹ️ No se encontraron PDFs para este elemento.")

            # --- GUARDAR ELEMENTO (PASO FINAL) ---
            print("  -> Finalizando equipo (Esperando 3 segundos para Guardar)...")
            
            # Pausa exacta de 3 segundos despues de "Subir" solicitada por el usuario
            page.wait_for_timeout(3000)
            
            # Llamamos a nuestra función de Guardar original, que busca botones reales de "Guardar" o "Aceptar"
            click_guardar(page)
            
            # Tiempo prudente para que la animación de guardado termine y la vista vuelva a la tabla principal
            # antes de intentar abrir el formulario del siguiente equipo
            page.wait_for_timeout(4000)

        print("✅ Carga finalizada.")
        input("ENTER para cerrar (o deja abierto para revisar)...")

    except Exception as e:
        print("ERROR:", e)
        input("ENTER para dejar el navegador abierto y revisar...")