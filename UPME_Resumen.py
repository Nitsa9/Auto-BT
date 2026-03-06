"""
UPME Dashboard - Scraper + Visualizador
========================================
Extrae los casos de la Bandeja de Entrada de UPME Bizagi
y abre un dashboard visual en el navegador.

Requisitos:
    pip install selenium webdriver-manager

Uso:
    python upme_dashboard.py
"""

import time
import json
import webbrowser
import http.server
import threading
import tempfile
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# ─────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────
URL_BASE = "https://automatizacion-upme.bizagi.com/"
USUARIO  = "1000396872"
PASSWORD = "jn1g0xMOqg"   # <- actualiza si cambiaste tu contrasena
HEADLESS = False
PORT     = 8765
# ─────────────────────────────────────────


# ══════════════════════════════════════════
#  SCRAPING
# ══════════════════════════════════════════

def init_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1600,900")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )


def login(driver):
    print("Abriendo portal UPME...")
    driver.get(URL_BASE)
    time.sleep(5)

    for sel in ["input[name='username']", "#txtUserID", "input[name='usuario']", "input[type='text']"]:
        try:
            f = driver.find_element(By.CSS_SELECTOR, sel)
            f.clear(); f.send_keys(USUARIO)
            print(f"  Usuario ingresado ({sel})")
            break
        except: continue

    for sel in ["input[name='password']", "#txtPassword", "input[type='password']"]:
        try:
            f = driver.find_element(By.CSS_SELECTOR, sel)
            f.clear(); f.send_keys(PASSWORD)
            print(f"  Contrasena ingresada ({sel})")
            break
        except: continue

    for sel in ["button[type='submit']", "#btnLogin", "input[type='submit']", "button"]:
        try:
            driver.find_element(By.CSS_SELECTOR, sel).click()
            print(f"  Login enviado ({sel})")
            break
        except: continue

    time.sleep(6)
    print(f"  URL: {driver.current_url}")


def navegar_bandeja(driver):
    print("Navegando a Bandeja de Entrada...")
    xpaths = [
        "//span[contains(text(),'Bandeja')]",
        "//a[contains(text(),'Bandeja')]",
        "//div[contains(text(),'Bandeja')]",
        "//li[contains(text(),'Bandeja')]",
    ]
    for xp in xpaths:
        try:
            el = WebDriverWait(driver, 6).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            el.click()
            print(f"  Bandeja abierta")
            time.sleep(4)
            return True
        except: continue

    driver.save_screenshot("upme_debug.png")
    print("  No encontro menu Bandeja. Screenshot: upme_debug.png")
    return False


def extraer_filas_tabla(driver):
    """Extrae filas de la tabla visible en pantalla."""
    time.sleep(2)
    filas_data = []

    # Encabezados
    encabezados = []
    for sel in ["table thead th", "[role='columnheader']", "th"]:
        ths = driver.find_elements(By.CSS_SELECTOR, sel)
        if ths:
            encabezados = [th.text.strip() for th in ths if th.text.strip()]
            if encabezados:
                break

    print(f"    Encabezados encontrados: {encabezados}")

    # Filas de datos
    filas = []
    for sel in ["table tbody tr", "[role='row']", "table tr", ".grid-row"]:
        filas = driver.find_elements(By.CSS_SELECTOR, sel)
        if filas:
            print(f"    {len(filas)} filas con selector: {sel}")
            break

    # Si no encontro nada con selectores normales, guardar HTML para debug
    if not filas:
        with open("upme_page.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("    Sin filas. HTML guardado: upme_page.html")
        return []

    keys_default = ["_icono", "Numero del caso", "Proceso", "Actividad",
                    "Fecha creacion caso", "Actividad vence en", "Fecha Solucion caso"]

    for fila in filas:
        celdas = fila.find_elements(By.CSS_SELECTOR, "td, [role='gridcell']")
        textos = [c.text.strip() for c in celdas]
        if not any(textos):
            continue

        if encabezados and len(encabezados) == len(textos):
            row = dict(zip(encabezados, textos))
        else:
            row = {}
            for i, t in enumerate(textos):
                k = keys_default[i] if i < len(keys_default) else f"col_{i}"
                row[k] = t

        row["_extraido"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        filas_data.append(row)

    return filas_data


def extraer_secciones(driver):
    """
    Recorre el menú lateral de la Bandeja y extrae cada categoría.
    Si no hay menú lateral, extrae directamente la vista actual.
    """
    print("Extrayendo datos...")
    secciones = {}

    # Intentar leer items del sidebar (Bizagi suele tenerlos en li o div con texto)
    items = []
    selectores_sidebar = [
        "//li[contains(@class,'item')]",
        "//div[contains(@class,'inbox-item')]",
        "//div[contains(@class,'workarea')]//li",
        "//ul[@class]//li",
        "//*[contains(@class,'category')]",
    ]
    for xp in selectores_sidebar:
        items = driver.find_elements(By.XPATH, xp)
        if items:
            print(f"  Menu lateral: {len(items)} items ({xp})")
            break

    # Si hay items en el sidebar, iterar por cada uno
    if items:
        for item in items:
            nombre = item.text.strip().split("\n")[0]  # solo primera linea
            if not nombre or len(nombre) < 3:
                continue
            try:
                item.click()
                time.sleep(3)
                filas = extraer_filas_tabla(driver)
                if filas:
                    # Limpiar contador del badge del nombre  ej: "Incentivos 84" -> "Incentivos"
                    nombre_limpio = nombre.rsplit(" ", 1)[0] if nombre[-1].isdigit() else nombre
                    secciones[nombre_limpio] = filas
                    print(f"  Seccion '{nombre_limpio}': {len(filas)} casos")
            except Exception as e:
                print(f"  Error en '{nombre}': {e}")
    else:
        # Sin sidebar: extraer directamente lo que hay en pantalla
        print("  Sin sidebar detectado, extrayendo vista actual...")
        filas = extraer_filas_tabla(driver)
        if filas:
            secciones["Bandeja de Entrada"] = filas
            print(f"  Total extraido: {len(filas)} casos")
        else:
            # Ultimo intento: capturar screenshot y HTML
            driver.save_screenshot("upme_bandeja.png")
            with open("upme_bandeja.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print("  Sin datos. Archivos de debug: upme_bandeja.png, upme_bandeja.html")

    return secciones


# ══════════════════════════════════════════
#  HTML DEL DASHBOARD
# ══════════════════════════════════════════

def generar_html(secciones):
    total = sum(len(v) for v in secciones.values())
    datos_json = json.dumps(secciones, ensure_ascii=False)

    # El HTML se construye con concatenacion para evitar conflictos con f-strings y llaves JS
    css = """
  :root {
    --green:#4a6741; --green-dark:#3a5233; --green-light:#e8ede7;
    --gold:#c8a94a; --red:#c0392b; --border:#d0d8cf;
    --text:#2c3e2d; --muted:#5a6b5b; --faint:#8a9b8b; --bg:#f4f5f4;
  }
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Inter',sans-serif; background:var(--bg); color:var(--text); height:100vh; display:flex; flex-direction:column; }

  .topbar { background:var(--green-dark); height:52px; display:flex; align-items:center; padding:0 20px; gap:6px; box-shadow:0 2px 8px rgba(0,0,0,.2); flex-shrink:0; }
  .logo { color:#fff; font-weight:800; font-size:16px; letter-spacing:1.5px; margin-right:24px; }
  .logo-badge { background:rgba(255,255,255,.15); border:1.5px solid rgba(255,255,255,.35); border-radius:6px; padding:3px 10px; }
  .nav-btn { background:transparent; border:none; color:rgba(255,255,255,.7); padding:7px 16px; border-radius:6px; cursor:pointer; font-size:13px; font-weight:500; transition:.15s; }
  .nav-btn:hover { background:rgba(255,255,255,.1); color:#fff; }
  .nav-btn.active { background:rgba(255,255,255,.15); color:#fff; border-bottom:2.5px solid var(--gold); border-radius:6px 6px 0 0; }
  .topbar-right { margin-left:auto; display:flex; align-items:center; gap:10px; }
  .search { background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.2); border-radius:6px; padding:6px 12px; color:#fff; font-size:13px; width:210px; outline:none; }
  .search::placeholder { color:rgba(255,255,255,.45); }
  .search:focus { background:rgba(255,255,255,.2); border-color:var(--gold); }
  .avatar { width:34px; height:34px; background:var(--gold); border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-weight:700; font-size:13px; }

  .layout { display:flex; flex:1; overflow:hidden; }

  .sidebar { width:238px; background:#fff; border-right:1px solid var(--border); overflow-y:auto; flex-shrink:0; padding:10px 0; }
  .sb-label { padding:8px 16px 4px; font-size:10.5px; font-weight:700; text-transform:uppercase; letter-spacing:.8px; color:var(--faint); }
  .sb-item { display:flex; align-items:center; justify-content:space-between; padding:9px 14px 9px 18px; cursor:pointer; font-size:13px; color:var(--muted); border-left:3px solid transparent; transition:.15s; }
  .sb-item:hover { background:var(--green-light); color:var(--green); }
  .sb-item.active { background:var(--green-light); color:var(--green-dark); font-weight:600; border-left-color:var(--green); }
  .sb-item .sb-name { flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .badge { background:var(--green); color:#fff; border-radius:10px; padding:1px 7px; font-size:11px; font-weight:600; flex-shrink:0; margin-left:6px; }
  .badge.gold { background:var(--gold); }

  .main { flex:1; overflow-y:auto; display:flex; flex-direction:column; }

  .stats { background:#fff; border-bottom:1px solid var(--border); padding:12px 22px; display:flex; gap:20px; align-items:center; flex-shrink:0; }
  .stat { display:flex; flex-direction:column; gap:1px; }
  .stat-n { font-size:22px; font-weight:700; color:var(--green-dark); }
  .stat-l { font-size:10.5px; color:var(--faint); text-transform:uppercase; letter-spacing:.5px; }
  .sdiv { width:1px; height:34px; background:var(--border); }

  .content { padding:16px 22px; flex:1; }
  .c-header { display:flex; align-items:center; justify-content:space-between; margin-bottom:14px; }
  .c-title { font-size:14.5px; font-weight:600; }
  .toolbar { display:flex; gap:8px; align-items:center; }
  .tb-btn { background:#fff; border:1px solid var(--border); border-radius:6px; padding:5px 12px; font-size:12px; cursor:pointer; color:var(--muted); display:flex; align-items:center; gap:5px; transition:.15s; }
  .tb-btn:hover { border-color:var(--green); color:var(--green); }
  .ppage { display:flex; align-items:center; gap:7px; font-size:12px; color:var(--muted); }
  .ppage select { border:1px solid var(--border); border-radius:4px; padding:4px 8px; font-size:12px; background:#fff; }

  .t-wrap { background:#fff; border:1px solid var(--border); border-radius:8px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,.06); }
  table { width:100%; border-collapse:collapse; font-size:12.5px; }
  thead tr { background:var(--green); }
  thead th { padding:10px 13px; text-align:left; color:#fff; font-weight:500; font-size:12px; white-space:nowrap; }
  thead th:first-child { width:44px; text-align:center; }
  tbody tr { border-bottom:1px solid #f0f2f0; transition:background .1s; cursor:pointer; }
  tbody tr:hover { background:var(--green-light); }
  tbody tr:last-child { border-bottom:none; }
  td { padding:9px 13px; vertical-align:middle; }
  td:first-child { text-align:center; }
  .cid { font-weight:600; color:var(--green-dark); }
  .act-cell { display:flex; align-items:center; gap:7px; }
  .dot { width:8px; height:8px; border-radius:50%; background:#e74c3c; flex-shrink:0; }
  .date { font-size:12px; color:var(--muted); white-space:nowrap; }
  .vence { font-size:12px; font-weight:500; white-space:nowrap; }
  .vence.v { color:var(--red); font-weight:700; }
  .vence.p { color:#e67e22; }
  .vence.ok { color:var(--muted); }
  .star { background:none; border:none; cursor:pointer; color:var(--gold); font-size:14px; }
  .acts { display:flex; flex-direction:column; align-items:center; gap:2px; }

  .footer { padding:9px 22px; font-size:11.5px; color:var(--faint); border-top:1px solid var(--border); background:#fff; display:flex; justify-content:space-between; flex-shrink:0; }

  .overlay { display:none; position:fixed; inset:0; background:rgba(0,0,0,.4); z-index:200; align-items:center; justify-content:center; }
  .overlay.open { display:flex; }
  .modal { background:#fff; border-radius:10px; width:580px; max-width:92vw; padding:26px; box-shadow:0 20px 60px rgba(0,0,0,.22); }
  .mh { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:18px; }
  .mt { font-size:16px; font-weight:700; color:var(--green-dark); }
  .ms { font-size:12px; color:var(--faint); margin-top:3px; }
  .mc { background:none; border:none; font-size:20px; cursor:pointer; color:var(--faint); }
  .mr { display:flex; gap:16px; margin-bottom:11px; }
  .mf { flex:1; }
  .mf label { font-size:10.5px; font-weight:700; color:var(--faint); text-transform:uppercase; letter-spacing:.5px; display:block; margin-bottom:3px; }
  .mf span { font-size:13px; font-weight:500; }
  .mdiv { height:1px; background:var(--border); margin:14px 0; }
  pre.raw { font-size:11px; color:var(--muted); white-space:pre-wrap; background:var(--bg); padding:10px; border-radius:6px; max-height:170px; overflow-y:auto; margin-top:4px; }
  ::-webkit-scrollbar { width:5px; height:5px; }
  ::-webkit-scrollbar-thumb { background:var(--border); border-radius:3px; }
"""

    js = """
const D = __DATOS__;
const hoy = new Date();
let SEC = "__todos__", BUS = "", _lista = [];

const $ = id => document.getElementById(id);

function todos() { return Object.values(D).flat(); }
function deSec(s) { return s === "__todos__" ? todos() : (D[s] || []); }

function parseFecha(s) {
  if (!s) return null;
  const m = s.match(/(\\d{1,2})\\/(\\d{1,2})\\/(\\d{4})\\s+(\\d{1,2}):(\\d{2})\\s*(am|pm)/i);
  if (!m) return null;
  let h = +m[4];
  if (/pm/i.test(m[6]) && h !== 12) h += 12;
  if (/am/i.test(m[6]) && h === 12) h = 0;
  return new Date(+m[3], m[2]-1, +m[1], h, +m[5]);
}

function clsV(s) {
  const f = parseFecha(s);
  if (!f) return "ok";
  const d = (f - hoy) / 86400000;
  return d < 0 ? "v" : d < 30 ? "p" : "ok";
}

function renderSidebar() {
  let h = "";
  for (const [k, v] of Object.entries(D)) {
    const label = k.length > 30 ? k.slice(0, 28) + "…" : k;
    h += `<div class="sb-item" data-sec="${k}">
      <span class="sb-name">📁 ${label}</span>
      <span class="badge gold">${v.length}</span>
    </div>`;
  }
  $("sb-secs").innerHTML = h;
  document.querySelectorAll("#sb-secs .sb-item").forEach(el => {
    el.addEventListener("click", () => setSec(el.dataset.sec, el));
  });
}

function setSec(s, el) {
  SEC = s;
  document.querySelectorAll(".sb-item").forEach(e => e.classList.remove("active"));
  if (el) el.classList.add("active");
  $("c-title").textContent = s === "__todos__" ? "Todos los casos" : s;
  render();
}

function render() {
  const pp = +$("perPage").value;
  let casos = deSec(SEC);
  if (BUS) casos = casos.filter(c => JSON.stringify(c).toLowerCase().includes(BUS));
  _lista = casos;

  const venc = casos.filter(c => clsV(c["Actividad vence en"] || c["actividad_vence"] || "") === "v").length;
  $("s-total").textContent = todos().length;
  $("s-activos").textContent = casos.length - venc;
  $("s-vencidos").textContent = venc;
  const ext = (casos[0] && casos[0]["_extraido"]) || "–";
  $("s-ext").textContent = ext;
  $("fext").textContent = ext;
  $("finfo").textContent = `Mostrando ${Math.min(pp, casos.length)} de ${casos.length} casos`;

  const show = casos.slice(0, pp);
  const tbody = $("tbody");

  if (!show.length) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:52px;color:#aaa">
      <div style="font-size:30px;margin-bottom:8px">📭</div>No hay casos para mostrar</td></tr>`;
    return;
  }

  tbody.innerHTML = show.map((c, i) => {
    const num = c["Número del caso"] || c["Numero del caso"] || c["numero_caso"] || c["col_1"] || "–";
    const pro = c["Proceso"] || c["proceso"] || c["col_2"] || "–";
    const act = c["Actividad"] || c["actividad"] || c["col_3"] || "–";
    const fcr = c["Fecha creación caso"] || c["Fecha creacion caso"] || c["fecha_creacion"] || c["col_4"] || "–";
    const fv  = c["Actividad vence en"] || c["actividad_vence"] || c["col_5"] || "–";
    const fs  = c["Fecha Solución caso"] || c["Fecha Solucion caso"] || c["fecha_solucion"] || c["col_6"] || "–";
    const cv  = clsV(fv);
    const tag = cv === "v" ? "⚠ " : cv === "p" ? "⏰ " : "";
    return `<tr onclick="openM(${i})">
      <td><div class="acts">
        <button class="star" onclick="event.stopPropagation()">★</button>
      </div></td>
      <td><span class="cid">${num}</span></td>
      <td>${pro}</td>
      <td><div class="act-cell"><div class="dot"></div><span>${act}</span></div></td>
      <td class="date">${fcr}</td>
      <td class="vence ${cv}">${tag}${fv}</td>
      <td class="date">${fs}</td>
    </tr>`;
  }).join("");
}

function openM(i) {
  const c = _lista[i]; if (!c) return;
  const get = (keys) => { for (const k of keys) if (c[k]) return c[k]; return "–"; };
  $("m-id").textContent   = get(["Número del caso","Numero del caso","numero_caso","col_1"]);
  $("m-proc").textContent = get(["Proceso","proceso","col_2"]);
  $("m-act").textContent  = get(["Actividad","actividad","col_3"]);
  $("m-fcr").textContent  = get(["Fecha creación caso","Fecha creacion caso","fecha_creacion","col_4"]);
  $("m-fv").textContent   = get(["Actividad vence en","actividad_vence","col_5"]);
  $("m-fs").textContent   = get(["Fecha Solución caso","Fecha Solucion caso","fecha_solucion","col_6"]);
  const clean = Object.fromEntries(Object.entries(c).filter(([k]) => !k.startsWith("_")));
  $("m-raw").textContent = JSON.stringify(clean, null, 2);
  $("overlay").classList.add("open");
}
function closeM() { $("overlay").classList.remove("open"); }

function exportCSV() {
  if (!_lista.length) return;
  const keys = Object.keys(_lista[0]).filter(k => !k.startsWith("_"));
  const rows = [keys.join(","), ..._lista.map(c => keys.map(k => `"${(c[k]||"").replace(/"/g,'""')}"`).join(","))];
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([rows.join("\\n")], {type:"text/csv;charset=utf-8"}));
  a.download = "upme_casos.csv"; a.click();
}

renderSidebar();
setSec("__todos__");
document.getElementById("gSearch").addEventListener("input", e => { BUS = e.target.value.toLowerCase(); render(); });
"""

    # Inyectar datos reales en el JS
    js_final = js.replace("__DATOS__", datos_json)

    html = (
        "<!DOCTYPE html>\n"
        "<html lang='es'>\n"
        "<head>\n"
        "<meta charset='UTF-8'>\n"
        "<meta name='viewport' content='width=device-width, initial-scale=1.0'>\n"
        "<title>UPME \u2013 Bandeja de Entrada</title>\n"
        "<link href='https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap' rel='stylesheet'>\n"
        "<style>" + css + "</style>\n"
        "</head>\n"
        "<body>\n"
        "\n"
        "<div class='topbar'>\n"
        "  <div class='logo'><span class='logo-badge'>UPME</span></div>\n"
        "  <button class='nav-btn'>\U0001f3e0 Mi Portal</button>\n"
        "  <button class='nav-btn active'>\U0001f4e5 Bandeja de entrada</button>\n"
        "  <button class='nav-btn'>\U0001f4cb Nuevo Caso \u25be</button>\n"
        "  <div class='topbar-right'>\n"
        "    <input class='search' id='gSearch' placeholder='\U0001f50d  Buscar caso...'>\n"
        "    <div class='avatar'>SD</div>\n"
        "  </div>\n"
        "</div>\n"
        "\n"
        "<div class='layout'>\n"
        "  <div class='sidebar'>\n"
        "    <div class='sb-label'>Bandeja</div>\n"
        "    <div class='sb-item active' id='sb-todos' onclick=\"setSec('__todos__', this)\">\n"
        "      <span class='sb-name'>\U0001f4c2 Todos los casos</span>\n"
        "      <span class='badge' id='b-todos'>" + str(total) + "</span>\n"
        "    </div>\n"
        "    <div id='sb-secs'></div>\n"
        "  </div>\n"
        "\n"
        "  <div class='main'>\n"
        "    <div class='stats'>\n"
        "      <div class='stat'><span class='stat-n' id='s-total'>" + str(total) + "</span><span class='stat-l'>Total casos</span></div>\n"
        "      <div class='sdiv'></div>\n"
        "      <div class='stat'><span class='stat-n' id='s-activos' style='color:#e67e22'>\u2013</span><span class='stat-l'>Activos</span></div>\n"
        "      <div class='sdiv'></div>\n"
        "      <div class='stat'><span class='stat-n' id='s-vencidos' style='color:#c0392b'>\u2013</span><span class='stat-l'>Vencidos</span></div>\n"
        "      <div class='sdiv'></div>\n"
        "      <div class='stat'><span class='stat-n' id='s-ext' style='font-size:12px;color:#5a6b5b'>\u2013</span><span class='stat-l'>Ultima actualizaci\u00f3n</span></div>\n"
        "    </div>\n"
        "\n"
        "    <div class='content'>\n"
        "      <div class='c-header'>\n"
        "        <span class='c-title' id='c-title'>Todos los casos</span>\n"
        "        <div class='toolbar'>\n"
        "          <div class='ppage'>Resultados por p\u00e1gina\n"
        "            <select id='perPage' onchange='render()'>\n"
        "              <option value='25'>25</option>\n"
        "              <option value='50'>50</option>\n"
        "              <option value='100' selected>100</option>\n"
        "              <option value='9999'>Todos</option>\n"
        "            </select>\n"
        "          </div>\n"
        "          <button class='tb-btn' onclick='exportCSV()'>\u2b07 CSV</button>\n"
        "          <button class='tb-btn' onclick='location.reload()'>\u21ba Actualizar</button>\n"
        "        </div>\n"
        "      </div>\n"
        "      <div class='t-wrap'>\n"
        "        <table>\n"
        "          <thead><tr>\n"
        "            <th></th>\n"
        "            <th>N\u00famero del caso</th>\n"
        "            <th>Proceso</th>\n"
        "            <th>Actividad</th>\n"
        "            <th>Fecha creaci\u00f3n caso</th>\n"
        "            <th>Actividad vence en</th>\n"
        "            <th>Fecha Soluci\u00f3n caso</th>\n"
        "          </tr></thead>\n"
        "          <tbody id='tbody'><tr><td colspan='7' style='text-align:center;padding:40px;color:#aaa'>Cargando\u2026</td></tr></tbody>\n"
        "        </table>\n"
        "      </div>\n"
        "    </div>\n"
        "\n"
        "    <div class='footer'>\n"
        "      <span id='finfo'>\u2013</span>\n"
        "      <span>UPME \u00b7 Sistema de Gesti\u00f3n de Casos \u00b7 Extra\u00eddo: <b id='fext'>\u2013</b></span>\n"
        "    </div>\n"
        "  </div>\n"
        "</div>\n"
        "\n"
        "<div class='overlay' id='overlay' onclick=\"if(event.target===this)closeM()\">\n"
        "  <div class='modal'>\n"
        "    <div class='mh'>\n"
        "      <div><div class='mt' id='m-id'>\u2013</div><div class='ms' id='m-proc'>\u2013</div></div>\n"
        "      <button class='mc' onclick='closeM()'>&#x2715;</button>\n"
        "    </div>\n"
        "    <div class='mr'><div class='mf' style='flex:1'><label>Actividad</label><span id='m-act'>\u2013</span></div></div>\n"
        "    <div class='mdiv'></div>\n"
        "    <div class='mr'>\n"
        "      <div class='mf'><label>Fecha creaci\u00f3n</label><span id='m-fcr'>\u2013</span></div>\n"
        "      <div class='mf'><label>Vence en</label><span id='m-fv'>\u2013</span></div>\n"
        "      <div class='mf'><label>Fecha soluci\u00f3n</label><span id='m-fs'>\u2013</span></div>\n"
        "    </div>\n"
        "    <div class='mdiv'></div>\n"
        "    <div class='mf'><label>Datos completos</label><pre class='raw' id='m-raw'></pre></div>\n"
        "  </div>\n"
        "</div>\n"
        "\n"
        "<script>\n" + js_final + "\n</script>\n"
        "</body>\n"
        "</html>\n"
    )
    return html


# ══════════════════════════════════════════
#  SERVIDOR HTTP LOCAL
# ══════════════════════════════════════════

_HTML_PATH = None

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        with open(_HTML_PATH, "rb") as f:
            self.wfile.write(f.read())
    def log_message(self, *a): pass

def start_server(path, port):
    global _HTML_PATH
    _HTML_PATH = path
    srv = http.server.HTTPServer(("localhost", port), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


# ══════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════

def main():
    print("=" * 52)
    print("  UPME Dashboard  |", datetime.now().strftime("%Y-%m-%d %H:%M"))
    print("=" * 52)

    driver = init_driver()
    secciones = {}

    try:
        login(driver)
        ok = navegar_bandeja(driver)
        if ok:
            secciones = extraer_secciones(driver)
        else:
            print("No se pudo abrir la Bandeja.")
    except Exception as e:
        print(f"Error durante scraping: {e}")
        try:
            driver.save_screenshot("upme_error.png")
        except:
            pass
    finally:
        driver.quit()
        print("Navegador cerrado.")

    # Fallback con datos del screenshot si no extrae nada
    if not secciones or not any(secciones.values()):
        print("Usando datos de ejemplo (no se extrajeron datos reales).")
        secciones = {
            "Incentivos Tributarios FNCE": [
                {"Número del caso": "FNCE_20251022", "Proceso": "Incentivos Tributarios FNCE", "Actividad": "Solicitar modificación de certificación", "Fecha creación caso": "26/02/2025 9:49 am", "Actividad vence en": "11/10/2025 9:47 pm", "Fecha Solución caso": "26/02/2025 9:49 am", "_extraido": "Ejemplo"},
                {"Número del caso": "FNCE_20251105", "Proceso": "Incentivos Tributarios FNCE", "Actividad": "Solicitar modificación de certificación", "Fecha creación caso": "28/02/2025 8:07 am", "Actividad vence en": "19/08/2025 7:55 pm", "Fecha Solución caso": "28/02/2025 8:07 am", "_extraido": "Ejemplo"},
                {"Número del caso": "FNCE_20251465", "Proceso": "Incentivos Tributarios FNCE", "Actividad": "Solicitar modificación de certificación", "Fecha creación caso": "12/03/2025 2:59 pm", "Actividad vence en": "20/10/2025 4:41 pm", "Fecha Solución caso": "12/03/2025 2:59 pm", "_extraido": "Ejemplo"},
                {"Número del caso": "FNCE_20251641", "Proceso": "Incentivos Tributarios FNCE", "Actividad": "Solicitar modificación de certificación", "Fecha creación caso": "17/03/2025 3:54 pm", "Actividad vence en": "21/09/2025 4:31 pm", "Fecha Solución caso": "17/03/2025 3:54 pm", "_extraido": "Ejemplo"},
                {"Número del caso": "FNCE_20251649", "Proceso": "Incentivos Tributarios FNCE", "Actividad": "Solicitar modificación de certificación", "Fecha creación caso": "17/03/2025 5:19 pm", "Actividad vence en": "22/06/2025 5:36 pm", "Fecha Solución caso": "17/03/2025 5:19 pm", "_extraido": "Ejemplo"},
                {"Número del caso": "FNCE_20251806", "Proceso": "Incentivos Tributarios FNCE", "Actividad": "Solicitar modificación de certificación", "Fecha creación caso": "25/03/2025 7:27 am", "Actividad vence en": "26/06/2025 5:38 pm", "Fecha Solución caso": "25/03/2025 7:27 am", "_extraido": "Ejemplo"},
            ]
        }

    # Generar HTML y abrir dashboard
    html = generar_html(secciones)
    tmp = Path(tempfile.gettempdir()) / "upme_dashboard.html"
    tmp.write_text(html, encoding="utf-8")
    print(f"Dashboard generado: {tmp}")

    start_server(str(tmp), PORT)
    url = "http://localhost:" + str(PORT)
    print(f"Abriendo dashboard en: {url}")
    webbrowser.open(url)

    print("\nDashboard activo. Presiona Ctrl+C para cerrar.\n")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Cerrando.")


if __name__ == "__main__":
    main()