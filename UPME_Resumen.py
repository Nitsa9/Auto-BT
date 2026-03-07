"""
UPME Dashboard - API REST Scraper + Visualizador
=================================================
Extrae los casos de la Bandeja de Entrada de UPME Bizagi
usando la API REST interna (sin scraping de DOM).

Requisitos:
    pip install selenium webdriver-manager

Uso:
    python UPME_Resumen.py
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

# ─────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────
URL_BASE  = "https://automatizacion-upme.bizagi.com/"
USUARIO   = "1000396872"
PASSWORD  = "jn1g0xMOqg"
HEADLESS  = False
PORT      = 8765
# ─────────────────────────────────────────


# ══════════════════════════════════════════
#  DRIVER
# ══════════════════════════════════════════

def init_driver():
    return webdriver.Safari()


# ══════════════════════════════════════════
#  LOGIN
# ══════════════════════════════════════════

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

    time.sleep(8)
    print(f"  URL post-login: {driver.current_url}")


# ══════════════════════════════════════════
#  CAPA API  (llamadas via jQuery del browser)
# ══════════════════════════════════════════

_JS_AJAX = """
var callback = arguments[arguments.length - 1];
var method   = arguments[0];
var url      = arguments[1];
var params   = arguments[2] || {};
var postData = arguments[3] || null;

if (method === 'GET' && Object.keys(params).length > 0) {
    var qs = Object.keys(params)
        .map(function(k){ return k + '=' + encodeURIComponent(params[k]); })
        .join('&');
    url = url + '?' + qs + '&_=' + Date.now();
}

$.ajax({
    url:      url,
    type:     method,
    data:     postData,
    dataType: 'json',
    success:  function(data)           { callback({ok: true,  data: data}); },
    error:    function(xhr, st, err)   { callback({ok: false, status: xhr.status,
                                                   error: err, text: xhr.responseText}); }
});
"""

def api_get(driver, path, params=None):
    result = driver.execute_async_script(_JS_AJAX, "GET", path, params or {}, None)
    if not result or not result.get("ok"):
        print(f"    [GET {path}] error: {result}")
        return None
    return result["data"]


def api_post(driver, path, form_data):
    result = driver.execute_async_script(_JS_AJAX, "POST", path, {}, form_data)
    if not result or not result.get("ok"):
        print(f"    [POST {path}] error: {result}")
        return None
    return result["data"]


# ══════════════════════════════════════════
#  HELPERS RENDER JSON
# ══════════════════════════════════════════

def _walk(elements):
    """Generador que recorre recursivamente todos los nodos render/container."""
    for el in elements:
        if "render" in el:
            yield ("render", el["render"]["properties"])
        for key in ("container", "nestedForm", "tab", "tabItem",
                    "group", "horizontal", "panel", "contentPanel"):
            if key in el:
                yield (key, el[key]["properties"])
                yield from _walk(el[key].get("elements", []))


def find_field(form_elements, xpath_substr):
    """Devuelve el 'value' del primer render cuyo xpath contiene xpath_substr."""
    for kind, props in _walk(form_elements):
        if kind == "render" and xpath_substr in props.get("xpath", ""):
            return props.get("value", "")
    return ""


def find_grid(form_elements, xpath_substr):
    """Devuelve las properties del primer grid cuyo xpath contiene xpath_substr."""
    for kind, props in _walk(form_elements):
        if kind == "render" and props.get("type") == "grid" and xpath_substr in props.get("xpath", ""):
            return props
    return None


def find_grid_with_parent(elements, xpath_substr, parent_id=None):
    """Devuelve (grid_props, parent_container_id) buscando recursivamente."""
    for el in elements:
        if "render" in el:
            props = el["render"]["properties"]
            if props.get("type") == "grid" and xpath_substr in props.get("xpath", ""):
                return props, parent_id
        for key in ("container", "nestedForm", "tab", "tabItem",
                    "group", "horizontal", "panel", "contentPanel"):
            if key in el:
                this_id  = el[key]["properties"].get("id", parent_id)
                children = el[key].get("elements", [])
                result, found_parent = find_grid_with_parent(children, xpath_substr, this_id)
                if result is not None:
                    return result, found_parent
    return None, None


# ══════════════════════════════════════════
#  EXTRACCION DE DATOS
# ══════════════════════════════════════════

def get_workflows(driver):
    print("Consultando workflows...")
    data = api_get(driver, "/Rest/Inbox/FullSummaryWithLiveProcesses", {"taskState": "all"})
    if not data:
        return []
    workflows = data.get("processes", {}).get("workflows", {}).get("workFlow", [])
    for wf in workflows:
        print(f"  - {wf['name']}  ({wf['count']} casos)")
    return workflows


def get_cases_for_workflow(driver, id_workflow):
    data = api_get(driver, "/Rest/Processes/CustomizedColumnsData", {
        "smartInboxFilter": "W10=",
        "pageSize": 200,
        "page": 1,
        "taskState": "all",
        "idWorkflow": id_workflow,
    })
    if not data:
        return []
    return data.get("cases", {}).get("rows", [])


def get_render(driver, id_case, id_workitem, id_task):
    return api_post(driver, "/Rest/Handlers/Render", {
        "h_action":     "LOADFORM",
        "h_devicetype": "0",
        "h_devicecode": "1920x1080",
        "h_idCase":     str(id_case),
        "h_idWorkitem": str(id_workitem),
        "h_idTask":     str(id_task),
    })


def get_historico(driver, render_data):
    """Llama a MultiAction para traer el historico del caso."""
    form          = render_data.get("form", {})
    page_cache_id = form.get("pageCacheId", "")
    elements      = form.get("elements", [])

    hist_grid, parent_id = find_grid_with_parent(elements, "xHistoticosolicitud")
    if not hist_grid:
        print(" [hist: grid no encontrado]", end="")
        return []

    grid_id    = hist_grid.get("id", "")
    grid_xpath = hist_grid.get("xpath", "")
    # h_idRender debe ser el contenedor padre; h_tag es el grid mismo
    h_id_render = parent_id or grid_id

    actions = json.dumps([{
        "p_sort":         "idpActividad.sDescripcion",
        "p_order":        "asc",
        "p_page":         1,
        "p_rows":         100,
        "h_action":       "PROCESSPROPERTYVALUE",
        "h_xpath":        grid_xpath,
        "h_idRender":     h_id_render,
        "h_xpathContext": "",
        "h_pageCacheId":  page_cache_id,
        "h_propertyName": "data",
        "h_tag":          grid_id,
    }])

    resp = api_post(driver, "/Rest/Handlers/MultiAction", {
        "h_action":  "multiaction",
        "h_actions": actions,
    })

    if not resp or not isinstance(resp, list):
        return []

    rows = resp[0].get("result", {})
    if isinstance(rows, dict):
        return rows.get("rows", [])
    return []


def procesar_casos(driver):
    """Recorre todos los workflows y extrae la informacion de cada caso."""
    workflows = get_workflows(driver)
    if not workflows:
        print("No se encontraron workflows.")
        return {}

    secciones = {}

    for wf in workflows:
        nombre  = wf["name"]
        id_wf   = wf["idWorkFlow"]
        total   = wf["count"]
        print(f"\n{'='*52}")
        print(f"Workflow: {nombre}  ({total} casos)")
        print(f"{'='*52}")

        cases = get_cases_for_workflow(driver, id_wf)
        casos_procesados = []

        for case in cases:
            case_id    = case["id"]
            task_state = case["taskState"]
            fields     = case["fields"]

            numero_caso    = fields[0] if len(fields) > 0 else ""
            fecha_creacion = fields[3] if len(fields) > 3 else ""

            # Actividad y workitem desde fields[2]
            act_info   = fields[2] if len(fields) > 2 else {}
            workitems  = act_info.get("workitems", []) if isinstance(act_info, dict) else []
            actividad  = workitems[0].get("TaskName",  "") if workitems else ""
            id_workitem = workitems[0].get("idWorkItem", "") if workitems else ""
            id_task    = workitems[0].get("idTask",    "") if workitems else ""

            # Fecha vencimiento desde fields[4]
            vence_info  = fields[4] if len(fields) > 4 else {}
            fecha_vence = ""
            if isinstance(vence_info, dict):
                v = vence_info.get("Actividad vence en", [])
                fecha_vence = v[0] if v else ""

            print(f"  {numero_caso} [{task_state}]...", end="", flush=True)

            # Detalle via Render
            render = None
            if id_workitem and id_task:
                render = get_render(driver, case_id, id_workitem, id_task)

            estado_solicitud   = ""
            nombre_proyecto    = ""
            solicitantes       = []
            registro_solicitud = ""
            respuesta_obs      = ""

            if render:
                form_els = render.get("form", {}).get("elements", [])

                estado_solicitud = find_field(form_els, "idpEstadosolicitud.sDescripcion")

                nombre_proyecto = find_field(form_els, "sNombredelproyecto")

                # Solo extraer detalle completo para casos NO certificados
                if estado_solicitud == "Certificado":
                    print(" [Certificado]")
                    casos_procesados.append({
                        "Número del caso":         numero_caso,
                        "Proceso":                 nombre,
                        "Estado":                  task_state,
                        "Estado Solicitud":        estado_solicitud,
                        "Actividad":               actividad,
                        "Nombre del proyecto":     nombre_proyecto,
                        "Fecha creación caso":     fecha_creacion,
                        "Actividad vence en":      fecha_vence,
                        "Registro de Solicitud":   "",
                        "Respuesta Observaciones": "",
                        "Solicitantes":            "[]",
                        "_extraido":               datetime.now().strftime("%Y-%m-%d %H:%M"),
                    })
                    continue

                # Solicitantes asociados (grid xSolAsoc)
                sol_grid = find_grid(form_els, "xSolAsoc")
                if sol_grid:
                    for row in sol_grid.get("data", {}).get("rows", []):
                        solicitantes.append({
                            "tipo_persona": row[1] if len(row) > 1 else "",
                            "nombre":       row[2] if len(row) > 2 else "",
                            "tipo_id":      row[3] if len(row) > 3 else "",
                            "numero_id":    str(row[4]) if len(row) > 4 else "",
                            "municipio":    row[6] if len(row) > 6 else "",
                        })

                # Historico
                hist_rows = get_historico(driver, render)
                for hr in hist_rows:
                    act_hist   = hr[1] if len(hr) > 1 else ""
                    fecha_hist = hr[2] if len(hr) > 2 else ""
                    if act_hist == "Registro de Solicitud" and not registro_solicitud:
                        registro_solicitud = fecha_hist
                    elif act_hist == "Respuesta Observaciones" and not respuesta_obs:
                        respuesta_obs = fecha_hist

            caso = {
                "Número del caso":        numero_caso,
                "Proceso":                nombre,
                "Estado":                 task_state,
                "Estado Solicitud":       estado_solicitud,
                "Actividad":              actividad,
                "Nombre del proyecto":    nombre_proyecto,
                "Fecha creación caso":    fecha_creacion,
                "Actividad vence en":     fecha_vence,
                "Registro de Solicitud":  registro_solicitud,
                "Respuesta Observaciones": respuesta_obs,
                "Solicitantes":           json.dumps(solicitantes, ensure_ascii=False),
                "_extraido":              datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
            casos_procesados.append(caso)
            print(f" OK  [{estado_solicitud or '—'}]  {nombre_proyecto or '—'}")

        if casos_procesados:
            secciones[nombre] = casos_procesados
            print(f"  → {len(casos_procesados)} casos incluidos (de {total} totales)")

    return secciones


# ══════════════════════════════════════════
#  HTML DEL DASHBOARD
# ══════════════════════════════════════════

def generar_html(secciones):
    total = sum(len(v) for v in secciones.values())
    datos_json = json.dumps(secciones, ensure_ascii=False)

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
  tbody tr { border-bottom:1px solid #f0f2f0; transition:background .1s; cursor:pointer; }
  tbody tr:hover { background:var(--green-light); }
  tbody tr:last-child { border-bottom:none; }
  td { padding:9px 13px; vertical-align:middle; }
  .cid { font-weight:600; color:var(--green-dark); }
  .chip { display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; }
  .chip-red  { background:#fde8e8; color:#c0392b; }
  .chip-green{ background:#e8f5e9; color:#2e7d32; }
  .chip-gold { background:#fdf6e3; color:#b8860b; }
  .date { font-size:12px; color:var(--muted); white-space:nowrap; }
  .vence { font-size:12px; font-weight:500; white-space:nowrap; }
  .vence.v  { color:var(--red); font-weight:700; }
  .vence.p  { color:#e67e22; }
  .vence.ok { color:var(--muted); }

  .footer { padding:9px 22px; font-size:11.5px; color:var(--faint); border-top:1px solid var(--border); background:#fff; display:flex; justify-content:space-between; flex-shrink:0; }

  .overlay { display:none; position:fixed; inset:0; background:rgba(0,0,0,.4); z-index:200; align-items:center; justify-content:center; }
  .overlay.open { display:flex; }
  .modal { background:#fff; border-radius:10px; width:640px; max-width:94vw; max-height:90vh; overflow-y:auto; padding:26px; box-shadow:0 20px 60px rgba(0,0,0,.22); }
  .mh { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:18px; }
  .mt { font-size:16px; font-weight:700; color:var(--green-dark); }
  .ms { font-size:12px; color:var(--faint); margin-top:3px; }
  .mc { background:none; border:none; font-size:20px; cursor:pointer; color:var(--faint); }
  .mr { display:flex; gap:16px; margin-bottom:11px; flex-wrap:wrap; }
  .mf { flex:1; min-width:140px; }
  .mf label { font-size:10.5px; font-weight:700; color:var(--faint); text-transform:uppercase; letter-spacing:.5px; display:block; margin-bottom:3px; }
  .mf span { font-size:13px; font-weight:500; }
  .mdiv { height:1px; background:var(--border); margin:14px 0; }
  .sol-table { width:100%; border-collapse:collapse; font-size:12px; margin-top:6px; }
  .sol-table th { background:var(--green-light); padding:5px 8px; text-align:left; font-weight:600; color:var(--green-dark); }
  .sol-table td { padding:5px 8px; border-bottom:1px solid #eee; }
  ::-webkit-scrollbar { width:5px; height:5px; }
  ::-webkit-scrollbar-thumb { background:var(--border); border-radius:3px; }
"""

    js = """
const D = __DATOS__;
const hoy = new Date();
let SEC = "__todos__", BUS = "", FEST = "__todos__", _lista = [];

const $ = id => document.getElementById(id);

function todos() { return Object.values(D).flat(); }
function deSec(s) { return s === "__todos__" ? todos() : (D[s] || []); }

function parseFecha(s) {
  if (!s) return null;
  // formato MM/DD/YYYY HH:MM
  const m = s.match(/(\\d{1,2})\\/(\\d{1,2})\\/(\\d{4})\\s+(\\d{1,2}):(\\d{2})/);
  if (m) return new Date(+m[3], +m[1]-1, +m[2], +m[4], +m[5]);
  return null;
}

function clsV(s) {
  const f = parseFecha(s);
  if (!f) return "ok";
  const d = (f - hoy) / 86400000;
  return d < 0 ? "v" : d < 30 ? "p" : "ok";
}

function chipEstado(e) {
  if (!e) return "<span class='chip chip-gold'>—</span>";
  const map = {
    "Red":   "chip-red",
    "Green": "chip-green",
    "Yellow":"chip-gold",
    "Black": "chip-gold",
  };
  return `<span class='chip ${map[e]||"chip-gold"}'>${e}</span>`;
}

function renderSidebar() {
  let h = "";
  for (const [k, v] of Object.entries(D)) {
    const label = k.length > 30 ? k.slice(0,28)+"…" : k;
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

function poblarFiltroEstado() {
  const vals = [...new Set(todos().map(c => c["Estado Solicitud"] || "").filter(Boolean))].sort();
  const sel = $("filtroEstado");
  sel.innerHTML = "<option value='__todos__'>Todos los estados</option>"
    + vals.map(v => `<option value="${v}">${v}</option>`).join("");
}

function render() {
  const pp = +$("perPage").value;
  let casos = deSec(SEC);
  if (BUS)                     casos = casos.filter(c => JSON.stringify(c).toLowerCase().includes(BUS));
  if (FEST !== "__todos__")    casos = casos.filter(c => (c["Estado Solicitud"] || "") === FEST);
  _lista = casos;

  const venc = casos.filter(c => clsV(c["Actividad vence en"]||"") === "v").length;
  $("s-total").textContent    = todos().length;
  $("s-activos").textContent  = casos.length - venc;
  $("s-vencidos").textContent = venc;
  const ext = (casos[0] && casos[0]["_extraido"]) || "–";
  $("s-ext").textContent = ext;
  $("fext").textContent  = ext;
  $("finfo").textContent = `Mostrando ${Math.min(pp, casos.length)} de ${casos.length} casos`;

  const show  = casos.slice(0, pp);
  const tbody = $("tbody");

  if (!show.length) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:52px;color:#aaa">
      <div style="font-size:30px;margin-bottom:8px">📭</div>Sin casos para mostrar</td></tr>`;
    return;
  }

  tbody.innerHTML = show.map((c, i) => {
    const num  = c["Número del caso"]    || "–";
    const proy = c["Nombre del proyecto"]|| "–";
    const act  = c["Actividad"]          || "–";
    const fcr  = c["Fecha creación caso"]|| "–";
    const fv   = c["Actividad vence en"] || "–";
    const cv   = clsV(fv);
    const tag  = cv==="v" ? "⚠ " : cv==="p" ? "⏰ " : "";
    return `<tr onclick="openM(${i})">
      <td><span class="cid">${num}</span></td>
      <td>${proy}</td>
      <td style="font-size:12px">${act}</td>
      <td class="date">${fcr}</td>
      <td class="vence ${cv}">${tag}${fv}</td>
      <td>${chipEstado(c["Estado"])}</td>
      <td style="font-size:12px;color:var(--muted)">${c["Estado Solicitud"]||"—"}</td>
    </tr>`;
  }).join("");
}

function renderSolicitantes(jsonStr) {
  let arr;
  try { arr = JSON.parse(jsonStr); } catch(e) { return ""; }
  if (!arr || !arr.length) return "<em style='color:#aaa'>Sin solicitantes</em>";
  const rows = arr.map(s =>
    `<tr><td>${s.tipo_persona}</td><td>${s.nombre}</td><td>${s.tipo_id} ${s.numero_id}</td><td>${s.municipio}</td></tr>`
  ).join("");
  return `<table class="sol-table">
    <tr><th>Tipo</th><th>Nombre / Razón social</th><th>Identificación</th><th>Municipio</th></tr>
    ${rows}</table>`;
}

function openM(i) {
  const c = _lista[i]; if (!c) return;
  $("m-id").textContent    = c["Número del caso"]    || "–";
  $("m-proc").textContent  = c["Proceso"]             || "–";
  $("m-proy").textContent  = c["Nombre del proyecto"] || "–";
  $("m-act").textContent   = c["Actividad"]           || "–";
  $("m-est").textContent   = c["Estado Solicitud"]    || "–";
  $("m-fcr").textContent   = c["Fecha creación caso"] || "–";
  $("m-fv").textContent    = c["Actividad vence en"]  || "–";
  $("m-reg").textContent   = c["Registro de Solicitud"]     || "–";
  $("m-resp").textContent  = c["Respuesta Observaciones"]   || "–";
  $("m-sol").innerHTML     = renderSolicitantes(c["Solicitantes"] || "[]");
  $("overlay").classList.add("open");
}
function closeM() { $("overlay").classList.remove("open"); }

function exportCSV() {
  if (!_lista.length) return;
  const skip = ["_extraido", "Solicitantes"];
  const keys = Object.keys(_lista[0]).filter(k => !k.startsWith("_") && !skip.includes(k));
  const rows = [keys.join(","), ..._lista.map(c =>
    keys.map(k => `"${(c[k]||"").toString().replace(/"/g,'""')}"`).join(",")
  )];
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([rows.join("\\n")], {type:"text/csv;charset=utf-8"}));
  a.download = "upme_casos.csv"; a.click();
}

renderSidebar();
poblarFiltroEstado();
setSec("__todos__");
document.getElementById("gSearch").addEventListener("input", e => {
  BUS = e.target.value.toLowerCase(); render();
});
document.getElementById("filtroEstado").addEventListener("change", e => {
  FEST = e.target.value; render();
});
"""

    js_final = js.replace("__DATOS__", datos_json)

    html = (
        "<!DOCTYPE html>\n"
        "<html lang='es'>\n"
        "<head>\n"
        "<meta charset='UTF-8'>\n"
        "<title>UPME \u2013 Bandeja de Entrada</title>\n"
        "<link href='https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap' rel='stylesheet'>\n"
        "<style>" + css + "</style>\n"
        "</head>\n<body>\n"

        "<div class='topbar'>\n"
        "  <div class='logo'><span class='logo-badge'>UPME</span></div>\n"
        "  <button class='nav-btn active'>\U0001f4e5 Bandeja de entrada</button>\n"
        "  <div class='topbar-right'>\n"
        "    <input class='search' id='gSearch' placeholder='\U0001f50d  Buscar caso...'>\n"
        "    <div class='avatar'>SD</div>\n"
        "  </div>\n"
        "</div>\n"

        "<div class='layout'>\n"
        "  <div class='sidebar'>\n"
        "    <div class='sb-label'>Bandeja</div>\n"
        "    <div class='sb-item active' id='sb-todos' onclick=\"setSec('__todos__', this)\">\n"
        "      <span class='sb-name'>\U0001f4c2 Todos los casos</span>\n"
        "      <span class='badge' id='b-todos'>" + str(total) + "</span>\n"
        "    </div>\n"
        "    <div id='sb-secs'></div>\n"
        "  </div>\n"

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
        "          <select id='filtroEstado' class='tb-btn' style='cursor:pointer'></select>\n"
        "          <button class='tb-btn' onclick='exportCSV()'>\u2b07 CSV</button>\n"
        "        </div>\n"
        "      </div>\n"
        "      <div class='t-wrap'>\n"
        "        <table>\n"
        "          <thead><tr>\n"
        "            <th>N\u00famero del caso</th>\n"
        "            <th>Nombre del proyecto</th>\n"
        "            <th>Actividad actual</th>\n"
        "            <th>Fecha creaci\u00f3n</th>\n"
        "            <th>Vence en</th>\n"
        "            <th>Estado</th>\n"
        "            <th>Estado Solicitud</th>\n"
        "          </tr></thead>\n"
        "          <tbody id='tbody'><tr><td colspan='7' style='text-align:center;padding:40px;color:#aaa'>Cargando\u2026</td></tr></tbody>\n"
        "        </table>\n"
        "      </div>\n"
        "    </div>\n"

        "    <div class='footer'>\n"
        "      <span id='finfo'>\u2013</span>\n"
        "      <span>UPME \u00b7 Extra\u00eddo: <b id='fext'>\u2013</b></span>\n"
        "    </div>\n"
        "  </div>\n"
        "</div>\n"

        "<div class='overlay' id='overlay' onclick=\"if(event.target===this)closeM()\">\n"
        "  <div class='modal'>\n"
        "    <div class='mh'>\n"
        "      <div><div class='mt' id='m-id'>\u2013</div><div class='ms' id='m-proc'>\u2013</div></div>\n"
        "      <button class='mc' onclick='closeM()'>&#x2715;</button>\n"
        "    </div>\n"
        "    <div class='mr'>\n"
        "      <div class='mf'><label>Nombre del proyecto</label><span id='m-proy'>\u2013</span></div>\n"
        "      <div class='mf'><label>Estado Solicitud</label><span id='m-est'>\u2013</span></div>\n"
        "    </div>\n"
        "    <div class='mr'>\n"
        "      <div class='mf'><label>Actividad</label><span id='m-act'>\u2013</span></div>\n"
        "    </div>\n"
        "    <div class='mdiv'></div>\n"
        "    <div class='mr'>\n"
        "      <div class='mf'><label>Fecha creaci\u00f3n</label><span id='m-fcr'>\u2013</span></div>\n"
        "      <div class='mf'><label>Vence en</label><span id='m-fv'>\u2013</span></div>\n"
        "    </div>\n"
        "    <div class='mr'>\n"
        "      <div class='mf'><label>Registro de Solicitud</label><span id='m-reg'>\u2013</span></div>\n"
        "      <div class='mf'><label>Respuesta Observaciones</label><span id='m-resp'>\u2013</span></div>\n"
        "    </div>\n"
        "    <div class='mdiv'></div>\n"
        "    <div class='mf'><label>Solicitantes asociados</label><div id='m-sol'></div></div>\n"
        "  </div>\n"
        "</div>\n"

        "<script>\n" + js_final + "\n</script>\n"
        "</body>\n</html>\n"
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

    driver   = init_driver()
    secciones = {}

    try:
        login(driver)
        secciones = procesar_casos(driver)
    except Exception as e:
        import traceback
        print(f"\nError durante la extraccion: {e}")
        traceback.print_exc()
        try:
            driver.save_screenshot("upme_error.png")
            print("Screenshot guardado: upme_error.png")
        except: pass
    finally:
        driver.quit()
        print("\nNavegador cerrado.")

    if not secciones or not any(secciones.values()):
        print("Sin datos extraidos — revisa upme_error.png")
        return

    total = sum(len(v) for v in secciones.values())
    print(f"\nTotal casos procesados: {total}")

    html = generar_html(secciones)
    tmp  = Path(tempfile.gettempdir()) / "upme_dashboard.html"
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
