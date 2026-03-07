"""
Drive Extractor — MGS_0025_SDS4_El Copey Oc
============================================
Descarga los archivos del proyecto desde Google Drive a una carpeta local.
"""

import io
import os
from pathlib import Path
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ===========================================================================
# CONFIGURACIÓN
# ===========================================================================

_BASE = Path(__file__).parent
GOOGLE_TOKEN  = str(_BASE / "Recursos" / "token.json")
OUTPUT_DIR    = str(_BASE / "data" / "MGS_0025")

# IDs mapeados de la estructura real en Drive
FILES_TO_DOWNLOAD = [
    # (carpeta_local, file_id, nombre_destino, es_google_sheet)
    ("03_Engineering/07_Material_quantities",
     "1liBiqxjJAndlw8L83F2Q_YwcByopFn4l",
     "Cope-ELE-MC-CantidadesMaterial.xlsx", False),

    ("06_Financial",
     "1n8CnoHjUkzAwWJVlybEiEkpzAsSmlWi3Cpr-JDxGibg",
     "PEPC_Cope.xlsx", True),

    ("06_Financial/06_BT/01_Tracker_FNCE_20251551/01_Formato1",
     "1mQkM9KuJiGTqtcjqM6xJGxOpCW3ukxES",
     "01 CAMARA DE COMERCIO AYURA - 21112024.pdf", False),

    ("03_Engineering/05_Final version/03_Layouts 2D",
     "1FTBbTqFqGbnMZAfgSnbFP1oIQxeb8sFo",
     "SDS4_Cope-CIV-PL-01_Plano de localizaciones y accesos.pdf", False),

    ("03_Engineering/05_Final version/01_Simulation",
     "1b73jCj1AxdxaXsVgw7Zv890HtChJPGdd",
     "Cope-INF-ELE-V2.pdf", False),

    ("06_Financial/06_BT",
     "1QY-1WovvCffyNAWdwG7rkrplwgNXJuWp",
     "Copey definitivo.xlsx", False),
]

# Carpetas completas (descargar todos los archivos)
FOLDERS_TO_DOWNLOAD = [
    # (carpeta_local, folder_id)
    ("06_Financial/06_BT/01_Tracker_FNCE_20251551/02_Formato3",
     "1ma3q0lsHW18Puov90uc9vdrcW7kT69r4"),

    ("06_Financial/06_BT/03_Materiales_FNCE_20253284/02_Formato3",
     "19xhtAKqCmcPkUARBfuAai5yZcc-byLzI"),

    ("06_Financial/06_BT/04_Servicios_Cacica",
     "1rL8oDbKLloRnwHl1pbtqIJSLTcyVwmR2"),

    ("06_Financial/06_BT/Servicios Definitivos",
     "1zHnzEcKaNCulNZv-aN9ZIz3OugN_wIw4"),
]

GSHEET_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ===========================================================================
# FUNCIONES
# ===========================================================================

def build_service():
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN)
    return build("drive", "v3", credentials=creds)


def download_file(service, file_id: str, dest_path: str, is_gsheet: bool = False):
    Path(dest_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        if is_gsheet:
            request = service.files().export_media(fileId=file_id, mimeType=GSHEET_MIME)
        else:
            request = service.files().get_media(fileId=file_id)

        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        with open(dest_path, "wb") as f:
            f.write(buf.getvalue())
        print(f"  OK  {dest_path}")
        return True
    except Exception as e:
        print(f"  ERR {dest_path}: {e}")
        return False


def download_folder(service, folder_id: str, local_dir: str):
    resp = service.files().list(
        q=f"'{folder_id}' in parents",
        fields="files(id,name,mimeType)",
        pageSize=50
    ).execute()
    for f in resp.get("files", []):
        mime = f["mimeType"]
        is_gsheet = mime == "application/vnd.google-apps.spreadsheet"
        name = f["name"] if not is_gsheet else f["name"] + ".xlsx"
        dest = os.path.join(local_dir, name)
        download_file(service, f["id"], dest, is_gsheet)


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    print("Conectando a Google Drive...")
    service = build_service()
    print("Conectado.\n")

    print("=== Archivos individuales ===")
    for rel_dir, file_id, name, is_gsheet in FILES_TO_DOWNLOAD:
        dest = os.path.join(OUTPUT_DIR, rel_dir, name)
        download_file(service, file_id, dest, is_gsheet)

    print("\n=== Carpetas completas ===")
    for rel_dir, folder_id in FOLDERS_TO_DOWNLOAD:
        local_dir = os.path.join(OUTPUT_DIR, rel_dir)
        print(f"\n  Carpeta: {rel_dir}")
        download_folder(service, folder_id, local_dir)

    print(f"\nListo. Archivos guardados en: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
