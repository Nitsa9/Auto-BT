"""
Módulo 4: Gestión de Firmas - FNCE
===================================
Recibe el documento .docx del Módulo 3 y gestiona las firmas según disponibilidad:

  ┌──────────────────────────────────────────────────────────────────────┐
  │  Firmante                   │ Firma en Drive │ Acción               │
  ├──────────────────────────────────────────────────────────────────────┤
  │  Solenium (CC 1152203750)   │      ✓         │ Descargar → insertar │
  │  COX ENERGY (CC 1026274625) │      ✗         │ Enviar correo        │
  │  Bancolombia (CC 79048722)  │      ✗         │ Enviar correo        │
  └──────────────────────────────────────────────────────────────────────┘

Flujo:
  Módulo 3 → [documento.docx] → Módulo 4 → [documento_firmado_parcial.docx]
                                          → correos enviados a firmantes externos
"""

import os
import sys

# Asegura que Python encuentre services/, utils/ y config/
# sin importar desde qué directorio se ejecute el script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import json
import logging
import shutil
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

try:
    from services.google_drive import GoogleDriveService
except ImportError:
    from services.google_drive_stub import GoogleDriveService

try:
    from services.gmail_service import GmailService
except ImportError:
    from services.gmail_stub import GmailService

from utils.docx_handler import DocxSignatureHandler
from utils.pdf_converter import PdfToImageConverter
from utils.email_templates import build_signature_request_email
from config.settings import Settings

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
)
logger = logging.getLogger("module4.signature_manager")


@dataclass
class Signatory:
    role: str
    cc_ce: str
    name: str
    email: Optional[str] = None
    drive_file_id: Optional[str] = None
    local_pdf_path: Optional[str] = None
    local_img_path: Optional[str] = None
    status: str = "pending"          # "signed" | "email_sent" | "error"
    email_message_id: Optional[str] = None
    error_detail: Optional[str] = None


@dataclass
class SignatureResult:
    success: bool
    output_docx: Optional[str] = None
    signed: list = field(default_factory=list)
    email_sent: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    summary: str = ""


class SignatureManager:
    # CCs/NITs cuyas firmas SÍ están en Drive (Solenium)
    SIGNATURES_IN_DRIVE: set = {"1152203750"}

    def __init__(self, settings: Optional[Settings] = None):
        self.settings = settings or Settings()
        self.drive = GoogleDriveService(
            credentials_path=self.settings.google_credentials_path,
            token_path=self.settings.google_token_path,
        )
        self.gmail = GmailService(
            credentials_path=self.settings.google_credentials_path,
            token_path=self.settings.google_token_path,
        )
        self.docx_handler = DocxSignatureHandler()
        self.pdf_converter = PdfToImageConverter()
        self.work_dir = Path(self.settings.signatures_work_dir)
        self.work_dir.mkdir(parents=True, exist_ok=True)

        if self.settings.signatures_available_in_drive:
            self.SIGNATURES_IN_DRIVE = set(
                self.settings.signatures_available_in_drive.split(",")
            )

    def process(self, input_docx: str, output_docx: Optional[str] = None) -> SignatureResult:
        result = SignatureResult(success=False)
        input_path = Path(input_docx)

        if not input_path.exists():
            logger.error(f"Documento no encontrado: {input_docx}")
            return result

        if output_docx is None:
            output_docx = str(input_path.parent / f"{input_path.stem}_firmado.docx")

        logger.info("=" * 60)
        logger.info("Módulo 4: Gestión de Firmas FNCE")
        logger.info("=" * 60)
        logger.info(f"Entrada : {input_docx}")
        logger.info(f"Salida  : {output_docx}")

        # Paso 1: Extraer firmantes
        logger.info("\n[1/3] Extrayendo firmantes...")
        signatories = self.docx_handler.extract_signatories(input_docx)
        if not signatories:
            logger.error("No se encontraron firmantes.")
            return result

        for s in signatories:
            logger.info(f"  → [{s.role}] {s.name} | CC/NIT: {s.cc_ce} | email: {s.email or 'N/A'}")

        # Paso 2: Clasificar y procesar
        logger.info("\n[2/3] Procesando firmantes...")
        for signatory in signatories:
            if signatory.cc_ce in self.SIGNATURES_IN_DRIVE:
                self._process_drive_signature(signatory)
            else:
                self._process_email_request(signatory, input_docx)

            if signatory.status == "signed":
                result.signed.append(signatory)
            elif signatory.status == "email_sent":
                result.email_sent.append(signatory)
            else:
                result.errors.append(signatory)

        # Paso 3: Insertar firmas disponibles
        logger.info("\n[3/3] Generando documento...")
        if result.signed:
            try:
                self.docx_handler.insert_signatures(
                    input_docx=input_docx,
                    output_docx=output_docx,
                    signatories=result.signed,
                )
                logger.info(f"  ✓ Documento guardado: {output_docx}")
            except Exception as exc:
                logger.error(f"  ✗ Error al insertar firmas: {exc}")
                for s in result.signed:
                    s.status = "error"
                    s.error_detail = str(exc)
                result.errors.extend(result.signed)
                result.signed.clear()
                output_docx = None
        else:
            shutil.copy2(input_docx, output_docx)
            logger.info(f"  ℹ Sin firmas de Drive disponibles. Documento copiado: {output_docx}")

        result.output_docx = output_docx
        result.success = len(result.errors) == 0
        self._log_summary(result)
        result.summary = self._build_summary(result)
        return result

    def _process_drive_signature(self, signatory: Signatory) -> None:
        logger.info(f"\n  [Drive] {signatory.name} (CC {signatory.cc_ce})")
        file_id = self._find_signature_in_drive(signatory)
        if not file_id:
            msg = f"PDF no encontrado en Drive para CC/NIT {signatory.cc_ce}"
            logger.warning(f"    ⚠ {msg}")
            signatory.status = "error"
            signatory.error_detail = msg
            return
        signatory.drive_file_id = file_id
        logger.info(f"    ✓ Encontrado en Drive: {file_id}")

        pdf_path = self._download_signature_pdf(signatory)
        if not pdf_path:
            msg = f"Error al descargar PDF para CC/NIT {signatory.cc_ce}"
            logger.error(f"    ✗ {msg}")
            signatory.status = "error"
            signatory.error_detail = msg
            return
        signatory.local_pdf_path = pdf_path
        logger.info(f"    ✓ PDF descargado: {pdf_path}")

        img_path = self.pdf_converter.convert(pdf_path, self.work_dir)
        if not img_path:
            msg = f"Error al convertir PDF a imagen para CC/NIT {signatory.cc_ce}"
            logger.error(f"    ✗ {msg}")
            signatory.status = "error"
            signatory.error_detail = msg
            return
        signatory.local_img_path = img_path
        logger.info(f"    ✓ Imagen lista: {img_path}")
        signatory.status = "signed"

    def _process_email_request(self, signatory: Signatory, docx_path: str) -> None:
        logger.info(f"\n  [Email] {signatory.name} (CC {signatory.cc_ce})")
        if not signatory.email:
            msg = f"Sin correo para CC/NIT {signatory.cc_ce}"
            logger.error(f"    ✗ {msg}")
            signatory.status = "error"
            signatory.error_detail = msg
            return

        email_req = build_signature_request_email(
            signatory=signatory,
            docx_path=docx_path,
            project_name=self.settings.project_name,
            sender_contact=self.settings.sender_contact_email,
        )

        try:
            response = self.gmail.send_email(
                to=email_req.to,
                subject=email_req.subject,
                body_html=email_req.body_html,
                attachment_path=email_req.attachment_path,
                attachment_name=email_req.attachment_name,
            )
            signatory.email_message_id = response.get("id")
            signatory.status = "email_sent"
            logger.info(f"    ✓ Correo enviado → {signatory.email} (id={signatory.email_message_id})")
        except Exception as exc:
            msg = f"Error al enviar correo a {signatory.email}: {exc}"
            logger.error(f"    ✗ {msg}")
            signatory.status = "error"
            signatory.error_detail = msg

    def _find_signature_in_drive(self, signatory: Signatory) -> Optional[str]:
        folder_id = self.settings.signatures_drive_folder_id
        cc = signatory.cc_ce
        for pattern in [f"firma_{cc}", f"signature_{cc}", f"{cc}_firma", cc]:
            query = f"name contains '{pattern}' and mimeType='application/pdf'"
            if folder_id:
                query += f" and '{folder_id}' in parents"
            files = self.drive.search_files(query=query, max_results=1)
            if files:
                return files[0]["id"]
        return None

    def _download_signature_pdf(self, signatory: Signatory) -> Optional[str]:
        if not signatory.drive_file_id:
            return None
        dest = str(self.work_dir / f"firma_{signatory.cc_ce}.pdf")
        try:
            self.drive.download_file(file_id=signatory.drive_file_id, destination_path=dest)
            if not self._is_valid_pdf(dest):
                logger.error(f"Archivo no es PDF válido: {dest}")
                os.remove(dest)
                return None
            return dest
        except Exception as exc:
            logger.error(f"Error en descarga: {exc}")
            return None

    @staticmethod
    def _is_valid_pdf(path: str) -> bool:
        try:
            with open(path, "rb") as f:
                return f.read(5) == b"%PDF-"
        except OSError:
            return False

    @staticmethod
    def _log_summary(result: SignatureResult) -> None:
        logger.info("\n" + "=" * 60)
        logger.info("RESUMEN")
        logger.info("=" * 60)
        for s in result.signed:
            logger.info(f"  ✓ [FIRMADO]       {s.name} (CC {s.cc_ce})")
        for s in result.email_sent:
            logger.info(f"  📧 [CORREO ENV.]  {s.name} (CC {s.cc_ce}) → {s.email}")
        for s in result.errors:
            logger.warning(f"  ✗ [ERROR]         {s.name} (CC {s.cc_ce}): {s.error_detail}")
        logger.info(f"\n  Documento salida: {result.output_docx}")

    @staticmethod
    def _build_summary(result: SignatureResult) -> str:
        lines = []
        if result.signed:
            lines.append(f"Firmas insertadas ({len(result.signed)}): " +
                         ", ".join(s.name for s in result.signed))
        if result.email_sent:
            lines.append(f"Correos enviados ({len(result.email_sent)}): " +
                         ", ".join(f"{s.name} → {s.email}" for s in result.email_sent))
        if result.errors:
            lines.append(f"Errores ({len(result.errors)}): " +
                         ", ".join(f"{s.name}: {s.error_detail}" for s in result.errors))
        return " | ".join(lines)


def main():
    # ----------------------------------------------------------------
    # Ruta del documento entregado por el Módulo 3
    # Cambiar este path cuando el archivo se mueva de ubicación
    # ----------------------------------------------------------------
    INPUT_DOCX = r"C:\Users\naty-\OneDrive\Escritorio\Hackaton\Plantilla de Firma FNCE (17).docx"

    manager = SignatureManager(settings=Settings())
    result = manager.process(input_docx=INPUT_DOCX)

    print(json.dumps({
        "success": result.success,
        "output_docx": result.output_docx,
        "summary": result.summary,
        "signed": [{"name": s.name, "cc_ce": s.cc_ce} for s in result.signed],
        "email_sent": [{"name": s.name, "cc_ce": s.cc_ce, "email": s.email,
                        "message_id": s.email_message_id} for s in result.email_sent],
        "errors": [{"name": s.name, "cc_ce": s.cc_ce, "detail": s.error_detail}
                   for s in result.errors],
    }, indent=2, ensure_ascii=False))
    sys.exit(0 if result.success else 1)


if __name__ == "__main__":
    main()