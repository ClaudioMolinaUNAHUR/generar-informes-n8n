import os
import io
import uuid
from pypdf import PdfReader, PdfWriter
from utils.helpers import DATA_DIR, log


def _build_output_path(output_name: str = None) -> str:
    output_dir = os.path.join(DATA_DIR, "generados")
    os.makedirs(output_dir, exist_ok=True)

    name = output_name or f"informe_unido_{uuid.uuid4().hex[:8]}"
    if not name.lower().endswith(".pdf"):
        name += ".pdf"

    return os.path.join(output_dir, name)


def merge_pdfs_from_bytes(pdfs_bytes: list, output_name: str = None) -> str:
    """
    Une una lista de PDFs recibidos como bytes crudos (archivos subidos
    directamente), en el mismo orden en que llegan, en un único PDF.
    Devuelve la ruta absoluta al PDF final.
    """
    if not pdfs_bytes or not isinstance(pdfs_bytes, list):
        raise ValueError("Se debe enviar una lista no vacía de archivos PDF")

    writer = PdfWriter()

    for i, pdf_bytes in enumerate(pdfs_bytes):
        try:
            if not pdf_bytes:
                raise ValueError("archivo vacío")
            reader = PdfReader(io.BytesIO(pdf_bytes))
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            raise ValueError(f"Error procesando el PDF en la posición {i}: {e}")

    output_path = _build_output_path(output_name)
    with open(output_path, "wb") as f:
        writer.write(f)

    log(f"📎 {len(pdfs_bytes)} PDFs unidos → {output_path}")
    return output_path


def merge_pdfs_from_paths(pdf_paths: list, output_name: str = None) -> str:
    """
    Une PDFs que ya existen en disco (rutas absolutas o relativas a DATA_DIR),
    en el orden en que llegan. Útil cuando los PDFs los generó el propio
    sistema y solo se pasa el path en vez del contenido.
    """
    if not pdf_paths or not isinstance(pdf_paths, list):
        raise ValueError("'pdf_paths' debe ser una lista no vacía de rutas")

    writer = PdfWriter()

    for path in pdf_paths:
        full_path = path if os.path.isabs(path) else os.path.join(DATA_DIR, path)
        if not os.path.isfile(full_path):
            raise ValueError(f"No existe el archivo PDF: {full_path}")
        reader = PdfReader(full_path)
        for page in reader.pages:
            writer.add_page(page)

    output_path = _build_output_path(output_name)
    with open(output_path, "wb") as f:
        writer.write(f)

    log(f"📎 {len(pdf_paths)} PDFs unidos → {output_path}")
    return output_path
