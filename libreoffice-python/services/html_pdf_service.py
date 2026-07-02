"""
services/html_pdf_service.py

Convierte un string HTML (o HTML en base64) a un archivo PDF
usando WeasyPrint, respetando estilos CSS inline y embebidos.
"""

import os
import base64
import uuid
import tempfile
from pathlib import Path

# WeasyPrint para conversión HTML → PDF de alta fidelidad
try:
    from weasyprint import HTML, CSS
    from weasyprint.text.fonts import FontConfiguration
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False


# Directorio de salida — mismo que el resto del proyecto
try:
    from utils.helpers import DATA_DIR
    OUTPUT_DIR = os.path.join(DATA_DIR, "generados")
except ImportError:
    OUTPUT_DIR = tempfile.gettempdir()

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ── CSS base opcional que se aplica a todos los documentos ──────────────────
_BASE_CSS = """
    @page {
        margin: 1.5cm;
    }
    body {
        font-family: Arial, Helvetica, sans-serif;
        font-size: 11pt;
        line-height: 1.4;
        color: #1a1a1a;
    }
    img {
        max-width: 100%;
    }
    table {
        border-collapse: collapse;
        width: 100%;
    }
    th, td {
        border: 1px solid #ccc;
        padding: 6px 10px;
        text-align: left;
    }
"""


def html_to_pdf(
    html_content: str,
    base_url: str | None = None,
    extra_css: str | None = None,
    filename_prefix: str = "documento",
) -> str:
    """
    Convierte un string HTML a PDF y lo guarda en OUTPUT_DIR.

    Parámetros
    ----------
    html_content : str
        HTML completo (con o sin <html>/<body>).
    base_url : str | None
        URL base para resolver recursos relativos (imágenes, CSS externo).
        Si es None se usa el directorio de trabajo.
    extra_css : str | None
        CSS adicional que se aplica sobre el base.
    filename_prefix : str
        Prefijo del nombre del archivo de salida (sin extensión).

    Retorna
    -------
    str
        Ruta absoluta al PDF generado.

    Lanza
    -----
    RuntimeError si WeasyPrint no está instalado.
    """
    if not WEASYPRINT_AVAILABLE:
        raise RuntimeError(
            "WeasyPrint no está instalado. "
            "Ejecutá: pip install weasyprint"
        )

    unique_id = uuid.uuid4().hex[:8]
    output_filename = f"{filename_prefix}_{unique_id}.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    font_config = FontConfiguration()

    css_sheets = [CSS(string=_BASE_CSS, font_config=font_config)]
    if extra_css:
        css_sheets.append(CSS(string=extra_css, font_config=font_config))

    html_obj = HTML(
        string=html_content,
        base_url=base_url or os.getcwd(),
    )
    html_obj.write_pdf(
        output_path,
        stylesheets=css_sheets,
        font_config=font_config,
    )

    return output_path


def html_base64_to_pdf(
    html_b64: str,
    base_url: str | None = None,
    extra_css: str | None = None,
    filename_prefix: str = "documento",
) -> str:
    """
    Decodifica un HTML en base64 y lo convierte a PDF.
    Delega en html_to_pdf() una vez decodificado.
    """
    try:
        html_content = base64.b64decode(html_b64).decode("utf-8")
    except Exception as exc:
        raise ValueError(f"No se pudo decodificar el HTML en base64: {exc}") from exc

    return html_to_pdf(
        html_content=html_content,
        base_url=base_url,
        extra_css=extra_css,
        filename_prefix=filename_prefix,
    )


def pdf_to_base64(pdf_path: str) -> str:
    """Lee un PDF del disco y lo devuelve como string base64."""
    with open(pdf_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")
