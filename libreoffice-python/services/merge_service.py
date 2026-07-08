import os
import io
import re
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


# Carpeta donde se guardan los comprobantes de orden de pago, relativa a DATA_DIR
ORDEN_DE_PAGO_SUBDIR = "orden-de-pago"

# Los archivos vienen nombrados como:
#   yyyy-mm_nombre_resto...pdf
#   yyyy-mm-dd_nombre_resto...pdf
_PATRON_NOMBRE_ARCHIVO = re.compile(
    r"^(?P<fecha>\d{4}-\d{2}(?:-\d{2})?)_(?P<nombre>[^_]+)(?:_(?P<resto>.*))?\.pdf$",
    re.IGNORECASE,
)


def _parsear_nombre_archivo(nombre_archivo: str):
    """
    Extrae (fecha, nombre, resto) del nombre de un archivo con el formato
    yyyy-mm_nombre_resto.pdf o yyyy-mm-dd_nombre_resto.pdf.
    Devuelve None si el archivo no matchea el patrón esperado.
    """
    match = _PATRON_NOMBRE_ARCHIVO.match(nombre_archivo)
    if not match:
        return None
    return match.group("fecha"), match.group("nombre"), match.group("resto") or ""


def buscar_pdfs_orden_de_pago(
    fecha: str = None,
    nombre: str = None,
    carpeta: str = ORDEN_DE_PAGO_SUBDIR,
) -> list:
    """
    Busca, dentro de DATA_DIR/<carpeta>, los PDFs cuyo nombre coincide con
    la fecha y/o el nombre indicados.

    - fecha: puede ser "yyyy-mm" o "yyyy-mm-dd". Si el archivo tiene fecha
      completa (yyyy-mm-dd) y se busca solo por "yyyy-mm", igual matchea
      (coincidencia por prefijo).
    - nombre: se compara sin importar mayúsculas/minúsculas contra el
      "nombre" que aparece justo después de la fecha en el archivo.

    Si un parámetro es None, no se filtra por ese criterio.
    Devuelve las rutas absolutas de los archivos encontrados, ordenadas
    por nombre de archivo (para que el merge quede en un orden consistente).
    """
    carpeta_completa = os.path.join(DATA_DIR, carpeta)
    if not os.path.isdir(carpeta_completa):
        raise ValueError(f"No existe la carpeta: {carpeta_completa}")

    nombre_normalizado = nombre.strip().lower() if nombre else None

    encontrados = []
    for archivo in os.listdir(carpeta_completa):
        if not archivo.lower().endswith(".pdf"):
            continue

        parsed = _parsear_nombre_archivo(archivo)
        if not parsed:
            # No matchea el patrón esperado (yyyy-mm_nombre... o yyyy-mm-dd_nombre...)
            continue

        fecha_archivo, nombre_archivo, _resto = parsed

        if fecha and not fecha_archivo.startswith(fecha):
            continue

        if nombre_normalizado and nombre_normalizado not in nombre_archivo.lower():
            continue

        encontrados.append(os.path.join(carpeta_completa, archivo))

    encontrados.sort()
    return encontrados


def merge_pdfs_por_coincidencia(
    fecha: str = None,
    nombre: str = None,
    output_name: str = None,
    carpeta: str = ORDEN_DE_PAGO_SUBDIR,
) -> str:
    """
    Busca en DATA_DIR/<carpeta> los PDFs que coincidan con la fecha y/o el
    nombre indicados, y los une en un único PDF (en orden alfabético de
    nombre de archivo).
    """
    if not fecha and not nombre:
        raise ValueError("Se debe indicar al menos 'fecha' o 'nombre' para buscar coincidencias")

    archivos = buscar_pdfs_orden_de_pago(fecha=fecha, nombre=nombre, carpeta=carpeta)

    if not archivos:
        raise ValueError(
            f"No se encontraron PDFs en '{carpeta}' que coincidan con fecha={fecha!r}, nombre={nombre!r}"
        )

    log(f"🔎 {len(archivos)} PDFs coinciden con fecha={fecha!r}, nombre={nombre!r}: {archivos}")

    return merge_pdfs_from_paths(archivos, output_name=output_name)


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