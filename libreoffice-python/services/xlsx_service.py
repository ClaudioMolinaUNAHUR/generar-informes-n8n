import os
import uuid
import subprocess
from openpyxl import load_workbook
from utils.helpers import DATA_DIR, log


def _set_print_area(xlsx_path: str, cell_range: str, sheet_name: str = None) -> str:
    """
    Setea el área de impresión de la hoja indicada (o la activa) y configura
    el ajuste a una sola página. Devuelve el nombre de la hoja usada.
    """
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    ws.print_area = cell_range

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.save(xlsx_path)
    return ws.title


def convert_xlsx_to_pdf(xlsx_bytes: bytes, cell_range: str = "A1:J36", sheet_name: str = None) -> str:
    """
    Convierte un Excel (bytes crudos del archivo) a PDF, imprimiendo
    únicamente el rango de celdas indicado (por defecto A1:J36) de la hoja
    indicada (o la activa).

    Devuelve la ruta absoluta al PDF generado.
    """
    if not xlsx_bytes:
        raise ValueError("El archivo xlsx recibido está vacío")

    if not cell_range:
        cell_range = "A1:J36"

    output_dir = os.path.join(DATA_DIR, "generados")
    os.makedirs(output_dir, exist_ok=True)

    file_id = uuid.uuid4().hex
    xlsx_path = os.path.join(output_dir, f"{file_id}.xlsx")

    with open(xlsx_path, "wb") as f:
        f.write(xlsx_bytes)

    try:
        used_sheet = _set_print_area(xlsx_path, cell_range, sheet_name)
        log(f"📄 Print area '{cell_range}' seteado en hoja '{used_sheet}' → {xlsx_path}")

        cmd = [
            "libreoffice", "--headless", "--norestore",
            f"-env:UserInstallation=file:///tmp/lo_profile_{file_id}",
            "--convert-to", "pdf:calc_pdf_Export",
            "--outdir", output_dir,
            xlsx_path,
        ]
        log(f"🛠️ Ejecutando: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        log(f"↩️ returncode={result.returncode} stdout={result.stdout!r} stderr={result.stderr!r}")

        if result.returncode != 0:
            raise RuntimeError(f"Error convirtiendo xlsx a pdf: {result.stderr or result.stdout}")

        pdf_path = os.path.join(output_dir, f"{file_id}.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError(
                f"No se generó el PDF esperado tras la conversión. "
                f"stdout={result.stdout!r} stderr={result.stderr!r}"
            )

        log(f"✅ xlsx → pdf generado: {pdf_path}")
        return pdf_path
    finally:
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)