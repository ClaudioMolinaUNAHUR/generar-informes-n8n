#!/usr/bin/env python3
import sys
import json
import base64
import subprocess
import os
import warnings
import uuid
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import matplotlib.pyplot as plt
from pypdf import PdfReader, PdfWriter
import requests
import textwrap
import numpy as np
from io import BytesIO

warnings.filterwarnings("ignore")
DATA_DIR = "/data"


# --------------------------------------------------------------
# UTILS
# --------------------------------------------------------------
def get_logo_from_base64(base64_string: str) -> BytesIO | None:
    """
    Decodifica una cadena Base64 y devuelve un objeto BytesIO con los datos de la imagen.
    Retorna None si la cadena está vacía o es inválida.
    """
    if not base64_string:
        return None
    try:
        image_data = base64.b64decode(base64_string)
        return BytesIO(image_data)
    except base64.binascii.Error:
        log("⚠️ Error de decodificación Base64. La cadena del logo podría ser inválida.")
        return None


def replace_placeholders(slide, replacements):
    """
    Busca un placeholder por su nombre (definido en el Panel de Selección de PowerPoint)
    y reemplaza su texto por el valor correspondiente.
    Si el valor es una lista de objetos, aplica formato (ej. negritas) creando runs.
    """
    
    for key, value in replacements.items():
        for shape in slide.shapes:
            if shape.has_text_frame:
                texto = shape.text.strip()
                if texto == key:
                    tf = shape.text_frame
                    tf.clear()  # Limpiar el contenido existente

                    if isinstance(value, list):
                        # Procesar como array de objetos con formato
                        for item in value:
                            if not isinstance(item, dict):
                                continue  # Saltar si no es dict
                            text_part = item.get("text", "")
                            if not tf.paragraphs:
                                p = tf.add_paragraph()
                            else:
                                p = tf.paragraphs[-1]
                            run = p.add_run()
                            run.text = text_part
                            if item.get("bold", False):
                                run.font.bold = True
                            # Aquí puedes agregar más formatos si es necesario, ej.:
                            # if item.get("italic", False):
                            #     run.font.italic = True
                    else:
                        # Tratar como string simple (comportamiento original)
                        if not isinstance(value, str):
                            val_str = str(value)
                        else:
                            val_str = value
                        val_str = val_str.replace("\\n", "\n").replace("\\\n", "\n")
                        tf.text = val_str

                    # Justificar el texto cuando sea contenido de texto largo (resumen/KPIs)
                    # Esto aplica al texto ya insertado y a cada párrafo existente.
                    for p in tf.paragraphs:
                        p.alignment = PP_ALIGN.JUSTIFY


def insert_image_scaled_by_width(slide, placeholder, image_path_or_stream):
    """
    Reemplaza un placeholder con una imagen, escalándola para que ocupe todo el ancho
    del placeholder y ajustando el alto proporcionalmente. La imagen se centra verticalmente.
    """
    parent = placeholder.element.getparent()
    if parent is None:
        return  # ya fue eliminado, evitar crash

    ph_left, ph_top = placeholder.left, placeholder.top
    ph_width, ph_height = placeholder.width, placeholder.height

    parent.remove(placeholder.element)

    pic = slide.shapes.add_picture(
        image_path_or_stream, ph_left, ph_top, width=ph_width
    )

    new_height = pic.height
    pic.top = ph_top + (ph_height - new_height) // 2


def insert_logo_preserving_aspect(slide, placeholder, logo_stream):
    # área disponible
    ph_left, ph_top = placeholder.left, placeholder.top
    ph_w, ph_h = placeholder.width, placeholder.height

    # borrar placeholder original
    placeholder.element.getparent().remove(placeholder.element)

    # insertar imagen sin escalar
    pic = slide.shapes.add_picture(logo_stream, ph_left, ph_top)

    # tamaño real de la imagen
    img_w, img_h = pic.width, pic.height

    # calcular factor manteniendo aspecto
    scale = min(ph_w / img_w, ph_h / img_h)

    # aplicar tamaño escalado
    new_w = int(img_w * scale)
    new_h = int(img_h * scale)

    pic.width = new_w
    pic.height = new_h

    # centrar dentro del placeholder original
    pic.left = ph_left + (ph_w - new_w) // 2
    pic.top = ph_top + (ph_h - new_h) // 2


def _insert_logo_with_scaling(slide, logo_stream):
    """
    Busca el primer placeholder de tipo 'Picture' (18) e inserta el logo
    dentro de sus límites, manteniendo la relación de aspecto y eliminando el placeholder original.
    """
    # 18 es el tipo de placeholder para 'Picture'
    LOGO_PLACEHOLDER_TYPE = 18

    if not logo_stream:
        return

    for shape in slide.placeholders:
        if shape.placeholder_format.type == LOGO_PLACEHOLDER_TYPE:

            insert_logo_preserving_aspect(slide, shape, logo_stream)
            break  # Solo insertamos el primer logo que encontremos


def log(msg):
    print(msg, flush=True)


def set_placeholder_text(slide, idx, text):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            if ph.has_text_frame:
                ph.text = text
            return
    # Si no existe, opcionalmente crear un cuadro de texto
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    tf = tb.text_frame
    tf.text = text


# --------------------------------------------------------------
# PORTADA
# --------------------------------------------------------------
def generar_portada(data, logo_stream):
    prs = Presentation(f"{DATA_DIR}/plantillas/plantilla_portada.pptx")
    slide = prs.slides[0]

    replacements = {
        "{{ph_titulo}}": data.get("titulo_portada", ""),
        "{{ph_subtitle}}": data.get("subtitulo_portada", ""),
        "{{ph_fecha}}": data.get("fecha_portada", ""),
        "{{ph_pie_l}}": data.get("pie_l", ""),
        "{{ph_pie_r}}": data.get("pie_r", ""),
    }

    replace_placeholders(slide, replacements)

    # Busca un placeholder de tipo imagen (18) para el logo.
    _insert_logo_with_scaling(slide, logo_stream)

    output = f"{DATA_DIR}/pptx-parts/portada.pptx"
    prs.save(output)
    return output


# --------------------------------------------------------------
# GRÁFICOS
# --------------------------------------------------------------
def create_matplotlib_chart(chart_info, friendly_names, output_file):
    plt.figure(figsize=(10, 5))
    ctype = chart_info.get("type")

    # Título si viene en la definición del gráfico
    # title = chart_info.get("title") or chart_info.get("titulo") or ""
    # if title:
    #     plt.title(title, fontsize=20, fontweight="bold", pad=20)

    # Aumentar tamaño de fuente para los ejes
    plt.xticks(fontsize=18)
    plt.yticks(fontsize=18)

    labels = chart_info.get("labels", [])
    x = range(len(labels))

    # Mapa de etiquetas amigables para claves comunes
    flat_friendly_names = {
        key: value for chart in friendly_names.values() for key, value in chart.items()
    }

    # Paleta de colores para series (se rotan si hay más series)
    palette = ["#4f81bd", "#9abb59", "#4bacc6", "#8064a2"]

    # Detectar series numéricas dinámicamente (manteniendo el orden del dict)
    series_keys = []
    for key, val in chart_info.items():
        if key in ("labels", "type", "title", "titulo"):
            continue
        # Considerar series que sean listas/tuplas de números
        if isinstance(val, (list, tuple)) and all(
            isinstance(v, (int, float)) for v in val
        ):
            series_keys.append(key)

    if ctype == "bar":
        # Barras agrupadas: calcular offsets según cantidad de series
        n = len(series_keys)
        if n == 0:
            # nada que dibujar
            return

        ind = np.arange(len(labels))  # posiciones base
        total_width = 0.7
        bar_width = total_width / n

        for idx, key in enumerate(series_keys):
            vals = list(chart_info.get(key) or [])
            # normalizar longitud de vals
            if len(vals) < len(labels):
                vals = vals + [0] * (len(labels) - len(vals))

            label_full = flat_friendly_names.get(key, key.replace("_", " ").capitalize())
            # Dividir etiquetas largas en varias líneas para que no encojan el gráfico
            label = textwrap.fill(label_full, width=22)

            color = palette[idx % len(palette)]

            # calcular posiciones para esta serie
            offset = (idx - (n - 1) / 2) * bar_width
            positions = ind + offset
            plt.bar(positions, vals, bar_width * 0.95, label=label, color=color)

        # Ajustar ticks al centro de los grupos
        plt.xticks(ind, labels, rotation=0, fontsize=16)
        plt.grid(axis="y", linestyle="-", color="#dcdcdc", linewidth=0.8)
        # Leyenda a la derecha, centrada verticalmente
        plt.legend(
            loc="center left",
            bbox_to_anchor=(1, 0.5),
            frameon=False,
            labelspacing=1.2,
            fontsize=18,
        )

    elif ctype == "line":
        for idx, key in enumerate(series_keys):
            vals = list(chart_info.get(key) or [])
            if len(vals) < len(labels):
                vals = vals + [None] * (len(labels) - len(vals))
            label_full = flat_friendly_names.get(key, key.replace("_", " ").capitalize())
            # Dividir etiquetas largas en varias líneas
            label = textwrap.fill(label_full, width=22)

            color = palette[idx % len(palette)]
            plt.plot(x, vals, label=label, marker="o", color=color)

        plt.xticks(x, labels, rotation=45, fontsize=16)
        plt.grid(axis="y", linestyle="-", color="#dcdcdc", linewidth=0.8)
        plt.legend(loc="best", fontsize=18)

    # Formato eje Y con separador de miles
    try:
        import matplotlib.ticker as mtick
        ax = plt.gca()
        
        # ESTA ES LA LÍNEA CLAVE: Fuerza a que los ticks sean solo números enteros
        ax.yaxis.set_major_locator(mtick.MaxNLocator(integer=True))
        
        # Tu formateador actual
        ax.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x, pos: f"{int(x):,}"))
    except Exception:
        pass

    # Ajusta el layout para asegurar que la leyenda no se corte
    plt.tight_layout(rect=[0, 0.03, 0.95, 0.97])
    plt.savefig(output_file, dpi=150, transparent=True)
    plt.close()


def add_charts(slide, charts, friendly_names, replacements_chart):
    # debug: enumera shape names y contenido para analizar placeholders y textos
    # for shape in slide.shapes:
    #     text = getattr(shape, "text", "")
    #     log(
    #         f"placeholder: '{shape.name}', type={shape.shape_type},"
    #         f" pos=({shape.left},{shape.top}), size=({shape.width},{shape.height}), text='{text.strip()}'"
    #     )

    shape_map = {s.name: s for s in slide.shapes}

    # Tomamos placeholders en el orden deseado de replacements_chart.
    placeholder_candidates = [shape_map[name] for name in replacements_chart if name in shape_map]

    # Asegurar ordenando por posición visual (top, left) como fallback
    placeholder_candidates = sorted(placeholder_candidates, key=lambda s: (s.top, s.left))

    # Tomar solo los placeholders necesarios para los charts disponibles
    chart_placeholders = placeholder_candidates[: len(charts)]

    for (name, chart_info), placeholder in zip(charts.items(), chart_placeholders):
        if not chart_info.get("title") and not chart_info.get("titulo") and name:
            chart_info["title"] = name.replace("_", " ").capitalize()

        # fn = os.path.join(DATA_DIR, f"{name}.png")
        fn = f"/tmp/{name}.png"
        create_matplotlib_chart(chart_info, friendly_names, fn)
        insert_image_scaled_by_width(slide, placeholder, fn)


# --------------------------------------------------------------
# CONTENIDO
# --------------------------------------------------------------


def generar_contenido(data, logo_stream):
    slides_data = data.get("slides", [])
    generated_files = []

    feet_l, feet_r, periodo = data.get("pie_l", ""), data.get("pie_r", ""), data.get("periodo", "")

    for i, slide_item in enumerate(slides_data):
        template_file = slide_item.get("file_slide", "plantilla_contenido.pptx")

        prs_resume = Presentation(f"{DATA_DIR}/plantillas/resume_{template_file}")
        prs = Presentation(f"{DATA_DIR}/plantillas/{template_file}")

        # Asumimos que la plantilla tiene una sola diapositiva
        slide_resume = prs_resume.slides[0]
        slide = prs.slides[0]  

        slide_content = slide_item.get("slide", {})

        # Cargamos los nombres amigables para las leyendas de los gráficos
        product_type = slide_item.get("type")
        friendly_names = {}
        try:
            with open(
                f"{DATA_DIR}/charts/chart_{product_type}.json", "r", encoding="utf-8"
            ) as f:
                friendly_names = json.load(f)
        except FileNotFoundError:
            log(
                f"⚠️  No se encontró el archivo de configuración de gráficos: chart_{product_type}.json"
            )

        # Diccionario de reemplazos
        replacements_resume = {
            "{{ph_titulo}}": slide_content.get("titulo", ""),
            "{{ph_resumen}}": slide_content.get("resumen", ""),
            "{{ph_pie_l}}": feet_l,
            "{{ph_pie_r}}": feet_r,
        }
        replacements = {
            "{{ph_titulo}}": slide_content.get("titulo", ""),
            "{{ph_periodo}}": periodo,
            "{{ph_title_1}}": slide_content.get("title_1", ""),
            "{{ph_kpis_1}}": slide_content.get("kpis_1", ""),
            "{{ph_title_2}}": slide_content.get("title_2", ""),
            "{{ph_kpis_2}}": slide_content.get("kpis_2", ""),
            "{{ph_title_3}}": slide_content.get("title_3", ""),
            "{{ph_kpis_3}}": slide_content.get("kpis_3", ""),
            "{{ph_title_4}}": slide_content.get("title_4", ""),
            "{{ph_kpis_4}}": slide_content.get("kpis_4", ""),
            "{{ph_pie_l}}": feet_l,
            "{{ph_pie_r}}": feet_r,
        }
        charts = slide_content.get("charts", {})
        # if product_type != "wazuh" and charts:
        #     replacements["{{ph_utilizacion}}"] = slide_content.get("kpi_title", "")
        #     replacements["{{ph_kpis}}"] = slide_content.get("kpis", "")
        #     replacements["{{ph_soporte}}"] = slide_content.get("soporte_title", "")
        #     replacements["{{ph_soporte_kpis}}"] = slide_content.get("soporte_kpi", "")

        # Reemplazar texto marcador
        replace_placeholders(slide_resume, replacements_resume)
        replace_placeholders(slide, replacements)
        replacements_chart = [
            "Marcador de posición de imagen 6",
            "Marcador de posición de imagen 7",
            "Marcador de posición de imagen 16",
            "Marcador de posición de imagen 17",
            "Marcador de posición de imagen 18",
        ]
        # Insertar gráficos
        if charts:
            add_charts(slide, charts, friendly_names, replacements_chart)

        _insert_logo_with_scaling(slide, logo_stream)
        _insert_logo_with_scaling(slide_resume, logo_stream)

        output_path_resume = f"{DATA_DIR}/pptx-parts/contenido_{product_type}_resume.pptx"
        output_path = f"{DATA_DIR}/pptx-parts/contenido_{product_type}.pptx"
        prs_resume.save(output_path_resume)
        prs.save(output_path)
        generated_files.append(output_path_resume)
        generated_files.append(output_path)
    return generated_files


# --------------------------------------------------------------
# CIERRE
# --------------------------------------------------------------
def generar_cierre(data, logo_stream):
    cierre = data["despedida"]
    prs = Presentation(f"{DATA_DIR}/plantillas/plantilla_cierre.pptx")
    slide = prs.slides[0]

    replacements = {
        "{{ph_titulo}}": cierre.get("titulo", ""),
        "{{ph_pie_l}}": data.get("pie_l", ""),
        "{{ph_pie_r}}": data.get("pie_r", ""),
    }

    replace_placeholders(slide, replacements)

    _insert_logo_with_scaling(slide, logo_stream)

    output = f"{DATA_DIR}/pptx-parts/cierre.pptx"
    prs.save(output)
    return output


# --------------------------------------------------------------
# PPTX → PDF
# --------------------------------------------------------------
def convert_to_pdf(pptx_file):
    output_dir = f"{DATA_DIR}/pdf-parts"
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.basename(pptx_file).replace(".pptx", ".pdf")
    pdf_file = os.path.join(output_dir, base_name)

    # Usar un directorio de instalación único para evitar bloqueos y problemas de permisos
    user_inst = f"-env:UserInstallation=file:///tmp/lo_{uuid.uuid4()}"
    cmd = [
        "libreoffice",
        user_inst,
        "--headless",
        "--convert-to",
        "pdf",
        pptx_file,
        "--outdir",
        output_dir,
    ]
    try:
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
    except subprocess.CalledProcessError as e:
        log(f"⚠️ Error en LibreOffice: {e.stderr.decode('utf-8', errors='replace')}")
        raise
    return pdf_file


def apply_background_to_pdf(content_pdf_path, background_pdf_path):
    """
    Aplica un fondo desde un PDF a otro PDF que contiene el contenido.
    El contenido se superpone sobre el fondo.
    """
    content_reader = PdfReader(content_pdf_path)
    background_reader = PdfReader(background_pdf_path)
    writer = PdfWriter()

    # Asume que el PDF de fondo tiene al menos tantas páginas como el de contenido
    for i, content_page in enumerate(content_reader.pages):
        # Obtiene la página de fondo correspondiente
        background_page = background_reader.pages[i % len(background_reader.pages)]
        # Superpone el contenido (que tiene fondo transparente/blanco) sobre el fondo
        background_page.merge_page(content_page)
        writer.add_page(background_page)

    with open(content_pdf_path, "wb") as f:
        writer.write(f)


# --------------------------------------------------------------
# UNIR PDFs
# --------------------------------------------------------------
def unir_pdfs(pdf_paths, empresa, type="", split=0):
    writer = PdfWriter()
    for pdf_path in pdf_paths:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)

    output_dir = f"{DATA_DIR}/generados"
    os.makedirs(output_dir, exist_ok=True)
    out = f"{output_dir}/informe_{empresa}{'.' + type if split == 1 else ''}.pdf"
    with open(out, "wb") as f:
        writer.write(f)
    return out


# --------------------------------------------------------------
# MAIN
# --------------------------------------------------------------
def main():
    raw = sys.argv[1]
    data = json.loads(base64.b64decode(raw))
    # with open(f"{DATA_DIR}/input.json", "r", encoding="utf-8") as f:
    #     data = json.load(f)
    data = data["data"]
    
    split = data.get("split")
    logo_stream = get_logo_from_base64(data.get("logo_base64"))

    empresa = data.get("logo")[:-4].lower() if data.get("logo") else ""

    portada = None
    cierre = None
    if data["save"]:
        portada = generar_portada(data, logo_stream)
        cierre = generar_cierre(data, logo_stream)
    contenido_files = generar_contenido(data, logo_stream)
    types = [slide.get("type", "") for slide in data.get("slides", [])]

    # Pre-convert common parts to avoid redundant processing
    portada_pdf = convert_to_pdf(portada) if (data["save"] and portada) else None
    cierre_pdf = convert_to_pdf(cierre) if (data["save"] and cierre) else None

    informe_name = []
    if split == 0:
        pdf_files_to_merge = []
        if portada_pdf:
            pdf_files_to_merge.append(portada_pdf)

        pdf_files_to_merge.extend([convert_to_pdf(f) for f in contenido_files])

        if cierre_pdf:
            pdf_files_to_merge.append(cierre_pdf)

        informe_name.append(unir_pdfs(pdf_files_to_merge, empresa))
    else:
        for idx, content_pptx in enumerate(contenido_files):

            pdf_files_to_merge = []
            if portada_pdf:
                pdf_files_to_merge.append(portada_pdf)

            pdf_files_to_merge.append(convert_to_pdf(content_pptx))

            if cierre_pdf:
                pdf_files_to_merge.append(cierre_pdf)

            informe_name.append(
                unir_pdfs(pdf_files_to_merge, empresa, types[idx], split)
            )
            # Aplicar fondo al PDF de contenido

    # with open(final_pdf, "rb") as f:
    #     b64 = base64.b64encode(f.read()).decode()

    print(json.dumps({"file_names": informe_name}))


if __name__ == "__main__":
    main()
