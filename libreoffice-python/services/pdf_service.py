import os
import json
import uuid
import subprocess
from pptx import Presentation
from pypdf import PdfReader, PdfWriter
from utils.helpers import (
    DATA_DIR,
    log,
    _insert_logo_with_scaling,
    apply_text_formatting,
    normalize_kpi_text,
    normalize_suggestion_text,
    insert_image_scaled_by_width,
)
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Pt
import re
from services.chart_service import create_matplotlib_chart


def replace_placeholders(slide, replacements):
    """
    Busca un placeholder por su nombre (definido en el Panel de Selección de PowerPoint)
    y reemplaza su texto por el valor correspondiente.
    Si el valor es una lista de objetos, aplica formato (ej. negritas) creando runs.
    Retorna una lista de tuplas (text_frame, key) modificados.
    """
    modified = []

    def _value_to_text(value):
        if isinstance(value, list):
            parts = []
            for item in value:
                if isinstance(item, dict):
                    parts.append(str(item.get("text", "")))
                else:
                    parts.append(str(item))
            return "".join(parts)
        if isinstance(value, str):
            return value
        return str(value)

    # 1) Reemplazo exacto: mantiene compatibilidad con placeholders que ocupan todo el cuadro.
    for key, value in replacements.items():
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

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
                    val_str = _value_to_text(value)
                    val_str = val_str.replace("\\n", "\n").replace("\\\n", "\n")
                    tf.text = val_str

                for p in tf.paragraphs:
                    p.alignment = PP_ALIGN.JUSTIFY

                modified.append((tf, key))

    # 2) Reemplazo embebido: soporta placeholders dentro de un texto mayor en el mismo cuadro.
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        tf = shape.text_frame
        original_text = tf.text or ""
        new_text = original_text
        replaced_keys = []

        for key, value in replacements.items():
            if key in new_text:
                val_str = _value_to_text(value)
                val_str = val_str.replace("\\n", "\n").replace("\\\n", "\n")
                new_text = new_text.replace(key, val_str)
                replaced_keys.append(key)

        if new_text != original_text:
            tf.text = new_text
            for p in tf.paragraphs:
                p.alignment = PP_ALIGN.JUSTIFY
            for key in replaced_keys:
                modified.append((tf, key))

    # 3) Limpieza final: si queda algún placeholder sin reemplazar, ocultarlo con blanco.
    unresolved_pattern = re.compile(r"\{\{ph_[^}]+\}\}")
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        original_text = tf.text or ""
        cleaned_text = unresolved_pattern.sub(" ", original_text)
        if cleaned_text != original_text:
            tf.text = cleaned_text

    return modified


def add_charts(slide, charts, friendly_names, replacements_chart):
    to_sort = [s for s in slide.shapes if s.name in replacements_chart]
    chart_placeholders = sorted(to_sort, key=lambda s: s.name)

    for i, (name, chart_info) in enumerate(charts.items()):
        if i >= len(chart_placeholders):
            break
        placeholder = chart_placeholders[i]
        if not chart_info.get("title") and not chart_info.get("titulo") and name:
            chart_info["title"] = name.replace("_", " ").capitalize()

        fn = f"/tmp/{name}_{uuid.uuid4().hex}.png"
        create_matplotlib_chart(chart_info, friendly_names, fn)
        insert_image_scaled_by_width(slide, placeholder, fn)


def convert_to_pdf(pptx_file: str) -> str:
    output_dir = f"{DATA_DIR}/pdf-parts"
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.basename(pptx_file).replace(".pptx", ".pdf")
    pdf_file = os.path.join(output_dir, base_name)

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
        raise Exception(f"LibreOffice error: {e.stderr.decode()}")

    return pdf_file


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

    # Ejecutamos el reemplazo normal
    modified = replace_placeholders(slide, replacements)

    for tf, key in modified:
        key_l = key.lower()
        if "titulo" in key_l:
            apply_text_formatting(tf, font_name="Calibri", size=18)
        elif "sub" in key_l:
            apply_text_formatting(tf, font_name="Calibri", size=14)
        elif "pie" in key_l:
            apply_text_formatting(tf, font_name=None, size=10)
        else:
            apply_text_formatting(tf, font_name=None, size=12)

    # Aplicamos el centrado manual a los campos específicos
    campos_a_centrar = ["{{ph_subtitle}}", "{{ph_fecha}}"]

    for shape in slide.shapes:
        if shape.has_text_frame:
            # Buscamos si el texto de la forma coincide con los valores ya reemplazados
            # (o puedes volver a iterar sobre las keys de replacements)
            for key in campos_a_centrar:
                valor_insertado = replacements[key]
                if valor_insertado and valor_insertado in shape.text:
                    for paragraph in shape.text_frame.paragraphs:
                        paragraph.alignment = PP_ALIGN.CENTER

    # Busca un placeholder de tipo imagen (18) para el logo
    _insert_logo_with_scaling(slide, logo_stream)

    output = f"{DATA_DIR}/pptx-parts/portada.pptx"
    prs.save(output)
    return output


def generar_contenido_slide(slide_item, data, logo_stream):
    slide_content = slide_item.get("slide", {})

    charts = slide_content.get("charts", {})
    num_charts = len(charts)
    if num_charts == 0:
        # No crear la hoja si no hay gráficos
        return None
    elif num_charts == 1:
        template_file = "plantilla_contenido1.pptx"
    elif num_charts == 2:
        template_file = "plantilla_contenido2.pptx"
    elif num_charts == 3:
        template_file = "plantilla_contenido3.pptx"
    else:
        # Si hay más de 4, usar la plantilla de 4 gráficos
        template_file = "plantilla_contenido4.pptx"

    prs = Presentation(f"{DATA_DIR}/plantillas/{template_file}")

    # Asumimos que la plantilla tiene una sola diapositiva
    slide = prs.slides[0]

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

    feet_l, feet_r, periodo = (
        data.get("pie_l", ""),
        data.get("pie_r", ""),
        data.get("periodo", ""),
    )

    # Diccionario de reemplazos
    replacements = {
        "{{ph_titulo}}": slide_content.get("titulo", ""),
        "{{ph_periodo}}": periodo,
        "{{ph_title_1}}": slide_content.get("title_1", ""),
        "{{ph_kpis_1}}": normalize_kpi_text(slide_content.get("kpis_1", ""), 250),
        "{{ph_title_2}}": slide_content.get("title_2", ""),
        "{{ph_kpis_2}}": normalize_kpi_text(slide_content.get("kpis_2", ""), 250),
        "{{ph_title_3}}": slide_content.get("title_3", ""),
        "{{ph_kpis_3}}": normalize_kpi_text(slide_content.get("kpis_3", ""), 250),
        "{{ph_title_4}}": slide_content.get("title_4", ""),
        "{{ph_kpis_4}}": normalize_kpi_text(slide_content.get("kpis_4", ""), 250),
        "{{ph_pie_l}}": feet_l,
        "{{ph_pie_r}}": feet_r,
        # Sugerencias (opcionales, vacío si no vienen)
        "{{ph_sugerencia_1}}": normalize_suggestion_text(
            slide_content.get("sugerencia_1", ""), 200
        ),
        "{{ph_sugerencia_2}}": normalize_suggestion_text(
            slide_content.get("sugerencia_2", ""), 200
        ),
        "{{ph_sugerencia_3}}": normalize_suggestion_text(
            slide_content.get("sugerencia_3", ""), 200
        ),
        "{{ph_sugerencia_4}}": normalize_suggestion_text(
            slide_content.get("sugerencia_4", ""), 200
        ),
        "{{ph_sugerencia4}}": normalize_suggestion_text(
            slide_content.get("sugerencia_4", ""), 200
        ),
    }
    charts = slide_content.get("charts", {})
    # if product_type != "wazuh" and charts:
    #     replacements["{{ph_utilizacion}}"] = slide_content.get("kpi_title", "")
    #     replacements["{{ph_kpis}}"] = slide_content.get("kpis", "")
    #     replacements["{{ph_soporte}}"] = slide_content.get("soporte_title", "")
    #     replacements["{{ph_soporte_kpis}}"] = slide_content.get("soporte_kpi", "")

    # Reemplazar texto marcador
    modified = replace_placeholders(slide, replacements)

    tf_flags = {}
    for tf, key in modified:
        key_l = key.lower()
        flags = tf_flags.setdefault(
            id(tf), {"tf": tf, "titulo": False, "kpis": False, "sugerencia": False}
        )
        if "titulo" in key_l:
            flags["titulo"] = True
        if "kpis" in key_l:
            flags["kpis"] = True
        if "sugerencia" in key_l:
            flags["sugerencia"] = True

    for flags in tf_flags.values():
        tf = flags["tf"]
        if flags["titulo"]:
            apply_text_formatting(tf, font_name="Calibri", size=18)
            continue

        if flags["kpis"] or flags["sugerencia"]:
            apply_text_formatting(tf, font_name="Calibri", size=12, set_line=False)
            for p in tf.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                p.space_before = Pt(0)
                p.space_after = Pt(0)
                p.line_spacing = 1.15

                # Destacar el título SUGERENCIAS en cursiva.
                if p.text.strip().upper() == "SUGERENCIAS:":
                    for run in p.runs:
                        run.font.italic = True
            continue

        apply_text_formatting(tf, font_name="Aptos", size=12)
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

    output_path = f"{DATA_DIR}/pptx-parts/contenido_{product_type}.pptx"
    prs.save(output_path)
    return output_path


def generar_slide_producto(
    resumen, product_type, data, logo_stream, pie_l="", pie_r=""
):
    prs = Presentation(f"{DATA_DIR}/plantillas/plantilla_{product_type}.pptx")
    slide = prs.slides[0]

    # Definir campos por tipo de producto
    campos_por_tipo = {
        "uas": ["usu_per", "usu_esp", "solicitudes", "revalida"],
        "beyondtrust": ["pra", "rs", "ps", "adb", "epm"],
        "whalemate": ["sim", "aca", "ana", "grh", "cad"],
    }

    # Mapear product_type a grupo
    tipo_grupo = ""
    if product_type.lower() == "uas":
        tipo_grupo = "uas"
    elif product_type.lower() == "beyondtrust":
        tipo_grupo = "beyondtrust"
    elif product_type.lower() == "whalemate":
        tipo_grupo = "whalemate"

    replacements = {
        "{{ph_resumen}}": resumen,
        "{{ph_pie_l}}": data.get("pie_l", ""),
        "{{ph_pie_r}}": data.get("pie_r", ""),
    }

    # Agregar los campos nuevos si corresponden
    if tipo_grupo:
        for campo in campos_por_tipo[tipo_grupo]:
            replacements[f"{{{{ph_{campo}}}}}"] = data.get(campo, " ")

    modified = replace_placeholders(slide, replacements)

    for tf, key in modified:
        key_l = key.lower()
        if "titulo" in key_l:
            apply_text_formatting(tf, font_name="Calibri", size=18)
        elif "sub" in key_l:
            apply_text_formatting(tf, font_name="Calibri", size=14)
        elif "kpis" in key_l or "sugerencia" in key_l:
            # Mantener KPIs y sugerencias compactos para compartir el alto disponible.
            apply_text_formatting(tf, font_name="Calibri", size=12, set_line=False)
            for p in tf.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                p.space_before = Pt(0)
                p.space_after = Pt(0)
        elif "pie" in key_l:
            apply_text_formatting(tf, font_name=None, size=10)
        else:
            apply_text_formatting(tf, font_name="Aptos", size=12)

    _insert_logo_with_scaling(slide, logo_stream)

    output = f"{DATA_DIR}/pptx-parts/producto_{product_type}.pptx"
    prs.save(output)
    return output


def generar_cierre(data, logo_stream):
    cierre = data["despedida"]
    prs = Presentation(f"{DATA_DIR}/plantillas/plantilla_cierre.pptx")
    slide = prs.slides[0]

    replacements = {
        "{{ph_titulo}}": cierre.get("titulo", ""),
        "{{ph_pie_l}}": data.get("pie_l", ""),
        "{{ph_pie_r}}": data.get("pie_r", ""),
    }

    modified = replace_placeholders(slide, replacements)

    for tf, key in modified:
        key_l = key.lower()
        if "titulo" in key_l:
            apply_text_formatting(tf, font_name="Calibri", size=18)
        elif "pie" in key_l:
            apply_text_formatting(tf, font_name=None, size=10)
        else:
            apply_text_formatting(tf, font_name=None, size=12)

    _insert_logo_with_scaling(slide, logo_stream)

    output = f"{DATA_DIR}/pptx-parts/cierre.pptx"
    prs.save(output)
    return output


def generar_buenas_practicas(data, logo_stream):
    prs = Presentation(f"{DATA_DIR}/plantillas/plantilla_buenas_practicas.pptx")
    slide = prs.slides[0]

    replacements = {
        "{{ph_pie_l}}": data.get("pie_l", ""),
        "{{ph_pie_r}}": data.get("pie_r", ""),
    }

    modified = replace_placeholders(slide, replacements)

    for tf, key in modified:
        key_l = key.lower()
        if "titulo" in key_l:
            apply_text_formatting(tf, font_name="Calibri", size=18)
        elif "pie" in key_l:
            apply_text_formatting(tf, font_name=None, size=10)
        else:
            apply_text_formatting(tf, font_name=None, size=12)

    # Asegurar que el título "BUENAS PRACTICAS" sea Calibri 18
    for shape in slide.shapes:
        if shape.has_text_frame and "BUENAS PRACTICAS" in shape.text.upper():
            apply_text_formatting(shape.text_frame, font_name="Calibri", size=18)

    _insert_logo_with_scaling(slide, logo_stream)

    output = f"{DATA_DIR}/pptx-parts/buenas_practicas.pptx"
    prs.save(output)
    return output


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
