import os
import json
import base64
from datetime import datetime, timedelta
from fastapi import FastAPI, HTTPException, Request, Query, UploadFile, File
from fastapi.responses import FileResponse
from models.schemas import GenerateRequest
from utils.helpers import (
    DATA_DIR,
    get_safe_path,
    get_logo_from_base64,
    create_composite_logo_from_base64_list,
    log,
)
from services.pdf_service import (
    generar_portada,
    generar_slide_producto,
    generar_contenido_slide,
    generar_cierre,
    generar_buenas_practicas,
    convert_to_pdf,
    unir_pdfs,
)
from services.structure_service import (
    build_slide_structure,
    ultimo_dia_mes,
    formatea_mes_anio_es,
)

app = FastAPI()

def robust_parse_data(val):
    """Intenta convertir strings (JSON) a objetos y desvuelve envoltorios 'data' de n8n."""
    if isinstance(val, str):
        try:
            val = json.loads(val)
        except Exception:
            pass
    if isinstance(val, dict) and "data" in val and len(val) <= 2:
        # Si es un dict y tiene 'data', es probable que sea el envoltorio de n8n
        return robust_parse_data(val["data"])
    return val

@app.post("/generate")
async def generate_report(request: Request):
    try:
        raw_body = await request.json()
        data = robust_parse_data(raw_body)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid JSON")

    if not isinstance(data, dict):
        raise HTTPException(status_code=400, detail="Data must be an object")

    split = data.get("split", 0)
    logo_stream = get_logo_from_base64(data.get("logo_base64"))

    logo_val = data.get("logo")
    empresa = logo_val[:-4].lower() if isinstance(logo_val, str) and len(logo_val) > 4 else "informe"

    slides_data = data.get("slides", [])

    generated_pptx = []
    if data.get("save", False):
        portada = generar_portada(data, logo_stream)
        generated_pptx.append(portada)

    for slide_item in slides_data:
        product_type = slide_item["type"]
        resumen = slide_item.get("slide", {}).get("resumen", "")
        producto_slide = generar_slide_producto(
            resumen, product_type, data, logo_stream
        )
        generated_pptx.append(producto_slide)
        content_pptx = generar_contenido_slide(slide_item, data, logo_stream)
        generated_pptx.append(content_pptx)

    if data.get("save", False):
        buenas_practicas = generar_buenas_practicas(data, logo_stream)
        generated_pptx.append(buenas_practicas)
        cierre = generar_cierre(data, logo_stream)
        generated_pptx.append(cierre)

    # Convertir a PDF y unir
    pdf_files_to_merge = [convert_to_pdf(f) for f in generated_pptx if f is not None]

    informe_name = []
    if split == 0:
        informe_name.append(unir_pdfs(pdf_files_to_merge, empresa))
    else:
        # Para split, asumir que se divide por producto, pero como ahora hay más slides, quizás ajustar.
        # Por simplicidad, mantener como estaba, pero con la nueva lista.
        # Pero el usuario no especificó para split, así que dejar como está, pero probablemente split no se usa.
        types = [slide.get("type", "") for slide in slides_data]
        # Para split, crear informes separados por producto, incluyendo portada, producto_slide, resume, content, buenas_practicas, cierre
        # Pero eso sería complejo, asumir split=0 por ahora.
        informe_name.append(unir_pdfs(pdf_files_to_merge, empresa))

    return {"file_names": informe_name}



@app.post("/build-structure")
async def build_structure(request: Request):
    try:
        body = await request.json()
        main = body.get("main", {})
        products = body.get("products", [])

        if not isinstance(main, dict):
            raise ValueError("'main' must be a dictionary")
        if not isinstance(products, list):
            raise ValueError("'products' must be a list")

        # Convertir la fecha_portada

        fecha_iso = main.get("fecha_portada")
        mes_actual = None
        if fecha_iso:
            try:
                dt_object = None

                # 🟢 Caso 1: String 'YYYY-MM'
                if isinstance(fecha_iso, str):
                    # Acepta 'YYYY-MM' y, por compatibilidad, ISO con día/hora
                    # Primero intentamos 'YYYY-MM'
                    try:
                        dt_object = datetime.strptime(fecha_iso, "%Y-%m")
                    except ValueError:
                        # Si viniera como 'YYYY-MM-DD...' (ISO), intentamos parseo estándar
                        # Reemplazamos la Z por +00:00 si existiera
                        try:
                            dt_object = datetime.fromisoformat(
                                fecha_iso.replace("Z", "+00:00")
                            )
                        except Exception:
                            dt_object = None

                    # Normalizamos al primer día del mes si venía 'YYYY-MM'
                    if dt_object:
                        dt_object = dt_object.replace(day=1)

                # 🟢 Caso 2: Número de Excel (por si aparece)
                elif isinstance(fecha_iso, (int, float)):
                    # Día base Excel (serial date): 1899-12-30
                    dt_object = datetime(1899, 12, 30) + timedelta(days=fecha_iso)
                    # Normalizamos al primer día del mes
                    dt_object = dt_object.replace(day=1)

                # Salida final
                if dt_object:
                    main["fecha_portada"] = formatea_mes_anio_es(dt_object)
                    mes_actual = dt_object
                else:
                    main["fecha_portada"] = "Fecha no válida"

            except Exception:
                main["fecha_portada"] = "Fecha no válida"

        else:
            main["fecha_portada"] = "Fecha no válida"

        if "periodo" not in main:
            if mes_actual:
                mes_anterior_ultimo = mes_actual.replace(day=1) - timedelta(days=1)
                mes_actual_ultimo = ultimo_dia_mes(mes_actual)
                main["periodo"] = (
                    f"Periodo: {mes_anterior_ultimo.day:02d}/{mes_anterior_ultimo.month:02d} - {mes_actual_ultimo.day:02d}/{mes_actual_ultimo.month:02d}"
                )
            else:
                main["periodo"] = (
                    f"Periodo: {main.get('fecha_portada', '')}"
                    if main.get("fecha_portada")
                    else ""
                )

        parse_products = {}
        actual_product = ""
        # separo productos { "uas": product_data[] }
        for product in products:
            actual_product = product.get("product", actual_product)
            if actual_product not in parse_products:
                parse_products[actual_product] = []
            parse_products[actual_product].append(product)

        # agrego contenidos slide
        if "slides" not in main or not isinstance(main["slides"], list):
            main["slides"] = []

        for product in parse_products:
            pointer_resumen = list(parse_products[product][0].keys())[1]
            if not parse_products[product] or not hasattr(parse_products[product][0], 'keys'):
                continue
            keys = list(parse_products[product][0].keys())
            if len(keys) < 2:
                continue
            pointer_resumen = keys[1]
            
            with open(
                f"{DATA_DIR}/charts/chart_{product}.json", "r", encoding="utf-8"
            ) as chart_file:
                chart = json.load(chart_file)
            resume = main.get(f"resume_{product.split('.')[0].lower()}", "")
            slide_data = build_slide_structure(
                parse_products[product], product, chart, pointer_resumen, resume
            )
            file_slide = {
                "uas": "plantilla_contenido.pptx",
                "wazuh": "plantilla_contenido.pptx",
                "akurtech": "plantilla_contenido.pptx",
                "invgate.asj": "plantilla_contenido.pptx",
                "invgate": "plantilla_contenido.pptx",
                "beyondtrust": "plantilla_contenido.pptx",
                "whalemate": "plantilla_contenido.pptx",
                "axur": "plantilla_contenido.pptx",
            }
            if slide_data:
                main["slides"].append(
                    {
                        "type": product,
                        "slide": slide_data,
                        "file_slide": file_slide[product],
                    }
                )
                        
        return {"status": "ok", "output_file": json.dumps(main, ensure_ascii=False)}
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error building structure: {str(e)}"
        )


@app.post("/generate-n-emp")
async def generate_pdf_n_emp(request: Request):
    try:
        raw_body = await request.json()
        data = robust_parse_data(raw_body)

        main_data = data.get("main", {})
        emp_codes = data.get("emp_codes", [])
        logos_base64_list = data.get("logos_base64", [])

        logo_stream = create_composite_logo_from_base64_list(logos_base64_list)

        empresa = ""
        length_emp_codes = len(emp_codes)
        for i, emp_code in enumerate(emp_codes):
            empresa += emp_code + "-" if i != length_emp_codes - 1 else emp_code
        empresa = empresa.lower()
        portada_pptx_file = generar_portada(main_data, logo_stream)
        cierre_pptx_file = generar_cierre(main_data, logo_stream)

        pdf_files_to_merge = []
        pdf_files_to_merge.append(convert_to_pdf(portada_pptx_file))

        full_informes_paths = [
            os.path.join(DATA_DIR, "generados", f"informe_{f.lower()}.pdf")
            for f in emp_codes
        ]
        pdf_files_to_merge.extend(full_informes_paths)

        pdf_files_to_merge.append(convert_to_pdf(cierre_pptx_file))

        final_pdf = unir_pdfs(pdf_files_to_merge, empresa)

        # with open(final_pdf, "rb") as f:
        #     b64 = base64.b64encode(f.read()).decode()
        return {"file_name": os.path.basename(f"informe_{empresa}")}
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Error generating PDF N Emp: {str(e)}"
        )


@app.get("/health")
def health():
    return {"status": "ok"}

@app.get("/files/read")
async def read_file_endpoint(path: str):
    """
    Permite descargar cualquier archivo dentro de la carpeta /data.
    Ejemplo: GET /files/read?path=charts/chart_uas.json
    """
    try:
        full_path = get_safe_path(path)
        
        # Log de depuración para entender el 404
        log(f"🔍 Accediendo a: {full_path} | Existe: {os.path.exists(full_path)} | Es archivo: {os.path.isfile(full_path)}")
        
        if not os.path.isfile(full_path):
            raise HTTPException(status_code=404, detail="Archivo no encontrado")
        return FileResponse(full_path)
    except ValueError as ve:
        raise HTTPException(status_code=403, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/files/list")
async def list_files_endpoint(path: str = "."):
    """
    Lista los archivos y carpetas dentro de una ruta en /data.
    Ejemplo: GET /files/list?path=config
    """
    try:
        full_path = get_safe_path(path)
        if not os.path.exists(full_path):
            raise HTTPException(status_code=404, detail="Ruta no encontrada")
        
        items = os.listdir(full_path)
        details = []
        for item in items:
            item_path = os.path.join(full_path, item)
            details.append({
                "name": item,
                "is_dir": os.path.isdir(item_path),
                "size": os.path.getsize(item_path) if os.path.isfile(item_path) else 0
            })
        return {"path": path, "items": details}
    except ValueError as ve:
        raise HTTPException(status_code=403, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/files/save")
async def save_file_endpoint(
    path: str = Query(...),
    file: UploadFile = File(...)
):
    try:
        # Leer archivo
        file_bytes = await file.read()

        # Armar path seguro
        full_path = get_safe_path(path)
        os.makedirs(os.path.dirname(full_path), exist_ok=True)

        # Guardar archivo
        with open(full_path, "wb") as f:
            f.write(file_bytes)

        return {
            "status": "ok",
            "filename": file.filename,
            "message": f"Archivo guardado en {path}"
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
