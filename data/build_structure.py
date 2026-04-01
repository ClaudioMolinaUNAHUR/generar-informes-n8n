#!/usr/bin/env python3
import sys
import json
import base64
from datetime import datetime, timedelta
import locale

DATA_DIR = "/data"
slide_info = ["resumen", "sugerencia_1", "sugerencia_2", "sugerencia_3 "]


MESES_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def formatea_mes_anio_es(dt: datetime) -> str:
    """Devuelve 'Mes Año' en español (p. ej., 'Diciembre 2025')."""
    return f"{MESES_ES.get(dt.month, 'Mes')} {dt.year}"


def ultimo_dia_mes(dt: datetime) -> datetime:
    """Devuelve el último día del mes de la fecha dada."""
    siguiente_mes = (dt.replace(day=28) + timedelta(days=4)).replace(day=1)
    return siguiente_mes - timedelta(days=1)


def chart(values, name, build, kpis):
    total = 0
    for v in values:
        count = sum(values[v])
        total += count
        if count > 0:
            kpis[v] = count
    chart_name = {"name": name, "used": False}
    if total > 0:
        chart_name["used"] = True

        build["charts"][name] = {
            "type": "bar",
            "labels": ["Semana 1", "Semana 2", "Semana 3", "Semana 4"],
            **values,
        }
    return chart_name


def build_slide(product, product_name, chart_definitions, pointer_resume, resume):
    product_name_clean = product_name.split(".")[0]
    build = {
        "titulo": product_name_clean.upper(),
        "resumen": resume + "\n",
        "periodo": "",
        "sugerencia_1": "",
        "sugerencia_2": "",
        "sugerencia_3": "",
        "charts": {},
    }
    # 2. Inicializamos dinámicamente los contenedores para los datos de los gráficos.
    chart_data = {
        chart_name: {serie: [] for serie in series}
        for chart_name, series in chart_definitions.items()
    }
    # También creamos un mapa plano de todas las series para facilitar la búsqueda.
    all_series = {
        serie: json_key
        for series in chart_definitions.values()
        for serie, json_key in series.items()
    }
    kpis = {}

    # 3. Procesamos los datos en un único bucle optimizado.
    for semana in product:
        semana_key = semana.get("Semana", "").strip()
        if semana_key in slide_info:
            valor = semana.get(pointer_resume, "")
            build[semana_key] += valor if valor != "null" else ""
        else:
            if semana_key == "sugerencia_3":
                break
            for chart_name, series_def in chart_definitions.items():
                for serie_name, json_key in series_def.items():
                    val = semana.get(json_key, 0)
                    try:
                        val = int(float(val))
                    except (ValueError, TypeError):
                        val = 0
                    chart_data[chart_name][serie_name].append(val)

    # 4. Generamos los gráficos y los KPIs a partir de los datos recolectados.
    chart_names_used = []
    for chart_name, data in chart_data.items():
        name = chart(data, chart_name, build, kpis)
        if name:
            chart_names_used.append(name)

    count = sum(1 for c in chart_names_used if c["used"]) 

    if count == 1:
       for chart_name in chart_names_used:
            if chart_name["used"] and chart_name["name"] == "soporte":
                build["periodo"] = ""
                build["charts"] = {}
                return build
    
    position = 1
    for i, chart_name in enumerate(chart_names_used):
        if chart_name["used"]:      
            if f"kpis_{position}" not in build or f"title_{position}" not in build:
                build[f"kpis_{position}"] = ""
                build[f"title_{position}"] = ""
            for serie_name in chart_definitions[chart_name["name"]].keys():
                if serie_name in kpis:
                    nombre_amigable = all_series.get(serie_name, serie_name)
                    build[f"title_{position}"] = chart_name["name"].capitalize().replace("_", " ")
                    build[f"kpis_{position}"] += f"{nombre_amigable}: {kpis[serie_name]}\n"
            if chart_name["name"] != "soporte":
                build[f"kpis_{position}"] += f"SUGERENCIAS: {build[f'sugerencia_{i+1}']}"
            position += 1
    return build


# --------------------------------------------------------------
# MAIN
# --------------------------------------------------------------
def main():
    raw = sys.argv[1]
    data = json.loads(base64.b64decode(raw))
    # with open("estructure.json", "r", encoding="utf-8") as f:
    #     data = json.load(f)

    main = data["main"]
    products = data["products"]

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
            mes_anterior_ultimo = (mes_actual.replace(day=1) - timedelta(days=1))
            mes_actual_ultimo = ultimo_dia_mes(mes_actual)
            main["periodo"] = (
                f"Periodo: {mes_anterior_ultimo.day:02d}/{mes_anterior_ultimo.month:02d} - {mes_actual_ultimo.day:02d}/{mes_actual_ultimo.month:02d}"
            )
        else:
            main["periodo"] = f"Periodo: {main.get('fecha_portada', '')}" if main.get('fecha_portada') else ""

    parse_products = {}
    actual_product = ""
    # separo productos { "uas": product_data[] }
    for product in products:
        actual_product = product.get("product", actual_product)
        if actual_product not in parse_products:
            parse_products[actual_product] = []
        parse_products[actual_product].append(product)

    # agrego contenidos slide
    for product in parse_products:
        pointer_resumen = list(parse_products[product][0].keys())[1]
        with open(
            f"{DATA_DIR}/charts/chart_{product}.json", "r", encoding="utf-8"
        ) as chart_file:
            chart = json.load(chart_file)
        resume = main.get(f"resume_{product.split('.')[0].lower()}", "")
        slide_data = build_slide(
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
        }
        if slide_data:
            main["slides"].append(
                {
                    "type": product,
                    "slide": slide_data,
                    "file_slide": file_slide[product],
                }
            )

    # with open("salida.json", "w", encoding="utf-8") as f:
    #     json.dump({"data": main}, f, indent=2, ensure_ascii=False)

    print(json.dumps({"status": "ok", "output_file": data["main"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
