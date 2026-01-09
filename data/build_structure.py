#!/usr/bin/env python3
import sys
import json
import base64
from datetime import datetime, timedelta
import locale

DATA_DIR = "/data"
slide_info = ["resumen", "sugerencia", "sugerencia_version"]


MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def formatea_mes_anio_es(dt: datetime) -> str:
    """Devuelve 'Mes Año' en español (p. ej., 'Diciembre 2025')."""
    return f"{MESES_ES.get(dt.month, 'Mes')} {dt.year}"


def chart(values, name, build, kpis):
    total = 0
    for v in values:
        count = sum(values[v])
        total += count
        if count > 0:
            kpis[v] = count

    if total > 0:
        build["charts"][name] = {
            "type": "bar",
            "labels": ["Semana 1", "Semana2", "Semana 3", "Semana 4"],
            **values,
        }


def build_slide(product, product_name, chart_definitions, pointer_resume):
    build = {
        "titulo": product_name.upper().split(".")[0],
        "kpis": "",
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
            build[semana_key] = valor if valor != "null" else ""
            if semana_key == "sugerencia_version":
                break
        else:
            for chart_name, series_def in chart_definitions.items():
                for serie_name, json_key in series_def.items():
                    val = semana.get(json_key, 0)
                    try:
                        val = int(float(val))
                    except (ValueError, TypeError):
                        val = 0
                    chart_data[chart_name][serie_name].append(val)

    # 4. Generamos los gráficos y los KPIs a partir de los datos recolectados.
    for chart_name, data in chart_data.items():
        chart(data, chart_name, build, kpis)

    # 5. Construimos la cadena de KPIs.
    for serie, total in kpis.items():
        nombre_amigable = all_series.get(serie, serie)
        build["kpis"] += f"{nombre_amigable}: {total}\n"

    return build


# --------------------------------------------------------------
# MAIN
# --------------------------------------------------------------
def main():
    raw = sys.argv[1]
    data = json.loads(base64.b64decode(raw))
    main = data["main"]
    products = data["products"]

    # Convertir la fecha_portada
     
    fecha_iso = main.get("fecha_portada")

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
                        dt_object = datetime.fromisoformat(fecha_iso.replace("Z", "+00:00"))
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
            else:
                main["fecha_portada"] = "Fecha no válida"

        except Exception:
            main["fecha_portada"] = "Fecha no válida"

    else:
        main["fecha_portada"] = "Fecha no válida"

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
        with open(f"{DATA_DIR}/charts/chart_{product}.json", "r", encoding="utf-8") as chart_file:
            chart = json.load(chart_file)

        slide_data = build_slide(
            parse_products[product], product, chart, pointer_resumen
        )
        file_slide = {
            "uas": "plantilla_contenido.pptx",
            "wazuh": "plantilla_contenido_no_kpis.pptx",
            "ardid": "plantilla_contenido.pptx",
            "invgate.asj": "plantilla_contenido_no_kpis.pptx",
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

    print(json.dumps({"status": "ok", "output_file": data["main"]}))


if __name__ == "__main__":
    main()
