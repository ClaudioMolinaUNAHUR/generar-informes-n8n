from datetime import datetime, timedelta

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

def formatea_mes_anio_es(dt: datetime) -> str:
    return f"{MESES_ES.get(dt.month, 'Mes')} {dt.year}"

def ultimo_dia_mes(dt: datetime) -> datetime:
    siguiente_mes = (dt.replace(day=28) + timedelta(days=4)).replace(day=1)
    return siguiente_mes - timedelta(days=1)

def _chart_internal(values, name, build, kpis):
    total = 0
    for v in values:
        count = sum(values[v])
        total += count
        if count > 0:
            kpis[v] = count
    
    chart_res = {"name": name, "used": total > 0}
    if total > 0:
        build["charts"][name] = {
            "type": "bar",
            "labels": ["Semana 1", "Semana 2", "Semana 3", "Semana 4"],
            **values,
        }
    return chart_res

def build_slide_structure(product_data, product_name, chart_definitions, pointer_resume, resume):
    product_name_clean = product_name.split(".")[0]
    build = {
        "titulo": product_name_clean.upper(),
        "resumen": resume + "\n",
        "periodo": "",
        "sugerencia_1": "", "sugerencia_2": "", "sugerencia_3": "",
        "charts": {},
    }
    
    chart_data = {
        chart_name: {serie: [] for serie in series}
        for chart_name, series in chart_definitions.items()
    }
    
    all_series = {
        serie: json_key
        for series in chart_definitions.values()
        for serie, json_key in series.items()
    }
    
    kpis = {}
    slide_info_fields = ["resumen", "sugerencia_1", "sugerencia_2", "sugerencia_3"]

    for semana in product_data:
        semana_key = semana.get("Semana", "").strip()
        if semana_key in slide_info_fields:
            valor = semana.get(pointer_resume, "")
            build[semana_key] += str(valor) if valor != "null" else ""
        else:
            if semana_key == "sugerencia_3":
                break
            for chart_name, series_def in chart_definitions.items():
                for serie_name, json_key in series_def.items():
                    try:
                        val = int(float(semana.get(json_key, 0)))
                    except:
                        val = 0
                    chart_data[chart_name][serie_name].append(val)

    chart_names_used = []
    for chart_name, data in chart_data.items():
        chart_names_used.append(_chart_internal(data, chart_name, build, kpis))

    # Lógica de limpieza para soporte solo
    count_used = sum(1 for c in chart_names_used if c["used"])
    if count_used == 1:
       for c_res in chart_names_used:
            if c_res["used"] and c_res["name"] == "soporte":
                build["periodo"] = ""
                build["charts"] = {}
                return build
    
    position = 1
    for c_res in chart_names_used:
        if c_res["used"]:      
            kpi_key = f"kpis_{position}"
            title_key = f"title_{position}"
            build[kpi_key] = ""
            build[title_key] = c_res["name"].capitalize().replace("_", " ")
            
            for serie_name in chart_definitions[c_res["name"]].keys():
                if serie_name in kpis:
                    nombre_amigable = all_series.get(serie_name, serie_name)
                    build[kpi_key] += f"{nombre_amigable}: {kpis[serie_name]}\n"
            
            if c_res["name"] != "soporte":
                sug_val = build.get(f"sugerencia_{position}", "")
                if sug_val:
                    build[kpi_key] += f"SUGERENCIAS: {sug_val}"
            position += 1
            
    return build