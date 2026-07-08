from datetime import datetime, timedelta
import re
import unicodedata

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

def formatea_mes_anio_es(dt: datetime) -> str:
    return f"{MESES_ES.get(dt.month, 'Mes')} {dt.year}"

def ultimo_dia_mes(dt: datetime) -> datetime:
    siguiente_mes = (dt.replace(day=28) + timedelta(days=4)).replace(day=1)
    return siguiente_mes - timedelta(days=1)

MAX_CHART_WEEKS = 4

LOGO_SAR = "sar.png"


def procesa_fecha_portada(fecha_str: str, logo: str) -> str:
    """
    Recibe fecha_portada del JSON de entrada (formato "YYYY-MM") y el logo,
    parsea la fecha y delega en fecha_portada para generar el periodo formateado.

        Ejemplo:
            procesa_fecha_portada("2026-05", "sar.png")  →  "2026\\n26/04 - 26/05"
            procesa_fecha_portada("2026-05", "otro.png") →  "2026\\n01/05 - 31/05"
    """
    try:
        fecha = datetime.strptime(fecha_str.strip(), "%Y-%m")
    except (ValueError, AttributeError):
        return fecha_str  # Si el formato no es el esperado, devolver tal cual
    return fecha_portada(fecha, logo)


def fecha_portada(fecha: datetime, logo: str) -> str:
    """
    Devuelve el periodo de la portada según el logo.

    - Caso general:  año\\nDD/MM - DD/MM  (primer y último día del mes)
    - Caso sar.png:  año\\nDD/MM - DD/MM  (día 26 del mes anterior al día 26 del mes)

    Ejemplos:
      fecha=junio-2026, logo=cualquiera  →  "2026\\n01/06 - 30/06"
      fecha=junio-2026, logo=sar.png     →  "2026\\n26/05 - 26/06"
    """
    anio = fecha.year

    if logo and logo.strip().lower() == LOGO_SAR:
        # Periodo: día 26 del mes anterior → día 26 del mes actual
        primer_dia = (fecha.replace(day=1) - timedelta(days=1)).replace(day=26)
        ultimo_dia = fecha.replace(day=26)
    else:
        # Periodo: primer y último día del mes
        primer_dia = fecha.replace(day=1)
        ultimo_dia = ultimo_dia_mes(fecha)

    return f"{anio}\n{primer_dia.strftime('%d/%m')} - {ultimo_dia.strftime('%d/%m')}"


def _chart_internal(values, name, build, kpis, series_defs, show_zeros=False):
    total = 0
    for v in values:
        series_values = values[v]
        if v == "bt_credenciales":
            count = series_values[-1] if series_values else 0
        else:
            count = sum(series_values)
        total += count
        if count > 0:
            kpis[v] = count
        elif show_zeros:
            # Registrar explícitamente el 0 para que aparezca en los KPIs
            kpis[v] = 0

    chart_res = {"name": name, "used": total > 0, "show_zeros": show_zeros, "show_kpi": True}
    if total > 0:
        labels = [f"Semana {i+1}" for i in range(MAX_CHART_WEEKS)]
        visible_values = {
            serie_name: serie_values
            for serie_name, serie_values in values.items()
            if _show_series_in_chart(series_defs.get(serie_name))
        }
        build["charts"][name] = {
            "type": "bar",
            "labels": labels,
            **visible_values,
        }
    return chart_res

# Nombres de charts que deben mostrar KPIs aunque su valor sea 0
CHARTS_SHOW_ZEROS = {"takedown_estado_resolucion"}


def _normalize_field_name(value):
    text = str(value or "").strip().lower()
    text = "".join(
        c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c)
    )
    return re.sub(r"\s+", " ", text)


def _normalize_field_name_loose(value):
    base = _normalize_field_name(value)
    tokens = []
    for tok in base.split(" "):
        # Tolerar singular/plural en encabezados del origen (usuario/usuarios, etc.)
        if len(tok) > 3 and tok.endswith("s"):
            tok = tok[:-1]
        tokens.append(tok)
    return " ".join(tokens)


def _field_candidates(field_def):
    if isinstance(field_def, dict):
        field_def = field_def.get("field") or field_def.get("fields") or field_def.get("label")
    if isinstance(field_def, (list, tuple)):
        return [str(v) for v in field_def if v is not None]
    return [str(field_def)]


def _series_label(field_def, fallback):
    if isinstance(field_def, dict):
        return str(field_def.get("label") or field_def.get("title") or fallback)
    return fallback


def _show_series_in_chart(field_def):
    if isinstance(field_def, dict):
        return field_def.get("chart", True)
    return True


def _get_week_value(week_row, field_def, default=0):
    candidates = _field_candidates(field_def)

    # 1) Intento exacto primero.
    for candidate in candidates:
        if candidate in week_row:
            return week_row.get(candidate, default)

    # 2) Fallback normalizado: ignora tildes, mayúsculas y espacios duplicados.
    normalized_week = {
        _normalize_field_name(k): v for k, v in week_row.items() if isinstance(k, str)
    }
    for candidate in candidates:
        val = normalized_week.get(_normalize_field_name(candidate))
        if val is not None:
            return val

    # 3) Fallback flexible: tolera plural/singular simple por token.
    normalized_week_loose = {
        _normalize_field_name_loose(k): v for k, v in week_row.items() if isinstance(k, str)
    }
    for candidate in candidates:
        val = normalized_week_loose.get(_normalize_field_name_loose(candidate))
        if val is not None:
            return val

    return default


def _primary_field_name(field_def):
    candidates = _field_candidates(field_def)
    return candidates[0] if candidates else ""


def build_slide_structure(product_data, product_name, chart_definitions, pointer_resume, resume, fecha: datetime = None, logo: str = "", main_data: dict = None):
    product_name_clean = product_name.split(".")[0]
    normalized_product_name = (product_name or "").lower().split(".")[0]
    
    # Base de la estructura
    build = {
        "titulo": product_name_clean.upper(),
        "resumen": str(resume or "").strip(),
        "periodo": fecha_portada(fecha, logo) if fecha else "",
        "sugerencia_1": "", "sugerencia_2": "", "sugerencia_3": "", "sugerencia_4": "",
        "sugerencia_version": "",
        "work_t": "",
        "work_n": "",
        "charts": {},
        "desc": str(resume or "").strip(),
    }

    # ── MODIFICACIÓN CLAVE: Mapeo de datos globales ──
    # Si viene el diccionario principal, arrastramos los campos faltantes a la slide
    if main_data and isinstance(main_data, dict):
        campos_globales = [
            "usu_per", "usu_esp", "solicitudes", "revalida", 
            "pra", "rs", "pws", "adb", "epm", 
            "sim", "aca", "ana", "grh", "cad", 
            "ddv", "snc", "enc", "ecn", "iav", 
            "isd", "iam", "fim",
            "pdm", "th", "cti", "fdd", "ddw", "tkd",
            "tra", "pgs", "lgn", "rfg", "alr", "bwl", "rdc", "rdt", "rdp", "rdl",
            "sca", "pdc", "shv", "cyd", "icc",
            "work_t", "work_n", "sugerencia_version"
        ]
        for campo in campos_globales:
            if campo in main_data:
                build[campo] = main_data[campo]

    print("[DEBUG] build_slide_structure: product_name=", product_name)
    print("[DEBUG] build_slide_structure: pointer_resume=", pointer_resume)
    print("[DEBUG] build_slide_structure: chart_definitions keys=", list(chart_definitions.keys()))
    print("[DEBUG] build_slide_structure: product_data ejemplo=", product_data[0] if product_data else None)

    def _is_metadata_key(name):
        return isinstance(name, str) and name.startswith("_")

    chart_data = {}
    row_labels = {}
    row_chart_meta = {}
    special_chart_labels = {}
    for chart_name, series in chart_definitions.items():
        chart_data[chart_name] = {
            **{serie: [] for serie in series if not _is_metadata_key(serie)}
        }
        if isinstance(series, dict):
            row_regex = series.get("_row_regex")
            if row_regex:
                try:
                    row_chart_meta[chart_name] = {
                        "matcher": re.compile(row_regex, flags=re.IGNORECASE),
                        "label_prefix": series.get("_label_prefix", ""),
                        "show_kpi": series.get("_show_kpi", True),
                    }
                    row_labels[chart_name] = []
                except re.error:
                    row_chart_meta[chart_name] = None

    all_series = {
        serie: _series_label(json_key, _primary_field_name(json_key))
        for series in chart_definitions.values()
        for serie, json_key in series.items()
        if not _is_metadata_key(serie)
    }

    kpis = {}
    slide_info_fields = [
        "resumen",
        "sugerencia_1",
        "sugerencias_1",
        "sugerencia_2",
        "sugerencias_2",
        "sugerencia_3",
        "sugerencias_3",
        "sugerencia_4",
        "sugerencias_4",
        "sugerencia_version",
        "sugerencias_version",
        "desc",
        "work_t",
        "work_n",
    ]

    for semana in product_data:
        print("[DEBUG] Semana procesada:", semana)
        semana_key = semana.get("Semana", "").strip()
        has_rating_value = bool(str(semana.get("Rating") or "").strip())

        # --- CAPTURA DE TEXTOS CON FALLBACK DINÁMICO ---
        if semana_key in slide_info_fields:
            valor = semana.get(pointer_resume, "")
            # Si el puntero por defecto no trae contenido válido, se busca dinámicamente en otra columna con texto
            if not str(valor).strip() or valor == "null":
                for k, v in semana.items():
                    if k != "Semana" and isinstance(v, str) and v.strip() and v != "null":
                        valor = v
                        break

            if str(valor).strip() == "null":
                continue

            target_key = semana_key.replace("sugerencias_", "sugerencia_")
            if target_key == "resumen" and build[target_key]:
                build[target_key] = f"{build[target_key]}\n{valor}" if str(valor).strip() else build[target_key]
            else:
                build[target_key] = f"{build[target_key]}{valor}" if build[target_key] else str(valor)

        # --- PROCESAMIENTO DE FILAS DE GRÁFICOS (SEMANAS / RATINGS) ---
        elif semana_key.startswith("Semana") or (normalized_product_name == "sonarqube" and ("rating" in semana_key.lower() or has_rating_value)):
            for chart_name, series_def in chart_definitions.items():
                if chart_name in row_chart_meta:
                    continue

                # Flujo especial: Gráfico de Vulnerabilidades de SonarQube (Eje X basado en Letras A-E)
                if normalized_product_name == "sonarqube" and chart_name == "vulnerabilidades":
                    rating_label = str(semana.get("Rating") or "").strip()
                    if not rating_label and "Rating" in semana_key:
                        rating_label = semana_key.replace("Rating", "").strip()
                    
                    if rating_label:
                        special_chart_labels.setdefault(chart_name, [])
                        if rating_label not in special_chart_labels[chart_name]:
                            special_chart_labels[chart_name].append(rating_label)
                    
                    for serie_name, json_key in series_def.items():
                        if _is_metadata_key(serie_name):
                            continue
                        try:
                            val = int(float(_get_week_value(semana, json_key, 0)))
                        except Exception:
                            val = 0
                        chart_data[chart_name][serie_name].append(val)
                    continue

                # Flujo estándar: limitado por la cantidad de semanas máximas
                for serie_name, json_key in series_def.items():
                    if _is_metadata_key(serie_name):
                        continue
                    if len(chart_data[chart_name][serie_name]) >= MAX_CHART_WEEKS:
                        continue
                    try:
                        val = int(float(_get_week_value(semana, json_key, 0)))
                    except Exception:
                        val = 0
                    chart_data[chart_name][serie_name].append(val)
        else:
            for chart_name, meta in row_chart_meta.items():
                if not meta or not meta["matcher"].search(semana_key):
                    continue
                label = semana_key
                prefix = meta.get("label_prefix", "")
                if prefix and label.lower().startswith(prefix.lower()):
                    label = label[len(prefix) :].strip()
                row_labels[chart_name].append(label)
                for serie_name, json_key in chart_definitions[chart_name].items():
                    if _is_metadata_key(serie_name):
                        continue
                    try:
                        val = int(float(_get_week_value(semana, json_key, 0)))
                    except Exception:
                        val = 0
                    chart_data[chart_name][serie_name].append(val)
                break

    chart_names_used = []
    print("[DEBUG] chart_data recolectado:", chart_data)
    for chart_name, data in chart_data.items():
        if chart_name in row_chart_meta:
            labels = row_labels.get(chart_name, [])
            total = sum(sum(vals) for vals in data.values())
            if labels and total > 0:
                build["charts"][chart_name] = {
                    "type": "bar",
                    "labels": labels,
                    **{serie: values for serie, values in data.items() if not _is_metadata_key(serie)},
                }
                chart_names_used.append({
                    "name": chart_name,
                    "used": True,
                    "show_zeros": False,
                    "show_kpi": row_chart_meta[chart_name].get("show_kpi", True),
                })
            else:
                chart_names_used.append({
                    "name": chart_name,
                    "used": False,
                    "show_zeros": False,
                    "show_kpi": row_chart_meta[chart_name].get("show_kpi", True),
                })
            continue

        # Inyección de Gráficos con etiquetas especiales (SonarQube)
        labels = special_chart_labels.get(chart_name)
        if labels:
            build["charts"][chart_name] = {
                "type": "bar",
                "labels": labels,
                **{serie: values for serie, values in data.items() if not _is_metadata_key(serie)},
            }
            total = sum(sum(vals) for vals in data.values())
            chart_names_used.append({
                "name": chart_name,
                "used": total > 0,
                "show_zeros": chart_name in CHARTS_SHOW_ZEROS,
                "show_kpi": True
            })
            
            # Acumular los KPIs del gráfico especial en el diccionario global
            for v, series_values in data.items():
                count = sum(series_values)
                if count > 0:
                    kpis[v] = count
            continue

        show_zeros = chart_name in CHARTS_SHOW_ZEROS
        chart_names_used.append(
            _chart_internal(
                data,
                chart_name,
                build,
                kpis,
                chart_definitions.get(chart_name, {}),
                show_zeros=show_zeros,
            )
        )
    print("[DEBUG] chart_names_used:", chart_names_used)

    # Lógica de limpieza para soporte solo
    count_used = sum(1 for c in chart_names_used if c["used"])
    if count_used == 1:
       for c_res in chart_names_used:
            if c_res["used"] and c_res["name"] == "soporte":
                build["charts"] = {}
                return build

    position = 1
    for c_res in chart_names_used:
        if c_res["used"]:
            kpi_key = f"kpis_{position}"
            title_key = f"title_{position}"
            build[kpi_key] = ""
            build[title_key] = c_res["name"].capitalize().replace("_", " ")
            try:
                if c_res.get("show_kpi", True):
                    print(f"[DEBUG] Procesando KPIs para chart: {c_res['name']}")
                    chart_def = chart_definitions[c_res["name"]]
                    for serie_name in chart_def.keys():
                        # Mostrar el KPI si tiene valor > 0, o si el chart permite mostrar 0s
                        if serie_name in kpis and (kpis[serie_name] > 0 or c_res.get("show_zeros")):
                            nombre_amigable = all_series.get(serie_name, serie_name)
                            build[kpi_key] += f"{nombre_amigable}: {kpis[serie_name]}\n"
            except Exception as e:
                print(f"[ERROR] Al procesar KPIs/titles para chart '{c_res['name']}': {e}")
            position += 1

    return build