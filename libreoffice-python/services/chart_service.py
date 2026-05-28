import os
import re
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import textwrap
from utils.helpers import insert_image_scaled_by_width, log

def create_matplotlib_chart(chart_info, friendly_names, output_file):
    plt.figure(figsize=(10, 5))
    ctype = chart_info.get("type")
    chart_title = chart_info.get("title") or chart_info.get("titulo") or "<no title>"

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
    palette = ["#185FA5", "#0F6E56", "#534AB7", "#5F5E5A"]

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
            if len(vals) > len(labels):
                log(
                    f"DEBUG: create_matplotlib_chart - chart '{chart_title}' serie '{key}' tiene {len(vals)} valores pero labels {len(labels)}; truncando a labels"
                )
                vals = vals[: len(labels)]
            if len(vals) < len(labels):
                vals.extend([0] * (len(labels) - len(vals)))  # pad with 0

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
            vals = vals[:len(labels)]  # truncate if longer
            if len(vals) < len(labels):
                vals.extend([None] * (len(labels) - len(vals)))  # pad with None

            label_full = flat_friendly_names.get(key, key.replace("_", " ").capitalize())
            # Dividir etiquetas largas en varias líneas
            label = textwrap.fill(label_full, width=22)

            color = palette[idx % len(palette)]
            plt.plot(x, vals, label=label, marker="o", color=color)

        plt.xticks(x, labels, rotation=45, fontsize=16)
        plt.grid(axis="y", linestyle="-", color="#dcdcdc", linewidth=0.8)
        plt.legend(loc="best", fontsize=18)

    # Formato eje Y con separador de miles y margen superior
    try:
        import matplotlib.ticker as mtick
        ax = plt.gca()

        # Fuerza a que los ticks sean solo números enteros
        ax.yaxis.set_major_locator(mtick.MaxNLocator(integer=True))

        # Formateador con separador de miles
        ax.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x, pos: f"{int(x):,}"))

        # Agregar margen superior: 20% por encima del valor máximo
        current_ylim = ax.get_ylim()
        ax.set_ylim(current_ylim[0], current_ylim[1] * 1.20)
    except Exception:
        pass

    # Ajusta el layout para asegurar que la leyenda no se corte
    plt.tight_layout(rect=[0, 0.03, 0.95, 0.97])
    log(
        f"DEBUG: create_matplotlib_chart - saving chart '{chart_title}' type={ctype} labels={labels} series_keys={series_keys}"
    )
    plt.savefig(output_file, dpi=150, transparent=True)
    plt.close()


def _extract_chart_placeholder_index(shape):
    if not shape.has_text_frame:
        return None
    text = shape.text or ""
    match = re.search(r"ph_chart_(\d+)", text)
    if match:
        return int(match.group(1))
    return None


def add_charts(slide, charts, friendly_names, _unused_list=None):
    """
    Busca placeholders en la diapositiva identificándolos por el texto marcador "{{ph_chart"
    en lugar de nombres internos fijos. Los ordena por número y posición para asignar
    los gráficos en el orden correcto.
    """
    log(f"DEBUG: add_charts - Buscando placeholders para {len(charts)} gráficos: {list(charts.keys())}")
    
    placeholder_candidates = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text or ""
        if "{{ph_chart" in text:
            index = _extract_chart_placeholder_index(shape)
            placeholder_candidates.append((shape, index))
            log(
                f"DEBUG: add_charts - Encontrado marcador en '{shape.name}' con texto '{text.strip()}' index={index}"
            )

    # Ordena por índice numérico y posición visual
    placeholder_candidates.sort(
        key=lambda item: (
            item[1] if item[1] is not None else float("inf"),
            item[0].top,
            item[0].left,
        )
    )
    placeholder_candidates = [shape for shape, _ in placeholder_candidates]
    log(f"DEBUG: add_charts - Total marcadores {{ph_chart}} encontrados: {len(placeholder_candidates)}")

    if len(charts) > len(placeholder_candidates):
        log(
            f"WARNING: add_charts - hay {len(charts)} gráficos pero solo {len(placeholder_candidates)} marcadores {{ph_chart}} encontrados"
        )

    # Tomamos tantos placeholders como gráficos tengamos para evitar errores de índice
    chart_placeholders = placeholder_candidates[: len(charts)]

    for (name, chart_info), placeholder in zip(charts.items(), chart_placeholders):
        log(f"DEBUG: add_charts - Insertando gráfico '{name}' en shape '{placeholder.name}'")
        if not chart_info.get("title") and not chart_info.get("titulo") and name:
            chart_info["title"] = name.replace("_", " ").capitalize()

        # fn = os.path.join(DATA_DIR, f"{name}.png")
        fn = f"/tmp/{name}.png"
        create_matplotlib_chart(chart_info, friendly_names, fn)
        insert_image_scaled_by_width(slide, placeholder, fn)