import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import textwrap
from utils.helpers import insert_image_scaled_by_width 

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
            vals = vals[:len(labels)]  # truncate if longer
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

