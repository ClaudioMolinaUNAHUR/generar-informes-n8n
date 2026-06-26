import os
import re
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import textwrap
from utils.helpers import insert_image_scaled_by_width, log

EMU_PER_INCH = 914400

def create_matplotlib_chart(chart_info, friendly_names, output_file, target_size=None):
    if target_size is None:
        target_size = (10, 5)

    width, height = target_size
    width = max(5, width)
    height = max(3, height)
    fig, ax = plt.subplots(figsize=(width, height), dpi=150)
    ctype = chart_info.get("type")
    chart_title = chart_info.get("title") or chart_info.get("titulo") or "<no title>"
    use_adaptive_scale = False

    def _normalized_chart_name(value):
        return re.sub(r"\s+", " ", str(value or "").replace("_", " ").strip().lower())

    def _compact_tick_label(value, _pos=None):
        abs_value = abs(value)
        if abs_value >= 1_000_000:
            return f"{value / 1_000_000:.1f}M"
        if abs_value >= 1_000:
            return f"{value / 1_000:.0f}K"
        return f"{int(value)}"

    def _should_use_adaptive_scale(title, series_values):
        # Antes esto solo se activaba para el gráfico "Alertas Controles".
        # Lo generalizamos: cualquier gráfico de barras donde una serie sea
        # mucho más chica que otra (p.ej. 9 vs 26777) sufre el mismo problema
        # de barras "invisibles" en escala lineal, así que aplicamos la
        # misma corrección sin importar el nombre del gráfico.
        positives = [v for values in series_values for v in values if v and v > 0]
        if len(positives) < 2:
            return False
        min_positive = min(positives)
        max_positive = max(positives)
        return min_positive > 0 and (max_positive / min_positive) >= 8

    # Título si viene en la definición del gráfico
    # title = chart_info.get("title") or chart_info.get("titulo") or ""
    # if title:
    #     plt.title(title, fontsize=20, fontweight="bold", pad=20)

    # Usar tamaños de fuente fijos para evitar solapamientos variables
    # Reducidos por defecto para mantener legibilidad y evitar solapamientos
    DEFAULT_X_TICK_SIZE = 8
    DEFAULT_Y_TICK_SIZE = 8
    DEFAULT_LEGEND_SIZE = 8
    DEFAULT_LABEL_SIZE = 11
    tick_fontsize = DEFAULT_X_TICK_SIZE
    legend_fontsize = DEFAULT_LEGEND_SIZE
    label_fontsize = DEFAULT_LABEL_SIZE

    labels = chart_info.get("labels", [])
    x = range(len(labels))

    # Mapa de etiquetas amigables para claves comunes
    def _friendly_label(key, value):
        if isinstance(value, dict):
            return value.get("label") or value.get("title") or key.replace("_", " ").capitalize()
        return value

    flat_friendly_names = {
        key: _friendly_label(key, value)
        for chart in friendly_names.values()
        for key, value in chart.items()
    }

    # Paleta de colores para series (se rotan si hay más series)
    palette = [
        "#4F81BD",  # azul
        "#9BBB59",  # verde
        "#8063a1",  # violeta
        "#87C2D3",  # naranja
        "#f79546",  # gris oscuro
        "#c43d3d",  # rojo/  # azul fuerte
        "#e7cf47",
        "#ee28b2",
    ]

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

    def _normalize_series_vals(values, labels_len):
        vals = list(values or [])
        if len(vals) > labels_len:
            vals = vals[:labels_len]
        if len(vals) < labels_len:
            vals.extend([0] * (labels_len - len(vals)))
        return vals

    def _has_visible_data(values):
        return any(v not in (None, 0) for v in values)

    # Mostrar solo series que realmente tienen datos (evita leyenda con series vacías).
    visible_series_keys = []
    for key in series_keys:
        vals = _normalize_series_vals(chart_info.get(key), len(labels))
        if _has_visible_data(vals):
            visible_series_keys.append(key)

    if ctype == "bar":
        # Barras agrupadas: calcular offsets según cantidad de series
        n = len(visible_series_keys)
        if n == 0:
            # nada que dibujar
            return

        ind = np.arange(len(labels))  # posiciones base
        total_width = 0.7
        bar_width = total_width / n
        plotted_series = []
        use_adaptive_scale = False

        for idx, key in enumerate(visible_series_keys):
            vals = list(chart_info.get(key) or [])
            if len(vals) > len(labels):
                log(
                    f"DEBUG: create_matplotlib_chart - chart '{chart_title}' serie '{key}' tiene {len(vals)} valores pero labels {len(labels)}; truncando a labels"
                )
            vals = _normalize_series_vals(vals, len(labels))
            plotted_series.append(vals)

            label_full = flat_friendly_names.get(key, key.replace("_", " ").capitalize())
            label = textwrap.fill(label_full, width=20)

            color = palette[idx % len(palette)]

            offset = (idx - (n - 1) / 2) * bar_width
            positions = ind + offset
            ax.bar(positions, vals, bar_width * 0.95, label=label, color=color, edgecolor="none", zorder=3)

        # Estética general: sacar bordes superior/derecho y suavizar el resto
        # para que el gráfico se vea más liviano (no afecta colores de series).
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#d9d9d9")
        ax.spines["bottom"].set_color("#d9d9d9")

        use_adaptive_scale = _should_use_adaptive_scale(chart_title, plotted_series)
        # Preparar etiquetas X: acortar o romper para evitar solapamientos
        def _shorten_label(lbl):
            m = re.match(r"^\s*Semana\s*(\d+)\s*$", str(lbl), flags=re.IGNORECASE)
            if m:
                # Mantener la palabra 'Semana' pero partir en dos líneas para evitar solapamiento
                return f"Semana\n{m.group(1)}"
            s = str(lbl)
            if len(s) > 12 and " " in s:
                # Insertar salto de línea en la primera separación para reducir ancho
                parts = s.split(" ", 1)
                return parts[0] + "\n" + parts[1]
            if len(s) > 12:
                return s[:10] + "..."
            return s

        display_labels = [_shorten_label(l) for l in labels]
        x_tick_fontsize = tick_fontsize
        if use_adaptive_scale:
            x_tick_fontsize = max(8, tick_fontsize - 2)

        ax.set_xticks(ind)
        ax.set_xticklabels(display_labels, rotation=0, fontsize=x_tick_fontsize, ha="center")
        ax.set_axisbelow(True)
        ax.grid(axis="y", linestyle="-", color="#ececec", linewidth=0.8, zorder=0)
        # Forzar tamaño de fuente constante para ticks e impedir solapamiento
        ax.tick_params(axis="x", labelsize=tick_fontsize)
        ax.tick_params(axis="y", labelsize=DEFAULT_Y_TICK_SIZE)
        if use_adaptive_scale:
            positive_values = [v for values in plotted_series for v in values if v and v > 0]
            min_positive = min(positive_values)
            max_positive = max(positive_values)
            # linthresh marca el límite entre la zona lineal (cerca de 0) y la
            # logarítmica: lo fijamos en la potencia de 10 más cercana al valor
            # mínimo positivo para que las barras chicas no queden pegadas al 0.
            linthresh = max(1, 10 ** np.floor(np.log10(min_positive)))
            ax.set_yscale("symlog", linthresh=linthresh)

            # Ticks en potencias de 10 entre el linthresh y el máximo: se ven
            # parejos y prolijos en vez de saltar entre min/mediana/max.
            decade_start = int(np.floor(np.log10(linthresh)))
            decade_end = int(np.ceil(np.log10(max_positive)))
            tick_values = [0] + [10 ** d for d in range(decade_start, decade_end + 1)]
            if max_positive not in tick_values:
                tick_values.append(max_positive)
            tick_values = sorted(set(tick_values))

            ax.set_yticks(tick_values)
            ax.set_ylim(0, max_positive * 1.25)
            ax.yaxis.set_major_formatter(mtick.FuncFormatter(_compact_tick_label))
            ax.tick_params(axis="y", labelsize=max(8, tick_fontsize - 2))

        # Colocar siempre la leyenda a la derecha; ajustar tamaño para marcadores estrechos
        needs_more_bottom = any("\n" in d for d in display_labels)
        bottom_margin = 0.22 if needs_more_bottom else 0.18
        # Ajustar tamaño de la leyenda según ancho del placeholder (reducido 2 pts)
        legend_font = 7 if width < 7 else 8
        legend_bbox = (1.02, 0.5)
        ax.legend(
            loc="center left",
            bbox_to_anchor=legend_bbox,
            frameon=False,
            labelspacing=0.6,
            fontsize=legend_font,
            ncol=1,
            borderaxespad=0.5,
        )
        # Reservar espacio a la derecha para la leyenda y dejar margen inferior para etiquetas X
        right_margin = 0.78 if width >= 8 else 0.72
        # para marcadores muy estrechos, ampliar el derecho para que la leyenda no tape el gráfico
        if width < 5:
            right_margin = 0.68
        fig.subplots_adjust(right=right_margin, bottom=bottom_margin)
    elif ctype == "line":
        for idx, key in enumerate(visible_series_keys):
            vals = list(chart_info.get(key) or [])
            vals = vals[:len(labels)]
            if len(vals) < len(labels):
                vals.extend([None] * (len(labels) - len(vals)))

            label_full = flat_friendly_names.get(key, key.replace("_", " ").capitalize())
            label = textwrap.fill(label_full, width=20)

            color = palette[idx % len(palette)]
            ax.plot(x, vals, label=label, marker="o", markersize=5, linewidth=2.2, color=color, zorder=3)

        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#d9d9d9")
        ax.spines["bottom"].set_color("#d9d9d9")

        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, fontsize=tick_fontsize)
        ax.set_axisbelow(True)
        ax.grid(axis="y", linestyle="-", color="#ececec", linewidth=0.8, zorder=0)
        ax.legend(loc="best", fontsize=legend_fontsize)

    # Formato eje Y con separador de miles y margen superior
    try:
        if not use_adaptive_scale:
            # Usar MaxNLocator para obtener ticks "bonitos" (enteros) y
            # formateador compacto para mostrar K/M de forma consistente.
            locator = mtick.MaxNLocator(nbins=6, integer=True)
            ax.yaxis.set_major_locator(locator)
            ax.yaxis.set_major_formatter(mtick.FuncFormatter(_compact_tick_label))

            # A veces MaxNLocator puede producir ticks muy cercanos o con
            # valores repetidos tras el redondeo; deduplicarlos para evitar
            # etiquetas idénticas (ej. dos '2K').
            try:
                raw_ticks = list(ax.get_yticks())
                # Deduplicar en base a la etiqueta final que se mostrará
                # (p.ej. '2K'). Así evitamos que dos ticks distintos que
                # formatean a la misma etiqueta aparezcan repetidos.
                seen_labels = set()
                unique_ticks = []
                for t in raw_ticks:
                    lbl = _compact_tick_label(t)
                    if lbl not in seen_labels:
                        seen_labels.add(lbl)
                        unique_ticks.append(t)
                if len(unique_ticks) >= 2:
                    ax.set_yticks(unique_ticks)
            except Exception:
                pass

        current_ylim = ax.get_ylim()
        # Asegurar un límite superior razonable y evitar que el tope sea 0
        new_upper = max(current_ylim[1] * 1.15, 1)
        ax.set_ylim(current_ylim[0], new_upper)
    except Exception:
        pass

    fig.tight_layout(rect=[0, 0.02, 0.92, 0.95])
    log(
        f"DEBUG: create_matplotlib_chart - saving chart '{chart_title}' type={ctype} labels={labels} series_keys={visible_series_keys} target_size=({width},{height})"
    )
    fig.savefig(output_file, transparent=True, bbox_inches="tight", pad_inches=0.04)
    plt.close(fig)


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

        fn = f"/tmp/{name}.png"
        target_width = placeholder.width / EMU_PER_INCH
        target_height = placeholder.height / EMU_PER_INCH
        create_matplotlib_chart(
            chart_info,
            friendly_names,
            fn,
            target_size=(target_width, target_height),
        )
        insert_image_scaled_by_width(slide, placeholder, fn)