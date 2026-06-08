#!/usr/bin/env python3
"""
generate_graf.py
Lee datos desde /data/input.json (escrito por el nodo Code de n8n),
genera 3 gráficos de torta y devuelve un JSON por stdout con las
imágenes en base64.
"""

import sys
import json
import io
import base64
from collections import defaultdict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches


# ── Paleta de colores ───────────────────────────────────────────────────────
PALETTE = [
    "#3B82F6", "#22C55E", "#EAB308", "#A855F7", "#EF4444",
    "#F97316", "#06B6D4", "#EC4899", "#14B8A6", "#8B5CF6",
    "#84CC16", "#F59E0B", "#6366F1", "#10B981", "#F43F5E",
    "#0EA5E9", "#D946EF", "#FB923C", "#34D399", "#818CF8",
]

PIE_KWARGS = dict(
    autopct="%1.1f%%",
    startangle=140,
    pctdistance=0.78,
    wedgeprops=dict(linewidth=1.5, edgecolor="white"),
    textprops=dict(fontsize=9, color="white", fontweight="bold"),
)


def consolidate(labels, values):
    """Suma entradas con la misma etiqueta (case-insensitive)."""
    grouped = defaultdict(float)
    for lbl, val in zip(labels, values):
        grouped[lbl.strip().title()] += val
    items = sorted(grouped.items(), key=lambda x: -x[1])
    return [i[0] for i in items], [i[1] for i in items]


def colors_for(n):
    return [PALETTE[i % len(PALETTE)] for i in range(n)]


def make_pie(ax, labels, values, title, unit="atenciones"):
    labels, values = consolidate(labels, values)
    total = sum(values)
    colors = colors_for(len(labels))

    wedges, texts, autotexts = ax.pie(
        values, labels=None, colors=colors, **PIE_KWARGS
    )

    # Ocultar porcentajes muy pequeños para no solapar
    for at, val in zip(autotexts, values):
        if val / total < 0.02:
            at.set_text("")

    ax.set_title(title, fontsize=12, fontweight="bold", pad=14, color="#1E293B")

    # Leyenda con etiqueta + valor + %
    legend_labels = [
        f"{lbl}: {int(val) if val == int(val) else round(val, 1)} ({val/total*100:.1f}%)"
        for lbl, val in zip(labels, values)
    ]
    patches = [mpatches.Patch(color=c, label=l)
               for c, l in zip(colors, legend_labels)]
    ax.legend(
        handles=patches,
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        fontsize=7.5,
        frameon=False,
    )

    ax.text(
        0, -1.35,
        f"Total de {unit}: {int(total) if total == int(total) else round(total, 1)}",
        ha="center", fontsize=9, color="#475569", fontweight="bold",
        transform=ax.transData,
    )


def fig_to_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def main():
    try:
        # Lee el JSON del primer argumento
        data = json.loads(sys.argv[1])
    except (IndexError, json.JSONDecodeError) as e:
        print(json.dumps({"error": f"Error leyendo argumento: {str(e)}"}))
        sys.exit(1)

    # n8n puede mandar lista o dict directo
    if isinstance(data, list):
        data = data[0]

    # Validar que existan los tres charts
    for key in ("chart_productos", "chart_horas", "chart_clientes"):
        if key not in data:
            print(json.dumps({"error": f"Falta la clave '{key}' en el JSON"}))
            sys.exit(1)

    cp = data["chart_productos"]
    ch = data["chart_horas"]
    cc = data["chart_clientes"]

    output = {}

    # ── Gráfico 1: Productos ────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="white")
    make_pie(ax, cp["labels"], cp["values"],
             "Porcentaje de atención por producto - Mensual")
    output["chart_productos_b64"] = fig_to_b64(fig)
    plt.close(fig)

    # ── Gráfico 2: Horas por persona ────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="white")
    make_pie(ax, ch["labels"], ch["values"],
             "Porcentaje de horas mensuales trabajadas por persona",
             unit="horas trabajadas")
    output["chart_horas_b64"] = fig_to_b64(fig)
    plt.close(fig)

    # ── Gráfico 3: Clientes ─────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="white")
    make_pie(ax, cc["labels"], cc["values"],
             "Porcentaje de atención por cliente")
    output["chart_clientes_b64"] = fig_to_b64(fig)
    plt.close(fig)

    # ── Devolver resultado por stdout para que n8n lo capture ───────────────
    print(json.dumps(output))


if __name__ == "__main__":
    main()
