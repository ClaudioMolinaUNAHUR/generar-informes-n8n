"""
services/graf_service.py
Lógica de generación de gráficos de torta (pie charts).
Extraída de generate_graf.py para uso como servicio interno de la API.
"""

import io
import base64
from collections import defaultdict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches


# ── Paleta de colores ───────────────────────────────────────────────────────
PALETTE = [
    "#3B82F6",  # 1. Azul Principal (Datos estándar)
    "#10B981",  # 2. Verde Esmeralda (Estado óptimo / "Al día")
    "#F59E0B",  # 3. Ámbar Intenso (Alertas / "En Atención")
    "#EF4444",  # 4. Rojo Alerta (Crítico / Retrasos)
    "#8B5CF6",  # 5. Violeta Real (Contraste independiente)
    "#06B6D4",  # 6. Cian Vivo
    "#F97316",  # 7. Naranja Corporativo
    "#EC4899",  # 8. Rosa Intenso
    "#6366F1",  # 9. Índigo Profundo
    "#84CC16",  # 10. Verde Lima
    "#14B8A6",  # 11. Turquesa
    "#D946EF",  # 12. Fucsia Eléctrico
    "#0EA5E9",  # 13. Azul Cielo
    "#FB923C",  # 14. Naranja Claro
    "#A855F7",  # 15. Púrpura Vibrante
    "#34D399",  # 16. Menta Suave
    "#EAB308",  # 17. Amarillo Sol
    "#F43F5E",  # 18. Rosa Coral
    "#818CF8",  # 19. Lavanda Claro
    "#475569"   # 20. Pizarra Neutro (Ideal para la categoría "Otros" o históricos)
]

PIE_KWARGS = dict(
    autopct="%1.1f%%",
    startangle=140,
    pctdistance=0.78,
    wedgeprops=dict(linewidth=1.5, edgecolor="white"),
    textprops=dict(fontsize=9, color="white", fontweight="bold"),
)


def consolidate(labels: list[str], values: list[float]) -> tuple[list[str], list[float]]:
    """Suma entradas con la misma etiqueta (case-insensitive) y ordena de mayor a menor."""
    grouped = defaultdict(float)
    for lbl, val in zip(labels, values):
        grouped[lbl.strip().title()] += val
    items = sorted(grouped.items(), key=lambda x: -x[1])
    return [i[0] for i in items], [i[1] for i in items]


def colors_for(n: int) -> list[str]:
    """Devuelve n colores de la paleta, ciclando si es necesario."""
    return [PALETTE[i % len(PALETTE)] for i in range(n)]


def make_pie(
    ax: plt.Axes,
    labels: list[str],
    values: list[float],
    title: str,
    unit: str = "atenciones",
) -> None:
    """Dibuja un gráfico de torta en el Axes recibido."""
    labels, values = consolidate(labels, values)
    total = sum(values)
    colors = colors_for(len(labels))

    wedges, texts, autotexts = ax.pie(
        values, labels=None, colors=colors, **PIE_KWARGS
    )

    # Ocultar porcentajes muy pequeños para no solapar
    for at, val in zip(autotexts, values):
        if total > 0 and val / total < 0.02:
            at.set_text("")

    ax.set_title(title, fontsize=12, fontweight="bold", pad=14, color="#1E293B")

    # Leyenda con etiqueta + valor + %
    legend_labels = [
        f"{lbl}: {int(val) if val == int(val) else round(val, 1)} ({val / total * 100:.1f}%)" if total > 0 else f"{lbl}: 0 (0.0%)"
        for lbl, val in zip(labels, values)
    ]
    patches = [
        mpatches.Patch(color=c, label=l)
        for c, l in zip(colors, legend_labels)
    ]
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
        ha="center",
        fontsize=9,
        color="#475569",
        fontweight="bold",
        transform=ax.transData,
    )


def fig_to_b64(fig: plt.Figure) -> str:
    """Convierte una figura matplotlib a string base64 PNG."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def normalize_input(data) -> dict:
    """
    Normaliza el payload de entrada:
    - Si viene como lista, toma el primer elemento.
    - Acepta claves con prefijo 'chart_' o 'graf_' indistintamente.
    """
    if isinstance(data, list):
        data = data[0] if data else {}

    # Alias: si vienen con prefijo 'chart_', renombrar a 'graf_'
    for old, new in (
        ("chart_productos", "graf_productos"),
        ("chart_horas",     "graf_horas"),
        ("chart_clientes",  "graf_clientes"),
        ("chart_tickets",   "graf_tickets"),
    ):
        if old in data and new not in data:
            data[new] = data.pop(old)

    return data


def generate_grafs(data) -> dict:
    """
    Genera los cuatro gráficos de torta a partir del dict de entrada.

    Acepta el payload como lista o dict, y claves con prefijo
    'graf_' o 'chart_' (compatibilidad con n8n).

    Parámetros esperados (tras normalización):
        graf_productos: { labels: [...], values: [...] }
        graf_horas:     { labels: [...], values: [...] }
        graf_clientes:  { labels: [...], values: [...] }
        graf_tickets:   { labels: [...], values: [...] }

    Retorna un dict con las claves:
        graf_productos_b64, graf_horas_b64, graf_clientes_b64, graf_tickets_b64
    """
    data = normalize_input(data)

    required_keys = ("graf_productos", "graf_horas", "graf_clientes", "graf_tickets")
    for key in required_keys:
        if key not in data:
            raise ValueError(f"Falta la clave '{key}' en el JSON de entrada")

    cp = data["graf_productos"]
    ch = data["graf_horas"]
    cc = data["graf_clientes"]
    ct = data["graf_tickets"]

    output = {}

    # ── Gráfico 1: Productos ────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="white")
    make_pie(
        ax,
        cp["labels"],
        cp["values"],
        "Porcentaje de atención por producto",
    )
    output["graf_productos_b64"] = fig_to_b64(fig)
    plt.close(fig)

    # ── Gráfico 2: Horas por persona ────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="white")
    make_pie(
        ax,
        ch["labels"],
        ch["values"],
        "Porcentaje de horas trabajadas por área",
        unit="horas trabajadas",
    )
    output["graf_horas_b64"] = fig_to_b64(fig)
    plt.close(fig)

    # ── Gráfico 3: Clientes ─────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="white")
    make_pie(
        ax,
        cc["labels"],
        cc["values"],
        "Porcentaje de atención por cliente",
    )
    output["graf_clientes_b64"] = fig_to_b64(fig)
    plt.close(fig)

    # ── Gráfico 4: Tickets (Estados Abierto/Concluido) ──────────────────────
    fig, ax = plt.subplots(figsize=(9, 6), facecolor="white")
    make_pie(
        ax,
        ct["labels"],
        ct["values"],
        "Porcentaje de tickets trabajados esta semana",
        unit="tickets",
    )
    output["graf_tickets_b64"] = fig_to_b64(fig)
    plt.close(fig)

    return output