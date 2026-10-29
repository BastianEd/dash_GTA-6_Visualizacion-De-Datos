# -*- coding: utf-8 -*-
"""
Dashboard GTA VI para Rockstar
Asignatura: ADY1104 Visualización de Datos
Pregunta central: ¿Qué tema debería priorizar Rockstar al comunicar GTA VI?

Uso:
1. Instala dependencias: pip install dash pandas openpyxl plotly
2. Asegúrate de tener el archivo "GTA_VI_Datos_Visualizacion.xlsx" en la carpeta data/
3. Ejecuta el script: python dashboard_gta_vi.py
"""

from __future__ import annotations

import math
from datetime import date
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import plotly.graph_objects as go
from dash import Dash, Input, Output, dash_table, dcc, html

# ---------------------------------------------------------------------------
# 1. Configuración y Constantes
# ---------------------------------------------------------------------------
NOMBRE_ARCHIVO = "GTA_VI_Datos_Visualizacion.xlsx"
CANDIDATAS = [
    Path.cwd() / "data" / NOMBRE_ARCHIVO,
    Path.cwd() / NOMBRE_ARCHIVO,
    Path.cwd().parent / "data" / NOMBRE_ARCHIVO,
]
DATA_PATH = next((p for p in CANDIDATAS if p.exists()), Path.cwd() / "data" / NOMBRE_ARCHIVO)

MIN_N = 30       # Bajo este tamaño de grupo se muestra una advertencia
MIN_CELDA = 10   # En el mapa de calor, celdas con menos comentarios no se pintan

# Fecha oficial de lanzamiento anunciada por Rockstar
LANZAMIENTO = date(2026, 11, 19)

ORDEN = ["Positivo", "Neutral", "Negativo"]
COLORES = {"Positivo": "#2EE59D", "Neutral": "#8E97B5", "Negativo": "#FF3C8E"}
PATRONES = {"Positivo": "+", "Neutral": "", "Negativo": "/"}

# Paleta "Vice City": neones sobre noche violeta
NEON = dict(rosa="#FF3C8E", cian="#22D8F0", naranja="#FF9A3C", violeta="#A66BFF", verde="#2EE59D", amarillo="#FFD23C")
TINTA, ACENTO, FONDO, BORDE, SUAVE = "#F5F1FF", NEON["rosa"], "#0B0A22", "#3A3478", "#B9BCE0"
PANEL = "#15123D"
F_DISPLAY = "'Anton', 'Bebas Neue', Impact, sans-serif"
F_TITULO = "'Space Grotesk', 'Segoe UI', sans-serif"
F_TEXTO = "'Source Sans 3', 'Segoe UI', sans-serif"
FUENTES = (
    "https://fonts.googleapis.com/css2?family=Anton&family=Space+Grotesk:wght@500;700"
    "&family=Source+Sans+3:wght@400;600&display=swap"
)

TITULAR = (
    "Rockstar puede comunicar con confianza la historia y los personajes, "
    "pero primero debe aclarar las dudas sobre plataformas y rendimiento."
)

ID_VOZ_CRITICA = "85b1c80e413dd03c99fa"
ID_VOZ_FAVORABLE = "8e924a1224d28a8b795b"
ID_REVISAR = "d3d2c52628d9a30ef064"


# ---------------------------------------------------------------------------
# 2. Carga de Datos y Funciones de Métricas
# ---------------------------------------------------------------------------
def cargar_datos(ruta: Path = DATA_PATH) -> pd.DataFrame:
    """Lee la hoja 'Datos' del Excel del curso."""
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo {ruta.name}. Verifique que esté en la carpeta data/.")
    return pd.read_excel(ruta, sheet_name="Datos")


def pct(valor: float) -> str:
    """Formato de porcentaje con punto decimal: 8.1%."""
    return f"{valor:.1f}%"


def filtrar(df: pd.DataFrame, **seleccion: list[str] | None) -> pd.DataFrame:
    """Aplica filtros por columna; una selección vacía significa 'todos'."""
    mask = pd.Series(True, index=df.index)
    for columna, valores in seleccion.items():
        if valores:
            mask &= df[columna].isin(valores)
    return df[mask]


def composicion(df: pd.DataFrame, columna: str) -> pd.DataFrame:
    """% de cada sentimiento por grupo (cada fila suma 100) más el tamaño n."""
    conteo = pd.crosstab(df[columna], df["sentimiento"]).reindex(columns=ORDEN, fill_value=0)
    n = conteo.sum(axis=1)
    return conteo.div(n, axis=0).mul(100).round(1).assign(n=n)


def contraste_por_tema(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve n, % Positivo, % Neutral, % Negativo, balance neto y polarización por tema."""
    c = composicion(df, "tema_principal")
    return c.assign(
        balance=(c["Positivo"] - c["Negativo"]).round(1),
        polarizacion=(c["Positivo"] + c["Negativo"]).round(1),
    ).sort_values("balance")


def preguntas_por_tema(df: pd.DataFrame) -> pd.DataFrame:
    """% de comentarios con signo de pregunta por tema, con tamaño n."""
    es_preg = df["es_pregunta"].eq("Sí")
    g = es_preg.groupby(df["tema_principal"])
    return pd.DataFrame({"n": g.size(), "pct": g.mean().mul(100).round(1)}).sort_values("pct")


def celda_mas_critica(df: pd.DataFrame):
    """Celda tema x video con mayor % negativo entre las que tienen al menos MIN_CELDA comentarios."""
    g = df.groupby(["tema_principal", "video"])
    t = pd.DataFrame({"n": g.size(), "neg": g["sentimiento"].apply(lambda s: int(s.eq("Negativo").sum()))})
    t = t[t["n"] >= MIN_CELDA]
    if t.empty:
        return None
    t = t.assign(p=t["neg"] / t["n"] * 100).sort_values(["p", "n"], ascending=False)
    (tema, video), f = next(t.iterrows())
    return tema, video, float(f["p"]), int(f["n"])


def estrellas(pct_neg: float) -> int:
    """Nivel de búsqueda: 1 estrella por cada 5 puntos de % negativo (máximo 5)."""
    return min(5, math.ceil(pct_neg / 5))


def cuenta_regresiva(hoy: date | None = None) -> str:
    """Texto de la cuenta regresiva hacia LANZAMIENTO."""
    dias = (LANZAMIENTO - (hoy or date.today())).days
    if dias > 1:
        return f"Faltan {dias} días · {LANZAMIENTO:%d/%m/%Y}"
    if dias == 1:
        return "Sale mañana"
    if dias == 0:
        return "Sale hoy"
    return f"Salió hace {-dias} días"


def indicadores(df: pd.DataFrame) -> list[tuple]:
    """Cuatro indicadores de la vista ejecutiva: (valor, etiqueta, nota, barra %, color)."""
    s = df["sentimiento"].value_counts(normalize=True).mul(100)
    n_preg = int(df["es_pregunta"].eq("Sí").sum())
    p_pos, p_neg, p_preg = s.get("Positivo", 0), s.get("Negativo", 0), n_preg / len(df) * 100
    return [
        (f"{len(df)}", "comentarios analizados", f"{df['video'].nunique()} videos públicos", 100, NEON["cian"]),
        (pct(p_pos), "positivos", f"Neutral: {pct(s.get('Neutral', 0))}", p_pos, COLORES["Positivo"]),
        (pct(p_neg), "negativos", "etiqueta asistida, sin validar", p_neg, COLORES["Negativo"]),
        (pct(p_preg), "son preguntas", f"{n_preg} comentarios", p_preg, NEON["amarillo"]),
    ]


# ---------------------------------------------------------------------------
# 3. Funciones Gráficas
# ---------------------------------------------------------------------------
def _estilo(fig: go.Figure, alto: int) -> go.Figure:
    """Estilo común (modo noche): fuente clara, fondo transparente y leyenda inferior."""
    fig.update_layout(
        height=alto, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family=F_TEXTO, size=13, color=TINTA),
        margin=dict(l=10, r=30, t=10, b=40),
        legend=dict(orientation="h", y=-0.18, x=0, font=dict(color=TINTA)),
        hoverlabel=dict(bgcolor=PANEL, bordercolor=ACENTO, font=dict(family=F_TEXTO, color=TINTA)),
    )
    fig.update_xaxes(gridcolor="#2A2660", zeroline=False, linecolor=BORDE)
    fig.update_yaxes(automargin=True, linecolor=BORDE)
    return fig


def _etiquetas(n: pd.Series) -> list[str]:
    """Etiqueta del grupo con su n; ⚠ si la base es menor a MIN_N."""
    return [f"{g} (n={v}){' ⚠' if v < MIN_N else ''}" for g, v in n.items()]


def _barra(sent: str, **kw) -> go.Bar:
    """Barra horizontal de un sentimiento con color y patrón consistentes."""
    pattern = dict(shape=PATRONES[sent], bgcolor=COLORES[sent], fgcolor=FONDO, size=6)
    return go.Bar(orientation="h", name=sent, marker=dict(color=COLORES[sent], pattern=pattern), **kw)


def _mezcla(a: str, b: str, t: float) -> str:
    """Mezcla dos colores hex (t=0 da a, t=1 da b)."""
    ca, cb = [int(a[i:i + 2], 16) for i in (1, 3, 5)], [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def _color_balance(balance: float, lim: float) -> str:
    """Rosa si hay más negativos que positivos, gris si empatan, verde si hay más positivos."""
    t = (max(-lim, min(lim, balance)) + lim) / (2 * lim)
    if t < 0.5:
        return _mezcla(COLORES["Negativo"], COLORES["Neutral"], t * 2)
    return _mezcla(COLORES["Neutral"], COLORES["Positivo"], (t - 0.5) * 2)


def fig_divergente(c: pd.DataFrame) -> go.Figure:
    """Barras divergentes: negativo a la izquierda, positivo a la derecha."""
    y = _etiquetas(c["n"])
    fig = go.Figure()
    for sent, signo in (("Negativo", -1), ("Positivo", 1)):
        fig.add_trace(_barra(
            sent, y=y, x=signo * c[sent], text=[pct(v) for v in c[sent]],
            textposition="outside", cliponaxis=False, customdata=c[["Neutral"]],
            hovertemplate=f"%{{y}}<br>{sent}: %{{text}}<br>Neutral: %{{customdata[0]:.1f}}%<extra></extra>",
        ))
    lim = (int(max(c["Positivo"].max(), c["Negativo"].max()) // 10) + 1) * 10
    ticks = list(range(-lim, lim + 1, 10))
    fig.update_layout(barmode="relative")
    fig.update_xaxes(tickvals=ticks, ticktext=[f"{abs(t)}%" for t in ticks],
                     range=[-lim, lim], title="← Negativo   |   Positivo →")
    fig.add_vline(x=0, line_width=1, line_color=TINTA)
    return _estilo(fig, 60 + 44 * len(c))


def fig_composicion(c: pd.DataFrame) -> go.Figure:
    """Barras apiladas al 100 % con la composición de sentimiento por grupo."""
    y = _etiquetas(c["n"])
    fig = go.Figure([
        _barra(s, y=y, x=c[s], text=[pct(v) if v >= 4 else "" for v in c[s]],
               textposition="inside", insidetextanchor="middle")
        for s in ORDEN
    ])
    fig.update_layout(barmode="stack")
    fig.update_xaxes(range=[0, 100], tickvals=[0, 25, 50, 75, 100],
                     ticktext=["0%", "25%", "50%", "75%", "100%"])
    return _estilo(fig, 90 + 56 * len(c))


def fig_simple(valores: pd.Series, n: pd.Series, texto: list[str]) -> go.Figure:
    """Barras horizontales de un solo color."""
    fig = go.Figure(go.Bar(
        y=_etiquetas(n), x=valores, orientation="h", marker_color=NEON["cian"],
        text=texto, textposition="outside", cliponaxis=False,
    ))
    fig.update_xaxes(range=[0, valores.max() * 1.25 if len(valores) else 1])
    return _estilo(fig, 60 + 36 * len(valores))


def fig_mapa_tension(c: pd.DataFrame) -> go.Figure:
    """Mapa de Leonida: cuánto se habla de cada tema (x) frente a cuánto divide (y)."""
    total = c["n"].sum()
    d = c.assign(peso=(c["n"] / total * 100).round(1)).sort_values("peso")
    d = d.assign(pos=(["top center", "bottom center"] * len(d))[:len(d)])
    prom_peso = 100 / len(d)
    prom_pol = float((d["polarizacion"] * d["n"]).sum() / total)
    lim = max(10.0, float(d["balance"].abs().max()))
    tam_ref = 2 * float(d["n"].max()) / (58 ** 2)

    fig = go.Figure()
    grupos = (
        (f"Base suficiente (n ≥ {MIN_N})", d[d["n"] >= MIN_N], False),
        (f"Base pequeña (n < {MIN_N}) ⚠", d[d["n"] < MIN_N], True),
    )
    for nombre, g, hueco in grupos:
        if g.empty:
            continue
        colores = [_color_balance(b, lim) for b in g["balance"]]
        fig.add_trace(go.Scatter(
            x=g["peso"], y=g["polarizacion"], mode="markers+text", name=nombre,
            text=[f"{t}{' ⚠' if hueco else ''}" for t in g.index], textposition=list(g["pos"]),
            textfont=dict(color=TINTA, size=12),
            marker=dict(size=list(g["n"]), sizemode="area", sizeref=tam_ref, sizemin=9,
                        color="rgba(0,0,0,0)" if hueco else colores, opacity=1 if hueco else 0.92,
                        line=dict(color=colores, width=3 if hueco else 1.5)),
            customdata=g[["n", "Positivo", "Negativo", "balance"]],
            hovertemplate=("<b>%{text}</b><br>Peso: %{x:.1f}% de los comentarios"
                           "<br>División: %{y:.1f}% (positivo + negativo)"
                           "<br>Positivo %{customdata[1]:.1f}% · Negativo %{customdata[2]:.1f}%"
                           "<br>Balance %{customdata[3]:+.1f} pp · n=%{customdata[0]}<extra></extra>"),
        ))
    fig.add_hline(y=prom_pol, line_dash="dash", line_color=NEON["cian"], opacity=0.55)
    fig.add_vline(x=prom_peso, line_dash="dash", line_color=NEON["cian"], opacity=0.55)
    for x, y, xa, ya, t in ((0.01, 0.99, "left", "top", "Nicho que divide"), (0.99, 0.99, "right", "top", "Foco masivo"),
                            (0.99, 0.01, "right", "bottom", "Conversación tranquila"), (0.01, 0.01, "left", "bottom", "Poco ruido")):
        fig.add_annotation(xref="paper", yref="paper", x=x, y=y, xanchor=xa, yanchor=ya, text=t.upper(),
                           showarrow=False, font=dict(family=F_DISPLAY, size=13, color="rgba(185,188,224,.6)"))
    fig.update_xaxes(title="Peso del tema (% de los comentarios)", ticksuffix="%",
                     range=[0, float(d["peso"].max()) * 1.18 + 2])
    fig.update_yaxes(title="División (% positivo + % negativo)", ticksuffix="%",
                     range=[0, max(float(d["polarizacion"].max()) * 1.25, prom_pol * 1.5, 10)])
    _estilo(fig, 470)
    fig.update_layout(margin=dict(l=10, r=30, t=10, b=90), legend=dict(orientation="h", y=-0.27, x=0))
    return fig


def fig_calor(df: pd.DataFrame) -> go.Figure:
    """Mapa de calor tema x video con el % de comentarios negativos."""
    n = pd.crosstab(df["tema_principal"], df["video"])
    neg = (df[df["sentimiento"].eq("Negativo")].groupby(["tema_principal", "video"]).size()
           .unstack(fill_value=0).reindex(index=n.index, columns=n.columns, fill_value=0))
    orden = (neg.sum(axis=1) / n.sum(axis=1)).sort_values(ascending=False).index
    n, neg = n.loc[orden], neg.loc[orden]
    p = neg / n * 100
    filas = [f"{t} (n={int(n.loc[t].sum())})" for t in orden]
    z = p.where(n >= MIN_CELDA)
    zmax = max(25.0, float(z.max().max())) if z.notna().any().any() else 25.0

    fig = go.Figure(go.Heatmap(
        z=z.values, x=list(n.columns), y=filas, zmin=0, zmax=zmax, xgap=4, ygap=4, hoverongaps=False,
        colorscale=[[0, "#1B1850"], [1, NEON["rosa"]]],
        colorbar=dict(title=dict(text="% negativo"), ticksuffix="%", thickness=12, len=0.85),
        hovertemplate="%{y}<br>%{x}: %{z:.1f}% negativo<extra></extra>",
    ))
    for i, tema in enumerate(orden):
        for video in n.columns:
            k = int(n.loc[tema, video])
            if k >= MIN_CELDA:
                texto, color = f"{p.loc[tema, video]:.0f}%<br><span style='font-size:10px'>n={k}</span>", TINTA
            else:
                texto, color = f"n={k}", "#6D70A0"
            fig.add_annotation(x=video, y=filas[i], text=texto, showarrow=False, font=dict(size=12, color=color))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    fig.update_xaxes(side="top", showgrid=False)
    return _estilo(fig, 120 + 48 * len(orden))


# ---------------------------------------------------------------------------
# 4. Componentes de Interfaz
# ---------------------------------------------------------------------------
S_TARJETA = dict(
    background="linear-gradient(160deg, rgba(36,31,98,.92), rgba(15,13,46,.94))",
    border=f"1px solid {BORDE}", borderRadius="14px", padding="18px 20px",
)
S_TITULO = dict(fontFamily=F_TITULO, fontSize="18px", fontWeight=700, margin="0 0 6px", color=TINTA, lineHeight="1.25")
S_NOTA = dict(fontSize="14px", color=SUAVE, margin="0 0 8px", lineHeight="1.45")


def tarjeta(titulo: str, *hijos, nota: str | None = None, acento: str = ACENTO, etiqueta: str | None = None) -> html.Section:
    cabecera = []
    if etiqueta:
        cabecera.append(html.Div(etiqueta, style=dict(fontFamily=F_DISPLAY, letterSpacing="2px", fontSize="13px",
                                                      color=acento, marginBottom="4px")))
    cabecera.append(html.H3(titulo, style=S_TITULO))
    if nota:
        cabecera.append(html.P(nota, style=S_NOTA))
    return html.Section([*cabecera, *hijos], className="tarjeta", style={
        **S_TARJETA, "borderTop": f"3px solid {acento}",
        "boxShadow": f"0 0 24px {acento}33, 0 10px 30px rgba(0,0,0,.45)",
    })


def indicador(valor: str, etiqueta: str, nota: str, barra: float | None = None, color: str | None = None) -> html.Div:
    color = color or NEON["cian"]
    hijos = [
        html.Div(valor, style=dict(fontFamily=F_DISPLAY, fontSize="44px", letterSpacing="1px", color=color,
                                   textShadow=f"0 0 18px {color}88", lineHeight="1.1")),
        html.Div(etiqueta, style=dict(fontWeight=600, textTransform="uppercase", letterSpacing="1px", fontSize="13px")),
        html.Div(nota, style=dict(fontSize="13px", color=SUAVE)),
    ]
    if barra is not None:
        relleno = html.Div(style=dict(width=f"{max(2.0, min(100.0, barra)):.1f}%", height="100%",
                                      background=color, boxShadow=f"0 0 10px {color}"))
        hijos.append(html.Div(relleno, style=dict(height="6px", background="#251F5E", borderRadius="3px",
                                                  marginTop="10px", overflow="hidden")))
    return html.Div(hijos, className="tarjeta", style={**S_TARJETA, "flex": "1 1 200px", "borderLeft": f"4px solid {color}"})


def cita(df: pd.DataFrame, id_comentario: str, etiqueta: str, color: str = ACENTO) -> html.Blockquote:
    coincidencias = df.loc[df["id_comentario"] == id_comentario]
    if coincidencias.empty:
        return html.Blockquote(f"Comentario {id_comentario} no encontrado en el dataset activo.")
    f = coincidencias.iloc[0]
    return html.Blockquote([
        html.P(f"“{f['comentario']}”", style=dict(margin="0 0 6px")),
        html.Footer(f"{etiqueta}: tema {f['tema_principal']}, {f['video']}, id {f['id_comentario']}",
                    style=dict(fontSize="13px", color=SUAVE)),
    ], style=dict(margin="0 0 12px", paddingLeft="12px", borderLeft=f"3px solid {color}"))


def tarjeta_busqueda(c: pd.DataFrame, base: pd.DataFrame) -> html.Section:
    foco = base["Negativo"].idxmax()
    k_foco = estrellas(base.loc[foco, "Negativo"])
    filas = []
    for tema, f in c.sort_values(["Negativo", "n"], ascending=False).iterrows():
        n, k = int(f["n"]), estrellas(f["Negativo"])
        filas.append(html.Div([
            html.Span(f"{tema}{' ⚠' if n < MIN_N else ''}", style=dict(flex="1 1 auto", fontWeight=600)),
            html.Span([html.Span("★" * k, className="estrella"), html.Span("☆" * (5 - k), style=dict(color="#4A4590"))],
                      title=f"{k} de 5 estrellas", style=dict(letterSpacing="2px", fontSize="18px", whiteSpace="nowrap")),
            html.Span(f"{pct(f['Negativo'])} · n={n}", style=dict(width="104px", textAlign="right", color=SUAVE, fontSize="13px")),
        ], style=dict(display="flex", alignItems="center", gap="10px", padding="7px 0", borderBottom=f"1px solid {BORDE}")))
    return tarjeta(
        f"{foco} tiene el nivel de búsqueda más alto entre los temas con base suficiente ({k_foco} de 5 ★)",
        *filas, acento=NEON["amarillo"], etiqueta="SE BUSCA",
        nota=f"1 ★ por cada 5 puntos de % negativo (máx. 5). Es un recurso visual, no una métrica validada. ⚠ = menos de {MIN_N} comentarios.",
    )


def grafico(fig: go.Figure) -> dcc.Graph:
    return dcc.Graph(figure=fig, config=dict(displayModeBar=False))


FILA = dict(display="flex", flexWrap="wrap", gap="16px", marginBottom="16px")
GRILLA = dict(display="grid", gridTemplateColumns="repeat(auto-fit, minmax(420px, 1fr))", gap="16px")
ANCHO_TOTAL = dict(gridColumn="1 / -1")


# ---------------------------------------------------------------------------
# 5. Vistas
# ---------------------------------------------------------------------------
def vista_resumen(df: pd.DataFrame) -> html.Div:
    """Vista 1: hallazgo, indicadores, comparación central, decisión y límite."""
    c = contraste_por_tema(df)
    base = c[c["n"] >= MIN_N]
    foco = base["polarizacion"].idxmax()
    seguro = base[base["Negativo"] == 0]["n"].idxmax()
    f = c.loc[foco]
    n_foco = int(f["n"])

    comparacion = tarjeta(
        f"{foco} es el tema más dividido: {pct(f['Positivo'])} a favor y {pct(f['Negativo'])} en contra",
        grafico(fig_divergente(c)),
        nota=(f"{seguro} no tiene críticas (0 de {int(c.loc[seguro, 'n'])}). "
              f"El resto de cada tema es neutral. ⚠ = menos de {MIN_N} comentarios."),
        acento=NEON["cian"], etiqueta="COMPARACIÓN CENTRAL",
    )
    lateral = html.Div([
        tarjeta("Decisión para Rockstar",
                html.P("Evaluar una pieza de comunicación que aclare qué esperar en plataformas y "
                       "rendimiento, antes de reforzar el mensaje sobre historia y personajes.", style=S_NOTA),
                acento=NEON["violeta"], etiqueta="MISIÓN"),
        tarjeta("Cómo comprobarla",
                html.P("Medir en nuevos videos el % de comentarios negativos y de preguntas sobre "
                       "plataformas y rendimiento. Si bajan, la aclaración ayudó (no prueba causa).", style=S_NOTA),
                acento=NEON["verde"], etiqueta="CHECKPOINT"),
        tarjeta("Límite de esta lectura",
                html.P(f"Con n={n_foco}, un solo comentario mueve el porcentaje {100 / n_foco:.1f} puntos. "
                       "Las etiquetas no tienen validación humana.", style=S_NOTA),
                acento=NEON["naranja"], etiqueta="ADVERTENCIA"),
    ], style=dict(display="flex", flexDirection="column", gap="16px", flex="1 1 280px"))

    voz = tarjeta("Voz de la audiencia",
                  cita(df, ID_VOZ_CRITICA, "Crítica", COLORES["Negativo"]),
                  cita(df, ID_VOZ_FAVORABLE, "Favorable", COLORES["Positivo"]),
                  acento=NEON["violeta"], etiqueta="COMENTARIOS REALES")

    return html.Div([
        html.Div([indicador(*k) for k in indicadores(df)], style=FILA),
        html.Div([html.Div(comparacion, style=dict(flex="2 1 520px")), lateral], style=FILA),
        html.Div([html.Div(tarjeta_busqueda(c, base), style=dict(flex="1 1 420px")),
                  html.Div(voz, style=dict(flex="1 1 420px"))], style=FILA),
    ])


def vista_temas(df: pd.DataFrame) -> html.Div:
    """Vista 2: mapa de tensión, sentimiento por video, mapa de calor, ranking y preguntas."""
    cv, ct = composicion(df, "video"), contraste_por_tema(df)
    neg = cv["Negativo"]
    t_video = (f"Entre videos, el sentimiento negativo va de {pct(neg.min())} a {pct(neg.max())}"
               if len(cv) > 1 else "Sentimiento del video seleccionado")

    base = ct[ct["n"] >= MIN_N]
    if base.empty:
        t_mapa = "Peso y división de cada tema"
    else:
        foco = base["polarizacion"].idxmax()
        fb = base.loc[foco]
        t_mapa = (f"{foco}: {pct(fb['n'] / ct['n'].sum() * 100)} de los comentarios "
                  f"y la mayor división ({pct(fb['polarizacion'])} positivo o negativo)")

    celda = celda_mas_critica(df)
    if celda is None:
        t_calor = "Crítica por tema y video (ninguna celda tiene base suficiente)"
    elif celda[2] == 0:
        t_calor = "Ninguna celda con base suficiente tiene comentarios negativos"
    else:
        t_calor = f"La celda más crítica es {celda[0]} en {celda[1]}: {pct(celda[2])} de {celda[3]} comentarios"

    conteo = df["tema_principal"].value_counts().sort_values()
    top = conteo.idxmax()
    nota_rank = ("“Otros temas” agrupa lo que las reglas de palabras clave no lograron clasificar."
                 if top == "Otros temas" else None)

    pt = preguntas_por_tema(df)
    elegibles = pt[pt["n"] >= MIN_N]
    t_preg = (f"{elegibles['pct'].idxmax()} tiene la mayor proporción de preguntas "
              f"({pct(elegibles['pct'].max())})" if not elegibles.empty else "Preguntas por tema")
    plat = "Plataformas y rendimiento"
    nota_preg = (f"En {plat.lower()} solo {pct(pt.loc[plat, 'pct'])} son preguntas: la inquietud "
                 "aparece más como opinión que como duda explícita." if plat in pt.index else None)

    return html.Div([
        html.Div(tarjeta(
            t_mapa, grafico(fig_mapa_tension(ct)), acento=NEON["cian"], etiqueta="MAPA DE LEONIDA",
            nota=("Cada burbuja es un tema: a la derecha se habla más de él y arriba divide más (positivo + negativo). "
                  "Tamaño = n. Color: verde = más positivos que negativos, rosa = al revés. "
                  f"Círculo hueco ⚠ = menos de {MIN_N} comentarios. Líneas punteadas = promedios."),
        ), style=ANCHO_TOTAL),
        tarjeta(t_video, grafico(fig_composicion(cv)), acento=NEON["rosa"], etiqueta="POR VIDEO",
                nota="Porcentajes por video; la base (n) cambia entre videos."),
        tarjeta(t_calor, grafico(fig_calor(df)), acento=NEON["naranja"], etiqueta="MAPA DE CALOR",
                nota=f"% de comentarios negativos por tema y video. Las celdas con menos de {MIN_CELDA} comentarios no se pintan."),
        tarjeta(f"{top} es el grupo más grande ({pct(conteo.max() / conteo.sum() * 100)})",
                grafico(fig_simple(conteo, conteo, [f"{v}" for v in conteo])), nota=nota_rank,
                acento=NEON["violeta"], etiqueta="RANKING"),
        tarjeta(t_preg, grafico(fig_simple(pt["pct"], pt["n"], [pct(v) for v in pt["pct"]])), nota=nota_preg,
                acento=NEON["amarillo"], etiqueta="PREGUNTAS"),
    ], style=GRILLA)


def vista_evidencia(df: pd.DataFrame, df_global: pd.DataFrame) -> html.Div:
    """Vista 3: comentarios seleccionados, estado de revisión y limitaciones."""
    cols = ["video", "tema_principal", "sentimiento", "es_pregunta", "comentario"]
    tabla = dash_table.DataTable(
        data=df[cols].to_dict("records"),
        columns=[{"name": c.replace("_", " ").capitalize(), "id": c} for c in cols],
        page_size=10, sort_action="native", filter_action="native",
        style_table=dict(overflowX="auto"),
        style_header=dict(fontWeight=700, backgroundColor="#221E5E", color=TINTA, fontFamily=F_TITULO,
                          border=f"1px solid {BORDE}"),
        style_filter=dict(backgroundColor="#0F0D33", color=TINTA, border=f"1px solid {BORDE}"),
        style_cell=dict(textAlign="left", whiteSpace="normal", height="auto", padding="8px",
                        fontFamily=F_TEXTO, fontSize="14px", backgroundColor=PANEL, color=TINTA,
                        border=f"1px solid {BORDE}"),
        style_cell_conditional=[{"if": {"column_id": "comentario"}, "minWidth": "340px", "width": "45%"}],
        style_data_conditional=[
            {"if": {"filter_query": '{sentimiento} = "Positivo"', "column_id": "sentimiento"},
             "color": COLORES["Positivo"], "fontWeight": 700},
            {"if": {"filter_query": '{sentimiento} = "Negativo"', "column_id": "sentimiento"},
             "color": COLORES["Negativo"], "fontWeight": 700},
            {"if": {"state": "active"}, "backgroundColor": "#2A2570", "border": f"1px solid {ACENTO}"},
        ],
    )
    sin_validar = df["estado_revision"].str.contains("sin validación").mean() * 100 if "estado_revision" in df.columns else 100.0
    n_mixtos = int(df["posible_mixto"].eq("Sí").sum()) if "posible_mixto" in df.columns else 0
    n_preg = int(df["es_pregunta"].eq("Sí").sum())
    otros = df["tema_principal"].eq("Otros temas").mean() * 100
    limites = [
        "Las etiquetas de sentimiento y tema son propuestas asistidas, sin validación humana definitiva.",
        "Son reacciones a tres videos, no la experiencia real de jugar GTA VI.",
        "La muestra no representa a toda la comunidad; no se pueden sacar conclusiones causales.",
        "Un comentario puede tener más de un tema o postura; las reglas de palabras clave simplifican.",
    ]
    revisar = tarjeta(
        "Una etiqueta que merece revisión humana",
        cita(df_global, ID_REVISAR, "Etiquetado Positivo", NEON["naranja"]),
        html.P("El texto pide no ilusionarse con los gráficos, pero la etiqueta asistida lo marcó Positivo. "
               "Es un ejemplo de por qué el sentimiento es una señal preliminar.", style=S_NOTA),
        acento=NEON["naranja"], etiqueta="CONTRAEVIDENCIA",
    )
    return html.Div([
        html.Div([
            indicador(f"{len(df)}", "comentarios en la selección", "según los filtros activos", 100, NEON["cian"]),
            indicador(pct(sin_validar), "sin validación humana", "estado de revisión del dataset", sin_validar, NEON["naranja"]),
            indicador(f"{n_mixtos}", "señales mixtas", "heurística de palabras; puede subestimar", None, NEON["violeta"]),
            indicador(pct(otros), "en “Otros temas”", f"{n_preg} comentarios son preguntas", otros, NEON["amarillo"]),
        ], style=FILA),
        tarjeta("Comentarios seleccionados: úsalos para comprobar cada lectura", tabla,
                nota="Filtra u ordena la tabla para ubicar los comentarios que respaldan cada lectura.",
                acento=NEON["cian"], etiqueta="SALA DE EVIDENCIA"),
        html.Div([
            html.Div(tarjeta("Alcance y limitaciones",
                             html.Ul([html.Li(t) for t in limites], className="lista-neon",
                                     style=dict(margin=0, paddingLeft="20px", lineHeight="1.6")),
                             acento=NEON["rosa"], etiqueta="LÍMITES"), style=dict(flex="1 1 420px")),
            html.Div(revisar, style=dict(flex="1 1 420px")),
        ], style={**FILA, "marginTop": "16px"}),
    ])


# ---------------------------------------------------------------------------
# 6. Aplicación y Servidor Dash
# ---------------------------------------------------------------------------
CSS = r"""
:root{
  --Dash-Fill-Inverse-Strong:#15123D; --Dash-Fill-Interactive-Strong:#FF3C8E; --Dash-Fill-Interactive-Weak:#2A2570;
  --Dash-Stroke-Strong:#3A3478; --Dash-Stroke-Weak:#2A2660;
  --Dash-Text-Strong:#F5F1FF; --Dash-Text-Primary:#F5F1FF; --Dash-Text-Weak:#B9BCE0; --Dash-Text-Disabled:#6D70A0;
}
html,body{background:#0B0A22;margin:0}
*{scrollbar-color:#FF3C8E #15123D}
.Select-control{background:#0F0D33;border:1px solid #3A3478;color:#F5F1FF}
.Select-menu-outer{background:#15123D;border:1px solid #FF3C8E;color:#F5F1FF}
.Select-option{background:#15123D;color:#F5F1FF}
.Select-option.is-focused{background:#2A2570}
.Select-value{background:#2A2570!important;border:1px solid #FF3C8E!important}
.Select-value-label,.Select-placeholder,.Select-input>input{color:#F5F1FF!important}
.Select-arrow{border-top-color:#FF3C8E!important}
.tarjeta{transition:transform .2s ease}
.tarjeta:hover{transform:translateY(-2px)}
.estrella{color:#FFD23C;text-shadow:0 0 10px #FFD23C99;animation:pulso 2.4s ease-in-out infinite}
@keyframes pulso{50%{opacity:.65}}
.lista-neon li::marker{color:#FF9A3C}
.chip{display:inline-block;padding:5px 12px;border-radius:999px;background:rgba(11,10,34,.72);border:1px solid rgba(255,255,255,.28);font-size:13px;font-weight:600;margin:0 8px 8px 0}
@media (prefers-reduced-motion:reduce){.estrella{animation:none}.tarjeta{transition:none}}
"""


def _arte_svg() -> str:
    """Ilustración del encabezado: sol retro con franjas y palmeras en silueta."""
    rayas = "".join(f'<rect x="1020" y="{150 + i * 17}" width="220" height="{3 + i * 2}" fill="#7A1B6E"/>' for i in range(5))

    def palma(x: int, base: int, alto: int, inclin: int) -> str:
        tx, ty = x + inclin, base - alto
        frondas = ""
        for dx, dy in ((-70, 10), (-55, -22), (-25, -40), (25, -40), (55, -22), (70, 10), (-48, 38), (48, 38)):
            frondas += f'<path d="M{tx} {ty} Q{tx + dx * 0.6} {ty + dy - 26} {tx + dx} {ty + dy + 22}" />'
        return (f'<path d="M{x} {base} Q{x + inclin * 0.2} {base - alto * 0.5} {tx} {ty}" stroke-width="7"/>'
                f'<g stroke-width="5">{frondas}</g>')

    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1400 250" preserveAspectRatio="xMaxYMax slice">'
        '<defs><linearGradient id="sol" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#FFD23C"/>'
        '<stop offset=".5" stop-color="#FF8A3C"/><stop offset="1" stop-color="#FF3C8E"/></linearGradient>'
        '<clipPath id="c"><circle cx="1130" cy="140" r="108"/></clipPath></defs>'
        '<g clip-path="url(#c)"><rect x="1010" y="30" width="240" height="240" fill="url(#sol)"/>' + rayas + '</g>'
        '<g stroke="#12062B" fill="none" stroke-linecap="round">' + palma(1290, 250, 150, -18) + palma(1345, 250, 110, 12) + '</g>'
        '</svg>'
    )


ARTE = f'url("data:image/svg+xml;utf8,{quote(_arte_svg())}")'
GRADIENTE = "linear-gradient(180deg, #2A0B52 0%, #7A1B6E 58%, #D6306E 88%, #FF8A3C 100%)"
S_PESTANA = dict(background=PANEL, color=SUAVE, border=f"1px solid {BORDE}", borderRadius="10px 10px 0 0",
                 fontFamily=F_DISPLAY, letterSpacing="1.5px", textTransform="uppercase", padding="12px 18px")
S_PESTANA_SEL = {**S_PESTANA, "background": "linear-gradient(90deg, #FF3C8E, #FF9A3C)", "color": FONDO, "border": "1px solid #FF3C8E"}


def _chip(texto: str, color: str = "white") -> html.Span:
    return html.Span(texto, className="chip", style=dict(color=color))


def _filtro(id_: str, etiqueta: str, columna: str, opciones: list) -> html.Div:
    """Lista desplegable múltiple; vacía significa 'todos'."""
    return html.Div([
        html.Label(etiqueta, htmlFor=id_, style=dict(fontWeight=600, fontSize="14px", color=NEON["cian"],
                                                      textTransform="uppercase", letterSpacing="1px")),
        dcc.Dropdown(id=id_, options=opciones, multi=True, placeholder="Todos"),
    ], style=dict(flex="1 1 220px"))


def crear_app(df: pd.DataFrame) -> Dash:
    app = Dash(__name__, external_stylesheets=[FUENTES], title="GTA VI: dashboard para Rockstar")
    server = app.server
    app.index_string = (
        '<!DOCTYPE html><html lang="es"><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}'
        "<style>" + CSS + "</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>"
    )

    opciones_video = sorted(df["video"].unique())
    opciones_sent = sorted(df["sentimiento"].unique())
    opciones_tema = sorted(df["tema_principal"].unique())

    app.layout = html.Div([
        html.Header([
            html.Div(style=dict(position="absolute", top=0, left=0, right=0, bottom=0, backgroundImage=ARTE,
                                backgroundSize="cover", backgroundPosition="right bottom", backgroundRepeat="no-repeat")),
            html.Div([
                html.Div([
                    html.Span("GTA VI", style=dict(fontFamily=F_DISPLAY, fontSize="46px", letterSpacing="2px",
                                                   background="linear-gradient(90deg, #FF3C8E, #FFB347)",
                                                   WebkitBackgroundClip="text", WebkitTextFillColor="transparent")),
                    html.Span("DASHBOARD PARA ROCKSTAR", style=dict(marginLeft="14px", fontSize="13px", letterSpacing="3px",
                                                                     color="#E7D9FF", fontWeight=600)),
                ]),
                html.H1(TITULAR, style=dict(fontFamily=F_TITULO, fontSize="28px", lineHeight="1.25", margin="6px 0 16px",
                                            maxWidth="860px", color="white", textShadow="0 2px 14px rgba(0,0,0,.55)")),
                html.Div([_chip(cuenta_regresiva(), NEON["amarillo"]), _chip(f"{len(df)} comentarios · {df['video'].nunique()} videos"),
                          _chip("Etiquetas asistidas, sin validación humana")]),
            ], style=dict(position="relative")),
        ], style=dict(position="relative", overflow="hidden", background=GRADIENTE, color="white", padding="26px 32px 22px",
                      borderBottom=f"3px solid {ACENTO}", boxShadow=f"0 6px 34px {ACENTO}88")),
        html.Main([
            html.Div([
                html.Div("PANEL DE CONTROL", style=dict(fontFamily=F_DISPLAY, letterSpacing="2px", fontSize="13px",
                                                        color=ACENTO, marginBottom="8px")),
                html.Div([_filtro("f-video", "Video", "video", opciones_video),
                          _filtro("f-sent", "Sentimiento", "sentimiento", opciones_sent),
                          _filtro("f-tema", "Tema principal", "tema_principal", opciones_tema)], style={**FILA, "marginBottom": "4px"}),
                html.P("Los filtros aplican a las vistas 2 y 3.", style={**S_NOTA, "fontSize": "13px", "margin": 0}),
            ], style={**S_TARJETA, "marginBottom": "18px"}),
            dcc.Tabs(id="tabs", value="v1", colors=dict(border=BORDE, primary=ACENTO, background=PANEL), children=[
                dcc.Tab(label="Vista 1 · Resumen ejecutivo", value="v1", style=S_PESTANA, selected_style=S_PESTANA_SEL),
                dcc.Tab(label="Vista 2 · Temas y comparaciones", value="v2", style=S_PESTANA, selected_style=S_PESTANA_SEL),
                dcc.Tab(label="Vista 3 · Evidencia y calidad", value="v3", style=S_PESTANA, selected_style=S_PESTANA_SEL),
            ]),
            html.Div(id="contenido", style=dict(marginTop="16px")),
        ], style=dict(padding="20px 32px", maxWidth="1300px", margin="0 auto")),
        html.Footer(
            f"Alcance: {len(df)} comentarios públicos anonimizados de 3 videos. Sentimiento y tema son propuestas "
            "asistidas sin validación humana; no describen la experiencia de juego ni a toda la comunidad.",
            style=dict(padding="16px 32px", fontSize="13px", color=SUAVE, borderTop=f"1px solid {BORDE}")),
    ], style=dict(background=FONDO, fontFamily=F_TEXTO, color=TINTA, minHeight="100vh"))

    @app.callback(
        Output("contenido", "children"),
        Input("tabs", "value"), Input("f-video", "value"), Input("f-sent", "value"), Input("f-tema", "value"),
    )
    def render(vista: str, videos, sentimientos, temas):
        if vista == "v1":
            return vista_resumen(df)
        dff = filtrar(df, video=videos, sentimiento=sentimientos, tema_principal=temas)
        if dff.empty:
            return tarjeta("Ningún comentario cumple los filtros", html.P("Quita algún filtro para ver resultados.", style=S_NOTA))
        return vista_temas(dff) if vista == "v2" else vista_evidencia(dff, df)

    return app


# ---------------------------------------------------------------------------
# 7. Punto de Entrada Principal
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    df_datos = cargar_datos()
    print(f"Cargados {len(df_datos)} comentarios, {df_datos['video'].nunique()} videos, {df_datos['tema_principal'].nunique()} temas.")
    app = crear_app(df_datos)
    app.run(debug=True, port=8088)