# -*- coding: utf-8 -*-
"""
Dashboard GTA VI para Rockstar
Asignatura: ADY1104 Visualización de Datos
Pregunta central: ¿Qué tema debería priorizar Rockstar al comunicar GTA VI?

Uso:
1. Instala dependencias: pip install -r requirements.txt
2. Asegúrate de tener el archivo "GTA_VI_Datos_Visualizacion.xlsx" en la carpeta data/
3. Ejecuta localmente: python app.py
4. En Windows, inicia el servidor WSGI con: waitress-serve --listen=127.0.0.1:8088 app:server
5. En Render, usa el Blueprint definido en render.yaml
"""

from __future__ import annotations

import math
import unicodedata
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
BASE_DIR = Path(__file__).resolve().parent
CANDIDATAS = [
    BASE_DIR / "data" / NOMBRE_ARCHIVO,
    BASE_DIR / NOMBRE_ARCHIVO,
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


def _sin_tildes(texto) -> str:
    """Minúsculas y sin tildes, para que 'grafico' encuentre 'gráfico'."""
    return "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn").lower()


def buscar(df: pd.DataFrame, texto: str | None) -> pd.DataFrame:
    """Comentarios que contienen TODAS las palabras escritas (sin importar mayúsculas ni tildes)."""
    palabras = _sin_tildes(texto or "").split()
    if not palabras:
        return df
    base = df["comentario"].map(_sin_tildes)
    mask = pd.Series(True, index=df.index)
    for palabra in palabras:
        mask &= base.str.contains(palabra, regex=False)
    return df[mask]


def filtrar_v3(df: pd.DataFrame, texto: str | None, preguntas: str | None) -> pd.DataFrame:
    """Filtros propios de la vista 3: búsqueda de palabras y tipo de comentario (pregunta o no)."""
    df = buscar(df, texto)
    if preguntas == "si":
        df = df[df["es_pregunta"].eq("Sí")]
    elif preguntas == "no":
        df = df[df["es_pregunta"].eq("No")]
    return df


def resumen_filtros(videos, sentimientos, temas, texto, preguntas) -> str:
    """Texto con los filtros activos, p. ej. 'Video 1 · Negativo · palabra “fps”'. Vacío si no hay ninguno."""
    partes = []
    if videos:
        partes.append(", ".join(videos))
    if sentimientos:
        partes.append(", ".join(sentimientos))
    if temas:
        partes.append(", ".join(temas))
    if texto and texto.strip():
        partes.append(f"contiene “{texto.strip()}”")
    if preguntas == "si":
        partes.append("solo preguntas")
    elif preguntas == "no":
        partes.append("sin preguntas")
    return " · ".join(partes)


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
        font=dict(family=F_TEXTO, size=14, color=TINTA),
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


def fig_barras_h(etiquetas: list[str], valores, textos: list[str], resaltar: int | None = None) -> go.Figure:
    """Barras horizontales simples: una sola serie, la barra más importante resaltada en rosa."""
    valores = [float(v) for v in valores]
    colores = [NEON["rosa"] if i == resaltar else NEON["cian"] for i in range(len(valores))]
    fig = go.Figure(go.Bar(
        y=etiquetas, x=valores, orientation="h", marker_color=colores, text=textos,
        textposition="outside", cliponaxis=False, textfont=dict(size=14, color=TINTA),
        hovertemplate="%{y}: %{text}<extra></extra>",
    ))
    tope = max(valores) * 1.3 if valores else 1.0
    fig.update_xaxes(range=[0, tope if tope > 0 else 1.0], showticklabels=False, showgrid=False)
    fig.update_yaxes(tickfont=dict(size=14))
    return _estilo(fig, 70 + 40 * len(valores))


# ---------------------------------------------------------------------------
# 4. Componentes de Interfaz
# ---------------------------------------------------------------------------
S_TARJETA = dict(
    background="linear-gradient(160deg, rgba(36,31,98,.92), rgba(15,13,46,.94))",
    border=f"1px solid {BORDE}", borderRadius="14px", padding="18px 20px", minWidth="0",
)
S_TITULO = dict(fontFamily=F_TITULO, fontSize="18px", fontWeight=700, margin="0 0 6px", color=TINTA, lineHeight="1.25")
S_NOTA = dict(fontSize="15px", color=SUAVE, margin="0 0 8px", lineHeight="1.45")


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
COLUMNA = dict(display="flex", flexDirection="column", gap="16px")


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


def bloque(num: int, pregunta: str, hallazgo: str, *contenido, como: str | None = None, acento: str = ACENTO) -> html.Section:
    """Tarjeta con estructura fija: pregunta -> conclusión en una frase -> gráfico -> cómo leerlo."""
    hijos = [html.P(["→ ", hallazgo], style=dict(fontSize="16px", fontWeight=600, color=TINTA, margin="0 0 10px", lineHeight="1.4")),
             *contenido]
    if como:
        hijos.append(html.P([html.Strong("Cómo leerlo: "), como], style={**S_NOTA, "margin": "8px 0 0"}))
    return tarjeta(pregunta, *hijos, acento=acento, etiqueta=f"PREGUNTA {num}")


def leyenda() -> html.Div:
    """Guía rápida de colores y símbolos que se usan en toda la vista."""
    def punto(color: str) -> html.Span:
        return html.Span(style=dict(display="inline-block", width="12px", height="12px", borderRadius="50%",
                                    background=color, marginRight="6px", verticalAlign="middle"))
    chips = [html.Span([punto(COLORES[k]), k], className="chip", style=dict(color=TINTA)) for k in ORDEN]
    chips.append(html.Span(f"⚠ = menos de {MIN_N} comentarios: tómalo con cuidado", className="chip", style=dict(color=NEON["naranja"])))
    return html.Div([html.Span("GUÍA RÁPIDA", style=dict(fontFamily=F_DISPLAY, letterSpacing="2px", fontSize="13px",
                                                          color=ACENTO, marginRight="12px")), *chips],
                    style=dict(display="flex", flexWrap="wrap", alignItems="center"))


def tabla_tema_video(df: pd.DataFrame) -> html.Div:
    """Tabla tema x video con el % de comentarios negativos; más rosa = más críticas."""
    n = pd.crosstab(df["tema_principal"], df["video"])
    neg = (df[df["sentimiento"].eq("Negativo")].groupby(["tema_principal", "video"]).size()
           .unstack(fill_value=0).reindex(index=n.index, columns=n.columns, fill_value=0))
    p = neg / n * 100
    n_tot, p_tot = n.sum(axis=1), neg.sum(axis=1) / n.sum(axis=1) * 100
    orden = p_tot.sort_values(ascending=False).index
    visibles = p.where(n >= MIN_CELDA)
    zmax = max(25.0, float(visibles.max().max())) if visibles.notna().any().any() else 25.0

    def celda(k: int, pc: float) -> html.Td:
        base = dict(textAlign="center", padding="8px 10px", border=f"1px solid {BORDE}", minWidth="110px")
        if k >= MIN_CELDA:
            fondo = _mezcla("#1B1850", NEON["rosa"], min(1.0, pc / zmax))
            return html.Td([html.Div(f"{pc:.0f}%", style=dict(fontFamily=F_DISPLAY, fontSize="22px", letterSpacing="1px")),
                            html.Div(f"de {k} comentarios", style=dict(fontSize="12px"))],
                           style={**base, "background": fondo, "color": "white"})
        texto = "sin comentarios" if k == 0 else f"muy pocos (solo {k})"
        return html.Td(texto, style={**base, "background": PANEL, "color": "#8C90BD", "fontSize": "13px"})

    th = dict(padding="8px 10px", fontFamily=F_TITULO, fontSize="14px", color=NEON["cian"], textAlign="center",
              borderBottom=f"2px solid {BORDE}")
    cabecera = html.Tr([html.Th("Tema", style={**th, "textAlign": "left"}),
                        *[html.Th(v, style=th) for v in n.columns], html.Th("Todos los videos", style=th)])
    filas = [html.Tr([
        html.Td(f"{t}{' ⚠' if n_tot[t] < MIN_N else ''}", style=dict(padding="8px 10px", fontWeight=600, fontSize="15px")),
        *[celda(int(n.loc[t, v]), float(p.loc[t, v])) for v in n.columns],
        celda(int(n_tot[t]), float(p_tot[t])),
    ]) for t in orden]
    return html.Div(html.Table([html.Thead(cabecera), html.Tbody(filas)],
                               style=dict(width="100%", borderCollapse="separate", borderSpacing="4px")),
                    style=dict(overflowX="auto"))


def vista_temas(df: pd.DataFrame) -> html.Div:
    """Vista 2: cinco preguntas simples sobre temas y videos, cada una con su conclusión."""
    n_tot = len(df)
    cv, ct = composicion(df, "video"), contraste_por_tema(df)
    base = ct[ct["n"] >= MIN_N]

    # 1. ¿De qué habla la gente?
    conteo = df["tema_principal"].value_counts().sort_values()
    top = conteo.idxmax()
    h1 = f"{top} es lo más comentado: {int(conteo.max())} de {n_tot} comentarios ({pct(conteo.max() / n_tot * 100)})."
    concretos = conteo.drop("Otros temas", errors="ignore")
    if top == "Otros temas" and not concretos.empty:
        h1 += f" Entre los temas concretos, el más comentado es {concretos.idxmax()} ({int(concretos.max())})."
    c1 = "Cada barra es un tema; más larga significa más comentarios. “Otros temas” junta lo que no encajó en ninguna categoría."
    b1 = bloque(1, "¿De qué habla la gente?", h1, grafico(fig_barras_h(
        list(conteo.index), conteo.values, [f"{v}  ({v / n_tot * 100:.1f}%)" for v in conteo.values],
        resaltar=len(conteo) - 1)), como=c1, acento=NEON["cian"])

    # 2. ¿Qué temas generan más críticas y más elogios?
    if base.empty:
        h2 = f"Ningún tema llega a {MIN_N} comentarios, así que no se puede comparar con confianza."
    elif base["Negativo"].max() == 0:
        h2 = "Ningún tema con base suficiente tiene comentarios negativos."
    else:
        crit, elog = base["Negativo"].idxmax(), base["Positivo"].idxmax()
        if crit == elog:
            h2 = (f"{crit} es el tema que más divide: {pct(base.loc[crit, 'Negativo'])} de sus comentarios son negativos "
                  f"y {pct(base.loc[crit, 'Positivo'])} positivos.")
        else:
            h2 = (f"{crit} es el tema con más críticas ({pct(base.loc[crit, 'Negativo'])} negativos) y "
                  f"{elog} el que recibe más elogios ({pct(base.loc[elog, 'Positivo'])} positivos).")
    c2 = ("Cada fila es un tema. La barra rosa hacia la izquierda son los comentarios negativos y la verde hacia la derecha los "
          "positivos. Lo que falta para llegar a 100 % es neutral y no se dibuja.")
    b2 = bloque(2, "¿Qué temas generan más críticas y más elogios?", h2,
                grafico(fig_divergente(ct.sort_values("Negativo"))), como=c2, acento=NEON["rosa"])

    # 3. ¿Cambia según el video?
    neg = cv["Negativo"]
    h3 = (f"El sentimiento negativo va de {pct(neg.min())} a {pct(neg.max())} según el video."
          if len(cv) > 1 else "Estás viendo un solo video.")
    c3 = "Cada barra es un video y suma 100 %. Verde = positivo, gris = neutral, rosa = negativo. n = comentarios de ese video."
    b3 = bloque(3, "¿Cambia la opinión según el video?", h3, grafico(fig_composicion(cv)), como=c3, acento=NEON["violeta"])

    # 4. ¿Dónde se concentran las críticas?
    celda = celda_mas_critica(df)
    if celda is None:
        h4 = f"Ninguna combinación de tema y video llega a {MIN_CELDA} comentarios."
    elif celda[2] == 0:
        h4 = "Ninguna combinación con base suficiente tiene comentarios negativos."
    else:
        h4 = f"La combinación más crítica es {celda[0]} en {celda[1]}: {pct(celda[2])} de {celda[3]} comentarios son negativos."
    c4 = (f"Cada casilla muestra el % de comentarios negativos de ese tema en ese video: más rosa = más críticas. "
          f"Si hay menos de {MIN_CELDA} comentarios no se calcula el porcentaje, porque un solo comentario cambiaría demasiado el resultado.")
    b4 = bloque(4, "¿En qué tema y video se concentran las críticas?", h4, tabla_tema_video(df), como=c4, acento=NEON["naranja"])

    # 5. ¿Qué dudas hace la gente?
    pt = preguntas_por_tema(df)
    elegibles = pt[pt["n"] >= MIN_N]
    h5 = (f"{elegibles['pct'].idxmax()} es el tema con más preguntas ({pct(elegibles['pct'].max())} de sus comentarios)."
          if not elegibles.empty else "Ningún tema tiene base suficiente para comparar preguntas.")
    plat = "Plataformas y rendimiento"
    nota_plat = ""
    if plat in pt.index:
        p_plat, n_plat = float(pt.loc[plat, "pct"]), int(pt.loc[plat, "n"])
        p_gen = float(df["es_pregunta"].eq("Sí").mean() * 100)
        if p_plat < p_gen:
            nota_plat = (f" En {plat.lower()} son preguntas el {pct(p_plat)} (promedio {pct(p_gen)}): "
                         "la inquietud aparece más como opinión que como duda explícita.")
        elif p_plat > p_gen:
            nota_plat = f" En {plat.lower()} son preguntas el {pct(p_plat)} (promedio {pct(p_gen)}): hay más dudas explícitas que en el resto."
        if n_plat < MIN_N and nota_plat:
            nota_plat += f" ⚠ Base pequeña (n={n_plat})."
    c5 = "Muestra qué % de los comentarios de cada tema contiene una pregunta. Entre paréntesis, cuántos comentarios tiene el tema." + nota_plat
    b5 = bloque(5, "¿Sobre qué temas hace preguntas la gente?", h5, grafico(fig_barras_h(
        [f"{t} (n={int(v)}){' ⚠' if v < MIN_N else ''}" for t, v in pt["n"].items()],
        pt["pct"].values, [pct(v) for v in pt["pct"]],
        resaltar=(list(pt.index).index(elegibles["pct"].idxmax()) if not elegibles.empty else None))),
        como=c5, acento=NEON["amarillo"])

    return html.Div([leyenda(), b1, b2, b3, b4, b5], style=COLUMNA)


def vista_evidencia(df: pd.DataFrame, df_global: pd.DataFrame, filtros_activos: str = "") -> html.Div:
    """Vista 3: los comentarios reales y qué tan confiables son las etiquetas."""
    prioridad = {"Negativo": 0, "Positivo": 1, "Neutral": 2}
    dt = df.assign(_o=df["sentimiento"].map(prioridad)).sort_values("_o", kind="stable")
    cols = ["video", "tema_principal", "sentimiento", "es_pregunta", "comentario"]
    nombres = {"video": "Video", "tema_principal": "Tema", "sentimiento": "Sentimiento",
               "es_pregunta": "¿Pregunta?", "comentario": "Comentario"}
    tabla = dash_table.DataTable(
        data=dt[cols].to_dict("records"),
        columns=[{"name": nombres[c], "id": c} for c in cols],
        page_size=10, sort_action="native",
        style_table=dict(overflowX="auto"),
        style_header=dict(fontWeight=700, backgroundColor="#221E5E", color=TINTA, fontFamily=F_TITULO,
                          border=f"1px solid {BORDE}", fontSize="15px"),
        style_cell=dict(textAlign="left", whiteSpace="normal", height="auto", padding="10px",
                        fontFamily=F_TEXTO, fontSize="15px", lineHeight="1.4", backgroundColor=PANEL, color=TINTA,
                        border=f"1px solid {BORDE}"),
        style_cell_conditional=[
            {"if": {"column_id": "video"}, "width": "80px", "minWidth": "80px"},
            {"if": {"column_id": "tema_principal"}, "width": "170px", "minWidth": "150px"},
            {"if": {"column_id": "sentimiento"}, "width": "110px", "minWidth": "100px"},
            {"if": {"column_id": "es_pregunta"}, "width": "90px", "minWidth": "80px"},
            {"if": {"column_id": "comentario"}, "minWidth": "320px"},
        ],
        style_data_conditional=[
            {"if": {"row_index": "odd"}, "backgroundColor": "#1B1850"},
            {"if": {"filter_query": '{sentimiento} = "Positivo"', "column_id": "sentimiento"},
             "color": COLORES["Positivo"], "fontWeight": 700},
            {"if": {"filter_query": '{sentimiento} = "Negativo"', "column_id": "sentimiento"},
             "color": COLORES["Negativo"], "fontWeight": 700},
            {"if": {"filter_query": '{es_pregunta} = "Sí"', "column_id": "es_pregunta"},
             "color": NEON["amarillo"], "fontWeight": 700},
        ],
    )
    vc = df["sentimiento"].value_counts()
    n_preg = int(df["es_pregunta"].eq("Sí").sum())
    sin_validar = df["estado_revision"].str.contains("sin validación").mean() * 100 if "estado_revision" in df.columns else 100.0
    otros = df["tema_principal"].eq("Otros temas").mean() * 100

    indic = html.Div([
        indicador(f"{len(df)}", "comentarios en la selección", "según los filtros de arriba", 100, NEON["cian"]),
        indicador(pct(sin_validar), "sin revisión humana", "las etiquetas las puso un programa", sin_validar, NEON["naranja"]),
        indicador(f"{n_preg}", "son preguntas", f"{pct(n_preg / len(df) * 100)} de la selección", n_preg / len(df) * 100, NEON["amarillo"]),
        indicador(pct(otros), "en “Otros temas”", "no encajaron en ningún tema", otros, NEON["violeta"]),
    ], style=FILA)

    paso1 = bloque(1, "¿Qué dicen exactamente los comentarios?",
                   (f"Estás viendo {len(df)} comentarios: {int(vc.get('Negativo', 0))} negativos, "
                    f"{int(vc.get('Positivo', 0))} positivos y {int(vc.get('Neutral', 0))} neutrales."),
                   html.P([html.Strong("Filtros activos: "), filtros_activos or "ninguno (se ven todos los comentarios)"],
                          style={**S_NOTA, "margin": "0 0 10px"}),
                   tabla, acento=NEON["cian"],
                   como=("Los negativos y positivos aparecen primero. Haz clic en el título de una columna para ordenarla y usa las flechas "
                         "de abajo para cambiar de página. Para ver otros comentarios, usa el panel de control de arriba: además de video, "
                         "sentimiento y tema, puedes buscar una palabra y quedarte solo con las preguntas."))

    limites = [
        "Las etiquetas de sentimiento y de tema las puso un programa con reglas simples; ninguna persona las revisó.",
        "Son reacciones a tres videos, no la experiencia de jugar GTA VI.",
        "Es una muestra de comentarios: no representa a toda la comunidad ni demuestra causas.",
        "Un comentario puede hablar de varios temas, pero aquí se le asigna solo uno.",
    ]
    confianza = tarjeta(
        "¿Cuánto se puede confiar en estos datos?",
        html.Ul([html.Li(t) for t in limites], className="lista-neon",
                style=dict(margin=0, paddingLeft="20px", lineHeight="1.7", fontSize="15px")),
        acento=NEON["rosa"], etiqueta="PREGUNTA 2")
    revisar = tarjeta(
        "Un ejemplo de etiqueta dudosa",
        cita(df_global, ID_REVISAR, "Etiquetado Positivo", NEON["naranja"]),
        html.P("Este comentario pide no ilusionarse con los gráficos, pero el programa lo marcó como positivo. "
               "Por eso el sentimiento es una pista inicial, no un resultado definitivo.", style=S_NOTA),
        acento=NEON["naranja"], etiqueta="CONTRAEJEMPLO")

    return html.Div([
        indic, paso1,
        html.Div([html.Div(confianza, style=dict(flex="1 1 380px", minWidth="0")),
                  html.Div(revisar, style=dict(flex="1 1 380px", minWidth="0"))], style=FILA),
    ], style=COLUMNA)


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
.tarjeta:has(.js-plotly-plot):hover{transform:none}
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


S_ETIQUETA_FILTRO = dict(fontWeight=600, fontSize="14px", color=NEON["cian"], textTransform="uppercase", letterSpacing="1px",
                         display="block", marginBottom="4px")


def _filtro(id_: str, etiqueta: str, columna: str, opciones: list) -> html.Div:
    """Lista desplegable múltiple; vacía significa 'todos'."""
    return html.Div([
        html.Label(etiqueta, htmlFor=id_, style=S_ETIQUETA_FILTRO),
        dcc.Dropdown(id=id_, options=opciones, multi=True, placeholder="Todos"),
    ], style=dict(flex="1 1 220px"))


def crear_app(df: pd.DataFrame) -> Dash:
    app = Dash(__name__, external_stylesheets=[FUENTES], title="GTA VI: dashboard para Rockstar")
    server = app.server
    app.index_string = (
        '<!DOCTYPE html><html lang="es"><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}'
        "<style>" + CSS + "</style></head><body>{%app_entry%}<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>"
    )

    def opciones(columna: str) -> list[dict]:
        """Opciones del desplegable con la cantidad de comentarios de cada una: 'Video 1 (263)'."""
        conteo = df[columna].value_counts()
        return [{"label": f"{v} ({int(conteo[v])})", "value": v} for v in sorted(conteo.index)]

    opciones_video, opciones_sent, opciones_tema = opciones("video"), opciones("sentimiento"), opciones("tema_principal")

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
                html.Div([
                    html.Div([
                        html.Label("Buscar en los comentarios", htmlFor="f-texto", style=S_ETIQUETA_FILTRO),
                        dcc.Input(id="f-texto", type="text", debounce=True, value="", autoComplete="off",
                                  placeholder="Escribe una palabra (fps, PC, precio…) y presiona Enter",
                                  style=dict(width="100%", boxSizing="border-box", padding="9px 12px", background="#0F0D33",
                                             color=TINTA, border=f"1px solid {BORDE}", borderRadius="6px",
                                             fontFamily=F_TEXTO, fontSize="15px")),
                    ], style=dict(flex="2 1 320px")),
                    html.Div([
                        html.Label("Tipo de comentario", style=S_ETIQUETA_FILTRO),
                        dcc.RadioItems(id="f-preg", value="todos",
                                       options=[{"label": "Todos", "value": "todos"},
                                                {"label": "Solo preguntas", "value": "si"},
                                                {"label": "Sin preguntas", "value": "no"}],
                                       labelStyle=dict(display="inline-block", marginRight="16px", fontSize="15px"),
                                       inputStyle=dict(marginRight="6px")),
                    ], style=dict(flex="1 1 280px")),
                    html.Div(html.Button("Limpiar filtros", id="btn-limpiar", n_clicks=0,
                                         style=dict(padding="9px 16px", background="transparent", color=TINTA,
                                                    border=f"1px solid {ACENTO}", borderRadius="6px", cursor="pointer",
                                                    fontFamily=F_TEXTO, fontSize="15px", fontWeight=600)),
                             style=dict(flex="0 0 auto")),
                ], id="filtros-v3", style={**FILA, "marginBottom": "4px", "alignItems": "flex-end", "display": "none"}),
                html.P(id="nota-filtros", style={**S_NOTA, "fontSize": "13px", "margin": 0}),
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
        Output("filtros-v3", "style"), Output("nota-filtros", "children"),
        Input("tabs", "value"),
    )
    def mostrar_filtros_v3(vista: str):
        """La búsqueda y el tipo de comentario solo aparecen (y solo aplican) en la vista 3."""
        estilo = {**FILA, "marginBottom": "4px", "alignItems": "flex-end", "display": "flex" if vista == "v3" else "none"}
        if vista == "v3":
            nota = "Video, sentimiento y tema filtran las vistas 2 y 3. La búsqueda y el tipo de comentario filtran solo esta vista."
        else:
            nota = "Los filtros de video, sentimiento y tema aplican a las vistas 2 y 3."
        return estilo, nota

    @app.callback(
        Output("f-video", "value"), Output("f-sent", "value"), Output("f-tema", "value"),
        Output("f-texto", "value"), Output("f-preg", "value"),
        Input("btn-limpiar", "n_clicks"), prevent_initial_call=True,
    )
    def limpiar(_clics):
        return [], [], [], "", "todos"

    @app.callback(
        Output("contenido", "children"),
        Input("tabs", "value"), Input("f-video", "value"), Input("f-sent", "value"), Input("f-tema", "value"),
        Input("f-texto", "value"), Input("f-preg", "value"),
    )
    def render(vista: str, videos, sentimientos, temas, texto, preguntas):
        if vista == "v1":
            return vista_resumen(df)
        dff = filtrar(df, video=videos, sentimiento=sentimientos, tema_principal=temas)
        if vista == "v3":
            dff = filtrar_v3(dff, texto, preguntas)
        if dff.empty:
            return tarjeta("Ningún comentario cumple los filtros",
                           html.P("Quita algún filtro, revisa la palabra buscada o pulsa “Limpiar filtros”.", style=S_NOTA))
        if vista == "v3":
            return vista_evidencia(dff, df, resumen_filtros(videos, sentimientos, temas, texto, preguntas))
        contenido = vista_temas(dff)
        if vista == "v2" and sentimientos:
            aviso = html.P(
                "⚠ Hay un filtro de sentimiento activo: los porcentajes de sentimiento (temas, videos y tabla) "
                "se calculan solo sobre esa selección y no representan la conversación completa.",
                style={**S_NOTA, "color": NEON["naranja"], "margin": "0 0 12px"})
            return html.Div([aviso, contenido])
        return contenido

    return app


# ---------------------------------------------------------------------------
# 7. Punto de Entrada Principal
# ---------------------------------------------------------------------------
df_datos = cargar_datos()
print(f"Cargados {len(df_datos)} comentarios, {df_datos['video'].nunique()} videos, {df_datos['tema_principal'].nunique()} temas.")
app = crear_app(df_datos)
server = app.server

if __name__ == "__main__":
    app.run(debug=False, port=8088)