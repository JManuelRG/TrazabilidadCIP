# -*- coding: utf-8 -*-
"""
Reporte automatizado en PDF - Bitácora digital de tanques móviles (Power Apps)

Lee la exportación de la lista de SharePoint (.xlsx), limpia los datos, calcula
indicadores de uso y de calidad de datos, genera gráficos (calendarios con
calplot + gráficos de barras con matplotlib) y arma un PDF con reportlab.

Requisitos (una sola vez, en la consola de Spyder o en Anaconda Prompt):
    pip install pandas openpyxl matplotlib reportlab calplot

Uso en Spyder:
    1) Ajusta RUTA_EXCEL (ruta completa al archivo, o deja el nombre si el
       archivo está en el directorio de trabajo de Spyder).
    2) Ejecuta el script completo (F5). El PDF se guarda en CARPETA_SALIDA.
"""

import re
import tempfile
from datetime import datetime, time as dtime
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

# =============================================================================
# CONFIGURACIÓN 
# =============================================================================
RUTA_EXCEL = Path("Tanques CL.xlsx")
CARPETA_SALIDA = Path(".")

TITULO = "Bitácora digital de tanques móviles"
SUBTITULO = "Reporte de uso de la aplicación y calidad de los datos"

# Criterios de análisis 
UMBRAL_DUPLICADO_H = 1        # registros del mismo tanque separados por menos de X h = probable captura duplicada
UMBRAL_REPROCESO_H = 24       # registros del mismo tanque separados por menos de X h = intervalo corto a revisar
MAX_SESION_MIN = 120          # sesiones más largas que esto se consideran no confiables (sesión abierta)
UMBRAL_CIERRE_REPETIDO = 10   # una misma hora de cierre repetida >= X veces se considera posible edición masiva

# Unifica nombres escritos de forma distinta 
ALIAS_USUARIOS = {
    "Dalia Itzel Coss Lopez": "Coss Lopez, Dalia Itzel", # Ejemplo de Alias
}

# =============================================================================
# CONSTANTES AUXILIARES
# =============================================================================
COL_PRIM_HEX = "#1F4E79"
COL_PRIM = colors.HexColor(COL_PRIM_HEX)
COL_SEC = "#7FB3D5"
COLORES_TIPO = ["#1F4E79", "#7FB3D5", "#E59866", "#A9CCE3"]

MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7,
         "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12}
MESES_NOMBRE = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MESES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep",
             "Oct", "Nov", "Dic"]
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

COLUMNAS_REQUERIDAS = ["Fecha_Registro", "Hora_Registro", "Número_TM",
                       "Tipo_limpieza", "Registro_Usuario",
                       "Hora_Inicio de sesion", "Hora_Cierre de sesion"]


# =============================================================================
# 1) LECTURA Y LIMPIEZA DE DATABASE
# =============================================================================
def _parse_fecha(valor):
    """Convierte '04-ago-2026' / '17-Sep-2026' (o fechas reales) a Timestamp."""
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor).normalize()
    m = re.match(r"^\s*(\d{1,2})[-/ ]([A-Za-zñÑ]+)\.?[-/ ](\d{4})\s*$", str(valor))
    if not m:
        return pd.to_datetime(valor, errors="coerce", dayfirst=True)
    dia, mes, anio = m.groups()
    n = MESES.get(mes.lower()) or MESES.get(mes.lower()[:3])
    return pd.Timestamp(int(anio), n, int(dia)) if n else pd.NaT


def _parse_hora(valor):
    """Convierte '10:56:25 AM' (o a. m./p. m., o 24 h) a datetime.time."""
    if pd.isna(valor):
        return None
    if isinstance(valor, datetime):
        return valor.time()
    if isinstance(valor, dtime):
        return valor
    t = str(valor).strip().upper()
    for a, b in (("A. M.", "AM"), ("P. M.", "PM"), ("A.M.", "AM"), ("P.M.", "PM")):
        t = t.replace(a, b)
    t = re.sub(r"\s+", " ", t)
    for fmt in ("%I:%M:%S %p", "%I:%M %p", "%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(t, fmt).time()
        except ValueError:
            continue
    return None


def _combinar(fechas, horas):
    salida = []
    for f, h in zip(fechas, horas):
        t = _parse_hora(h)
        salida.append(datetime.combine(f.date(), t) if (pd.notna(f) and t) else pd.NaT)
    return pd.Series(pd.to_datetime(salida), index=fechas.index)


def _limpiar_usuario(nombre):
    n = re.sub(r"\s+", " ", str(nombre)).strip()
    n = re.sub(r"\s+,", ",", n)
    return ALIAS_USUARIOS.get(n, n)


def cargar_datos(ruta):
    df = pd.read_excel(ruta, sheet_name=0)
    df.columns = [re.sub(r"\s+", " ", str(c)).strip().replace(" _", "_")
                  for c in df.columns]
    faltan = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltan:
        raise KeyError(f"Faltan columnas en el Excel: {faltan}")
    if "ID" in df.columns:
        df = df.drop_duplicates(subset="ID")

    df["fecha"] = df["Fecha_Registro"].map(_parse_fecha)
    df["t_registro"] = _combinar(df["fecha"], df["Hora_Registro"])
    df["t_inicio"] = _combinar(df["fecha"], df["Hora_Inicio de sesion"])
    df["t_cierre"] = _combinar(df["fecha"], df["Hora_Cierre de sesion"])
    df["tm"] = pd.to_numeric(df["Número_TM"], errors="coerce")
    df["tipo"] = df["Tipo_limpieza"].astype(str).str.strip()
    df["usuario"] = df["Registro_Usuario"].map(_limpiar_usuario)
    df = df.dropna(subset=["fecha", "t_registro", "tm"]).copy()
    df["tm"] = df["tm"].astype(int)

    # --- banderas de calidad de la sesión ---
    df["dur_min"] = (df["t_cierre"] - df["t_inicio"]).dt.total_seconds() / 60
    rep = df["Hora_Cierre de sesion"].map(df["Hora_Cierre de sesion"].value_counts())
    df["q_sin_inicio"] = df["t_inicio"].isna()
    df["q_sin_cierre"] = df["t_cierre"].isna()
    df["q_cierre_masivo"] = rep.fillna(0) >= UMBRAL_CIERRE_REPETIDO
    df["q_sesion_larga"] = df["dur_min"] > MAX_SESION_MIN
    df["q_dur_negativa"] = df["dur_min"] < 0
    df["sesion_confiable"] = ~(df[["q_sin_inicio", "q_sin_cierre", "q_cierre_masivo",
                                   "q_sesion_larga", "q_dur_negativa"]].any(axis=1))
    return df.sort_values("t_registro").reset_index(drop=True)


# =============================================================================
# 2) INDICADORES
# =============================================================================
def calcular(df):
    m = {"n": len(df)}
    m["n_tm"] = df["tm"].nunique()
    m["f_ini"], m["f_fin"] = df["fecha"].min(), df["fecha"].max()
    m["dias_activos"] = df["fecha"].nunique()
    m["dias_periodo"] = (m["f_fin"] - m["f_ini"]).days + 1
    m["n_usuarios"] = df["usuario"].nunique()

    m["por_dia"] = df.groupby("fecha").size()
    m["por_mes"] = df.groupby([df["fecha"].dt.year, df["fecha"].dt.month]).size()
    m["tipos"] = df["tipo"].value_counts()
    m["por_tm"] = df.groupby("tm").size().sort_values(ascending=False)
    m["usuario_tipo"] = (df.groupby(["usuario", "tipo"]).size().unstack(fill_value=0)
                         .assign(total=lambda x: x.sum(axis=1))
                         .sort_values("total", ascending=False))
    m["por_diasem"] = df.groupby(df["fecha"].dt.weekday).size().reindex(range(7), fill_value=0)
    m["por_hora"] = df.groupby(df["t_registro"].dt.hour).size().reindex(range(24), fill_value=0)

    # --- intervalos entre registros consecutivos del mismo tanque ---
    d = df.sort_values(["tm", "t_registro"]).copy()
    d["h_prev"] = d.groupby("tm")["t_registro"].diff().dt.total_seconds() / 3600
    d["tipo_prev"] = d.groupby("tm")["tipo"].shift()
    con_prev = d.dropna(subset=["h_prev"])
    m["intervalos"] = con_prev
    m["dup"] = con_prev[con_prev["h_prev"] < UMBRAL_DUPLICADO_H]
    m["corto"] = con_prev[(con_prev["h_prev"] >= UMBRAL_DUPLICADO_H)
                          & (con_prev["h_prev"] < UMBRAL_REPROCESO_H)]
    m["corto_mismo_tipo"] = int((m["corto"]["tipo"] == m["corto"]["tipo_prev"]).sum())
    m["mediana_intervalo_dias"] = con_prev["h_prev"].median() / 24 if len(con_prev) else np.nan

    # --- calidad ---
    m["calidad"] = {
        "Sin hora de inicio de sesión": int(df["q_sin_inicio"].sum()),
        "Sin hora de cierre de sesión": int(df["q_sin_cierre"].sum()),
        "Cierre idéntico repetido (posible edición masiva)": int(df["q_cierre_masivo"].sum()),
        f"Sesión mayor a {MAX_SESION_MIN} min (posible sesión abierta)": int(df["q_sesion_larga"].sum()),
        "Duración negativa": int(df["q_dur_negativa"].sum()),
    }
    m["n_confiables"] = int(df["sesion_confiable"].sum())
    return m


def _fmt_f(f):
    return f"{f.day:02d}/{f.month:02d}/{f.year}"


def generar_hallazgos(m):
    h = []
    h.append(f"Se analizaron {m['n']:,} registros de {m['n_tm']} tanques móviles entre el "
             f"{_fmt_f(m['f_ini'])} y el {_fmt_f(m['f_fin'])} ({m['dias_activos']} días con captura).")
    tipos = "; ".join(f"{k}: {v} ({v / m['n']:.0%})" for k, v in m["tipos"].items())
    h.append(f"Tipo de limpieza registrada: {tipos}.")
    meses = "; ".join(f"{MESES_NOMBRE[mes - 1]} {anio}: {v}" for (anio, mes), v in m["por_mes"].items())
    dia_max = m["por_dia"].idxmax()
    h.append(f"Registros por mes: {meses}. El día de mayor captura fue el {_fmt_f(dia_max)} "
             f"con {m['por_dia'].max()} registros.")
    top5 = m["usuario_tipo"]["total"].head(5).sum() / m["n"]
    h.append(f"{m['n_usuarios']} usuarios registraron en la app; los 5 con mayor captura concentran "
             f"{top5:.0%} de los registros.")
    n_corto = len(m["dup"]) + len(m["corto"])
    h.append(f"{n_corto} registros ocurren a menos de {UMBRAL_REPROCESO_H} h del registro anterior del "
             f"mismo tanque: {len(m['dup'])} a menos de {UMBRAL_DUPLICADO_H} h (probables capturas "
             f"duplicadas) y {len(m['corto'])} entre {UMBRAL_DUPLICADO_H} y {UMBRAL_REPROCESO_H} h, de "
             f"los cuales {m['corto_mismo_tipo']} repiten el mismo tipo de limpieza. La mediana entre "
             f"registros consecutivos de un mismo tanque es de {m['mediana_intervalo_dias']:.1f} días.")
    h.append(f"{m['n_confiables']} de {m['n']} registros ({m['n_confiables'] / m['n']:.0%}) tienen datos de "
             f"sesión confiables. Las incidencias de calidad no son excluyentes entre sí (ver tabla).")
    h.append("La base no incluye consumo de agua, receta CIP ni producto previo, por lo que este reporte "
             "mide uso y frecuencia de registro, no ahorro de agua. Un intervalo corto es un indicio a "
             "validar con el personal de lavado, no una prueba de re-lavado evitable.")
    return h


# =============================================================================
# 3) GRÁFICOS
# =============================================================================
def _guardar(fig, ruta):
    fig.savefig(ruta, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return ruta


def _calendario_matplotlib(serie, titulo, ruta, cmap):
    """Alternativa si calplot no está instalado o falla: mapa de calor tipo calendario."""
    ini = serie.index.min() - pd.Timedelta(days=serie.index.min().weekday())
    fin = serie.index.max() + pd.Timedelta(days=6 - serie.index.max().weekday())
    dias = pd.date_range(ini, fin)
    n_sem = len(dias) // 7
    mat = np.full((7, n_sem), np.nan)
    for i, dia in enumerate(dias):
        if dia in serie.index:
            mat[dia.weekday(), i // 7] = serie[dia]
    fig, ax = plt.subplots(figsize=(max(7, n_sem * 0.6 + 2), 3.0))
    cm = plt.get_cmap(cmap).copy()
    cm.set_bad("whitesmoke")
    im = ax.imshow(np.ma.masked_invalid(mat), cmap=cm, aspect="equal", vmin=0)
    for (r, c), v in np.ndenumerate(mat):
        if not np.isnan(v):
            ax.text(c, r, f"{int(v)}", ha="center", va="center", fontsize=6)
    for j in range(n_sem):
        d0 = dias[j * 7]
        if j == 0 or d0.month != dias[(j - 1) * 7].month:
            ax.text(j, -0.9, MESES_ABR[d0.month - 1], fontsize=8, ha="left")
    ax.set_yticks(range(7))
    ax.set_yticklabels(DIAS, fontsize=7)
    ax.set_xticks([])
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_title(titulo, fontsize=10, pad=18)
    fig.colorbar(im, ax=ax, shrink=0.7, pad=0.01)
    return _guardar(fig, ruta)


def calendario(serie, titulo, ruta, cmap="YlGn"):
    """Calendario de actividad con calplot (con alternativa en matplotlib)."""
    serie = serie[serie > 0]  # los días sin actividad quedan en gris
    try:
        import calplot
        fig, _ = calplot.calplot(serie, cmap=cmap, suptitle=titulo, colorbar=True,
                                 how="sum", figsize=(11, 2.8), fillcolor="whitesmoke")
        return _guardar(fig, ruta)
    except Exception as e:  # calplot ausente o incompatible con la versión de pandas/matplotlib
        print(f"  [aviso] calplot no disponible ({type(e).__name__}: {e}). "
              f"Se usa el calendario alternativo en matplotlib.")
        return _calendario_matplotlib(serie, titulo, ruta, cmap)


def generar_graficos(df, m, carpeta):
    g = {}
    carpeta = Path(carpeta)

    # --- calendarios (calplot) ---
    g["cal_total"] = calendario(m["por_dia"], "Registros por día (todos los tipos)",
                                carpeta / "cal_total.png", "YlGn")
    for i, (tipo, cmap) in enumerate(zip(m["tipos"].index, ["Blues", "Oranges", "Purples", "Greys"])):
        serie = df[df["tipo"] == tipo].groupby("fecha").size()
        g[f"cal_tipo_{i}"] = calendario(serie, f"Registros por día: {tipo}",
                                        carpeta / f"cal_tipo_{i}.png", cmap)

    # --- por usuario (apilado por tipo) ---
    ut = m["usuario_tipo"].drop(columns="total").iloc[::-1]
    fig, ax = plt.subplots(figsize=(8, max(2.5, 0.4 * len(ut) + 1)))
    izq = np.zeros(len(ut))
    for i, tipo in enumerate(ut.columns):
        ax.barh(ut.index, ut[tipo], left=izq, label=tipo, color=COLORES_TIPO[i % 4])
        izq += ut[tipo].values
    for y, tot in enumerate(izq):
        ax.text(tot + 0.5, y, f"{int(tot)}", va="center", fontsize=8)
    ax.set_xlabel("Registros")
    ax.set_title("Registros por usuario y tipo de limpieza", fontsize=10)
    ax.legend(fontsize=8, loc="lower right")
    ax.tick_params(axis="y", labelsize=8)
    ax.spines[["top", "right"]].set_visible(False)
    g["usuario"] = _guardar(fig, carpeta / "usuario.png")

    # --- día de la semana y hora ---
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(10, 3.2))
    a1.bar(DIAS, m["por_diasem"].values, color=COL_PRIM_HEX)
    a1.set_title("Registros por día de la semana", fontsize=10)
    a2.bar(m["por_hora"].index, m["por_hora"].values, color=COL_PRIM_HEX)
    a2.set_title("Registros por hora del día", fontsize=10)
    a2.set_xlabel("Hora")
    a2.set_xticks(range(0, 24, 2))
    for a in (a1, a2):
        a.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    g["semana_hora"] = _guardar(fig, carpeta / "semana_hora.png")

    # --- tanques con más registros ---
    top = m["por_tm"].head(10).iloc[::-1]
    fig, ax = plt.subplots(figsize=(6, 3.3))
    ax.barh([f"TM {i}" for i in top.index], top.values, color=COL_PRIM_HEX)
    for y, v in enumerate(top.values):
        ax.text(v + 0.05, y, str(v), va="center", fontsize=8)
    ax.set_title("10 tanques con más registros", fontsize=10)
    ax.set_xlabel("Registros")
    ax.spines[["top", "right"]].set_visible(False)
    g["top_tm"] = _guardar(fig, carpeta / "top_tm.png")

    # --- distribución de intervalos entre registros del mismo tanque ---
    dias = (m["intervalos"]["h_prev"] / 24).clip(upper=30)
    fig, ax = plt.subplots(figsize=(8, 3.2))
    ax.hist(dias, bins=30, color=COL_PRIM_HEX, edgecolor="white")
    ax.axvline(UMBRAL_REPROCESO_H / 24, color="#C0392B", ls="--", lw=1)
    ax.text(UMBRAL_REPROCESO_H / 24 + 0.3, ax.get_ylim()[1] * 0.9,
            f"{UMBRAL_REPROCESO_H} h", color="#C0392B", fontsize=8)
    ax.set_xlabel("Días entre registros consecutivos del mismo tanque (acotado a 30)")
    ax.set_ylabel("Frecuencia")
    ax.set_title("Intervalos entre registros por tanque", fontsize=10)
    ax.spines[["top", "right"]].set_visible(False)
    g["intervalos"] = _guardar(fig, carpeta / "intervalos.png")
    return g


# =============================================================================
# 4) PDF (reportlab)
# =============================================================================
def _estilos():
    base = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle("t", parent=base["Title"], fontName="Helvetica-Bold",
                                 fontSize=20, textColor=COL_PRIM, alignment=0, spaceAfter=2),
        "sub": ParagraphStyle("s", parent=base["Normal"], fontSize=11, textColor=colors.grey),
        "h1": ParagraphStyle("h1", parent=base["Heading1"], fontSize=13, textColor=COL_PRIM,
                             spaceBefore=10, spaceAfter=6),
        "txt": ParagraphStyle("p", parent=base["Normal"], fontSize=9.5, leading=13),
        "bullet": ParagraphStyle("b", parent=base["Normal"], fontSize=9.5, leading=13,
                                 leftIndent=12, bulletIndent=0, spaceAfter=3),
        "pie": ParagraphStyle("f", parent=base["Normal"], fontSize=8, leading=10,
                              textColor=colors.grey),
    }


def _img(ruta, ancho, max_alto=None):
    """Imagen escalada a 'ancho'; si excede 'max_alto', se reduce manteniendo proporción."""
    w, h = ImageReader(str(ruta)).getSize()
    alto = ancho * h / w
    if max_alto and alto > max_alto:
        ancho, alto = ancho * max_alto / alto, max_alto
    return Image(str(ruta), width=ancho, height=alto)


def _tabla(filas, anchos, col_num_desde=1):
    t = Table(filas, colWidths=anchos, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), COL_PRIM),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F6FA")]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C8D3DE")),
        ("ALIGN", (col_num_desde, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def _pie_pagina(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(2 * cm, 1.1 * cm, TITULO)
    canvas.drawRightString(letter[0] - 2 * cm, 1.1 * cm, f"Página {doc.page}")
    canvas.restoreState()


def construir_pdf(df, m, g, ruta_pdf):
    st = _estilos()
    ancho = 17 * cm
    doc = SimpleDocTemplate(str(ruta_pdf), pagesize=letter, leftMargin=2 * cm,
                            rightMargin=2 * cm, topMargin=1.8 * cm, bottomMargin=1.8 * cm,
                            title=TITULO)
    s = []

    # ---- Encabezado y KPIs ----
    ahora = datetime.now()
    s += [Paragraph(TITULO, st["titulo"]), Paragraph(SUBTITULO, st["sub"]),
          Paragraph(f"Generado el {ahora.day:02d}/{ahora.month:02d}/{ahora.year} "
                    f"a las {ahora:%H:%M}", st["pie"]), Spacer(1, 10)]

    pm = m["por_tm"]
    celda = ParagraphStyle("k", parent=st["txt"], fontSize=8.5, leading=10.5)
    datos = [
        ("Registros totales", f"{m['n']:,}", "Periodo", f"{_fmt_f(m['f_ini'])} a {_fmt_f(m['f_fin'])}"),
        ("Tanques móviles distintos", f"{m['n_tm']}", "Días con captura", f"{m['dias_activos']} de {m['dias_periodo']}"),
        ("Usuarios que registraron", f"{m['n_usuarios']}", "Registros por tanque (media / mediana / máx.)",
         f"{pm.mean():.1f} / {pm.median():.0f} / {pm.max()}"),
        ("Mediana entre registros del mismo tanque", f"{m['mediana_intervalo_dias']:.1f} días",
         "Sesiones con datos confiables", f"{m['n_confiables'] / m['n']:.0%}"),
    ]
    kpis = [[Paragraph(f"<b>{a}</b>", celda), Paragraph(b, celda),
             Paragraph(f"<b>{c}</b>", celda), Paragraph(d, celda)] for a, b, c, d in datos]
    t = Table(kpis, colWidths=[4.6 * cm, 2.8 * cm, 5.4 * cm, 4.2 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F2F6FA")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#C8D3DE")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    s += [t, Paragraph("Hallazgos principales", st["h1"])]
    for texto in generar_hallazgos(m):
        s.append(Paragraph(texto, st["bullet"], bulletText="•"))

    # ---- Uso en el tiempo ----
    s += [PageBreak(), Paragraph("1. Uso de la aplicación en el tiempo", st["h1"]),
          Paragraph("Los calendarios muestran el número de registros por día; los días sin actividad "
                    "aparecen en gris.", st["txt"]), Spacer(1, 6)]
    for clave in [k for k in g if k.startswith("cal_")]:
        s += [_img(g[clave], ancho, max_alto=4.6 * cm), Spacer(1, 4)]
    s.append(_img(g["semana_hora"], ancho, max_alto=5.3 * cm))

    # ---- Usuarios ----
    s += [PageBreak(), Paragraph("2. Registros por usuario", st["h1"]), _img(g["usuario"], 14 * cm, max_alto=9 * cm),
          Spacer(1, 8)]
    tipos = [c for c in m["usuario_tipo"].columns if c != "total"]
    filas = [["Usuario"] + tipos + ["Total", "% del total"]]
    for usuario, r in m["usuario_tipo"].iterrows():
        filas.append([usuario] + [int(r[c]) for c in tipos] + [int(r["total"]), f"{r['total'] / m['n']:.1%}"])
    s.append(_tabla(filas, [5.8 * cm] + [3.7 * cm] * len(tipos) + [1.6 * cm, 2.1 * cm]))

    # ---- Tanques e intervalos ----
    s += [PageBreak(), Paragraph("3. Tanques móviles e intervalos entre registros", st["h1"]),
          _img(g["top_tm"], 11 * cm, max_alto=6.5 * cm), Spacer(1, 6), _img(g["intervalos"], 14 * cm, max_alto=6.5 * cm), Spacer(1, 8)]
    ip = m["intervalos"]
    filas = [["Intervalo respecto al registro anterior del mismo tanque", "Registros", "%"],
             [f"Menos de {UMBRAL_DUPLICADO_H} h (probable captura duplicada)", len(m["dup"]),
              f"{len(m['dup']) / len(ip):.1%}"],
             [f"Entre {UMBRAL_DUPLICADO_H} y {UMBRAL_REPROCESO_H} h (intervalo corto a revisar)",
              len(m["corto"]), f"{len(m['corto']) / len(ip):.1%}"],
             ["   de los cuales, mismo tipo de limpieza que el anterior", m["corto_mismo_tipo"], ""],
             [f"{UMBRAL_REPROCESO_H} h o más", len(ip) - len(m["dup"]) - len(m["corto"]),
              f"{(len(ip) - len(m['dup']) - len(m['corto'])) / len(ip):.1%}"]]
    s.append(_tabla(filas, [11.5 * cm, 2.8 * cm, 2.7 * cm]))
    s.append(Paragraph(f"Base: {len(ip)} registros que tienen un registro previo del mismo tanque.", st["pie"]))

    # ---- Calidad de datos ----
    filas = [["Indicador", "Registros", "% del total"]]
    for k, v in m["calidad"].items():
        filas.append([k, v, f"{v / m['n']:.1%}"])
    filas.append(["Registros con sesión confiable", m["n_confiables"], f"{m['n_confiables'] / m['n']:.1%}"])
    s.append(KeepTogether([Paragraph("4. Calidad de los datos de sesión", st["h1"]),
                           _tabla(filas, [11.5 * cm, 2.8 * cm, 2.7 * cm]),
                           Paragraph("Un registro puede caer en más de una incidencia.", st["pie"])]))

    # ---- Notas ----
    s += [Paragraph("5. Notas metodológicas", st["h1"])]
    notas = [
        "Fuente: exportación de la lista de SharePoint de la aplicación (una fila por ciclo registrado).",
        f"Intervalo duplicado: menos de {UMBRAL_DUPLICADO_H} h; intervalo corto: menos de "
        f"{UMBRAL_REPROCESO_H} h. Son criterios configurables en el script, no definiciones normativas.",
        f"Sesión no confiable: sin hora de inicio o cierre, duración negativa o mayor a {MAX_SESION_MIN} min, "
        f"o una misma hora de cierre repetida {UMBRAL_CIERRE_REPETIDO} o más veces.",
        "La hora de registro se toma tal como viene en la base; no se verificó contra el proceso físico.",
        "El reporte no incluye consumo de agua ni receta CIP porque esos datos no están en la base.",
    ]
    for n in notas:
        s.append(Paragraph(n, st["bullet"], bulletText="•"))

    doc.build(s, onFirstPage=_pie_pagina, onLaterPages=_pie_pagina)


# =============================================================================
# 5) PRINCIPAL
# =============================================================================
def main():
    print(f"Leyendo: {RUTA_EXCEL}")
    df = cargar_datos(RUTA_EXCEL)
    print(f"  {len(df)} registros válidos.")
    m = calcular(df)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    ruta_pdf = CARPETA_SALIDA / f"Reporte_tanques_moviles_{datetime.now():%Y%m%d_%H%M}.pdf"

    with tempfile.TemporaryDirectory() as tmp:
        print("Generando gráficos...")
        g = generar_graficos(df, m, tmp)
        print("Construyendo PDF...")
        construir_pdf(df, m, g, ruta_pdf)

    print(f"Listo: {ruta_pdf.resolve()}")
    return ruta_pdf


if __name__ == "__main__":
    main()
