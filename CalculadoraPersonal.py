"""
Calcula el minimo de personal necesario (check-in y gate) por dia y por
estacion, a partir de una malla de itinerario de vuelos en Excel.

Genera:
1. 'personal_requerido.xlsx'          -> detalle diario tabular por estacion
2. 'dashboard_dimensionamiento.html'  -> dashboard interactivo (un solo archivo,
                                          se abre con doble clic en cualquier navegador)

Este es el UNICO archivo fuente para este calculo. Si en el futuro hace
falta ajustar una regla (duracion de turno, horas de check-in, etc.),
se edita SOLO aqui, para evitar que existan dos versiones del script
que se desincronicen (eso fue lo que causo que un bug ya corregido
volviera a aparecer en una version anterior).

Instalar dependencias:
    pip install pandas openpyxl plotly numpy --break-system-packages
"""

from __future__ import annotations

from datetime import datetime, timedelta
import datetime as dt_module
import heapq
import json
import math
import os
import sys
import threading
import traceback
import webbrowser

import tkinter as tk
from tkinter import filedialog, messagebox

# Soporte opcional de "arrastrar el Excel sobre la ventana ya abierta".
# Si el paquete no esta disponible, el programa sigue funcionando
# normalmente con el boton "Seleccionar archivo...".
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_DISPONIBLE = True
except ImportError:
    _DND_DISPONIBLE = False

# pandas/numpy/plotly son librerias pesadas: importarlas de una vez al
# abrir el programa es lo que hacia que la ventana tardara varios
# segundos en aparecer. Se importan aqui como "None" y solo se cargan
# de verdad la primera vez que hacen falta (dentro de un hilo aparte,
# ver _cargar_dependencias_pesadas), asi la ventana aparece al instante.
pd = None
np = None
go = None
make_subplots = None


def _cargar_dependencias_pesadas():
    global pd, np, go, make_subplots
    if pd is not None:
        return
    import numpy as _np
    import pandas as _pd
    import plotly.graph_objects as _go
    from plotly.subplots import make_subplots as _make_subplots
    np, pd, go, make_subplots = _np, _pd, _go, _make_subplots

# ============================================================
# ---- Parametros configurables ----
# ============================================================
BUFFER_CAMBIO_SALA = timedelta(minutes=15)

# Duracion total del turno de trabajo (la hora de comida va INCLUIDA
# dentro de estas 8 horas, no se suma aparte).
DURACION_TURNO = timedelta(hours=8)

REGLAS = {
    "nacional": {
        "checkin": {"duracion": timedelta(hours=2), "personas": 2},
        "gate": {"duracion": timedelta(hours=1), "personas": 2},
    },
    "internacional": {
        "checkin": {"duracion": timedelta(hours=3), "personas": 3},
        "gate": {"duracion": timedelta(hours=1), "personas": 3},
    },
}

# Completa con los codigos IATA de las estaciones nacionales de tu red.
# Si AMBOS extremos del vuelo estan aqui -> nacional. Si no -> internacional.
ESTACIONES_NACIONALES = {
    "MTR", "BOG", "MDE", "CTG", "BAQ", "CLO", "SMR", "ADZ", "PEI", "BGA",
}

DIAS_SEMANA = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]

COL_CHECKIN = "personal_checkin (turnos 8h)"
COL_GATE = "personal_gate (turnos 8h)"
COL_TOTAL = "total_personal (turnos 8h)"


# ============================================================
# ---- Parsing de la malla ----
# ============================================================
def dia_que_opera(dow) -> int:
    """DOW viene como un solo digito 1=Lunes ... 7=Domingo -> indice 0=Lunes."""
    n = int(float(str(dow).strip()))
    if not 1 <= n <= 7:
        raise ValueError(f"Valor de DOW fuera de rango (1-7): {dow!r}")
    return n - 1


def parsear_hora(valor) -> tuple:
    """
    Extrae (hora, minuto) de forma robusta frente a distintos formatos
    en que Excel puede entregar una hora:
    - datetime.time / datetime.datetime -> se usa directo
    - texto "06:30" o "6:30"
    - texto/numero "0630", "630", 630, 630.0 (el ".0" se descarta bien,
      NO se concatena como digitos -- ese fue un bug de una version anterior)
    """
    if pd.isna(valor):
        return 0, 0

    if isinstance(valor, (dt_module.time, datetime)):
        return valor.hour, valor.minute

    s = str(valor).strip()

    if ":" in s:
        partes = s.split(":")
        return int(partes[0]), int(partes[1])

    # Convertir via float primero elimina un posible ".0" sin mezclar
    # el punto decimal con los digitos de la hora.
    try:
        s_num = str(int(float(s)))
    except ValueError:
        s_num = "".join(filter(str.isdigit, s)) or "0"

    s_num = s_num.zfill(4)
    return int(s_num[:2]), int(s_num[2:4])


def expandir_vuelos(df: pd.DataFrame) -> pd.DataFrame:
    filas = []
    for _, fila in df.iterrows():
        desde = pd.to_datetime(fila["EFFECTIVE FROM"]).date()
        hasta = pd.to_datetime(fila["EFFECTIVE TO"]).date()
        dia_ok = dia_que_opera(fila["DOW"])
        h, m = parsear_hora(fila["ETD"])

        tipo = (
            "nacional"
            if fila["DEPARTURESTATION"] in ESTACIONES_NACIONALES
            and fila["ARRIVALSTATION"] in ESTACIONES_NACIONALES
            else "internacional"
        )

        dia = desde
        while dia <= hasta:
            if dia.weekday() == dia_ok:
                filas.append({
                    "vuelo": fila["FLIGHTNUMBER"],
                    "tipo": tipo,
                    "fecha": dia,
                    "estacion": fila["DEPARTURESTATION"],
                    "hora_salida": datetime.combine(dia, datetime.min.time())
                    + timedelta(hours=h, minutes=m),
                })
            dia += timedelta(days=1)

    return pd.DataFrame(filas)


def generar_tareas(vuelos: pd.DataFrame, servicio: str):
    tareas = []
    for _, fila in vuelos.iterrows():
        regla = REGLAS[fila["tipo"]][servicio]
        fin = fila["hora_salida"]
        inicio = fin - regla["duracion"]
        for _ in range(regla["personas"]):
            tareas.append((inicio, fin))
    return tareas


def generar_tareas_detalle(vuelos: pd.DataFrame, servicio: str):
    """
    Igual que generar_tareas, pero conservando a que vuelo pertenece cada
    tarea. Es lo que permite luego decir "este turno cubre los vuelos
    AB123, AB456" en vez de solo mostrar una hora de entrada/salida.
    """
    tareas = []
    for _, fila in vuelos.iterrows():
        regla = REGLAS[fila["tipo"]][servicio]
        fin = fila["hora_salida"]
        inicio = fin - regla["duracion"]
        for _ in range(regla["personas"]):
            tareas.append({"inicio": inicio, "fin": fin, "vuelo": str(fila["vuelo"])})
    return tareas


def ventanas_servicio(vuelos: pd.DataFrame, servicio: str):
    """Una fila por vuelo (sin duplicar por personas) -- para dibujar el Gantt."""
    out = []
    for _, fila in vuelos.iterrows():
        regla = REGLAS[fila["tipo"]][servicio]
        fin = fila["hora_salida"]
        inicio = fin - regla["duracion"]
        out.append({
            "vuelo": str(fila["vuelo"]),
            "inicio": inicio.isoformat(),
            "fin": fin.isoformat(),
        })
    return out


def asignar_turnos(tareas_detalle, duracion_turno: timedelta = DURACION_TURNO):
    """
    Asigna cada tarea (una persona-vuelo) a un turno especifico, en vez de
    solo contar cuantos turnos hacen falta. Un turno dura duracion_turno
    desde que se crea, y solo puede tener una tarea a la vez (una persona
    no puede atender 2 vuelos al mismo tiempo). Se reutiliza el primer
    turno ya creado que este libre y en el que la tarea quepa antes de
    que termine su jornada; si ninguno sirve, se crea uno nuevo.

    Esto da el mismo numero de turnos que el metodo por demanda agregada,
    pero ademas deja registrado que vuelos cubre cada turno (trazabilidad).
    """
    if not tareas_detalle:
        return []

    ordenadas = sorted(tareas_detalle, key=lambda t: t["inicio"])
    turnos = []

    for tarea in ordenadas:
        elegido = None
        for turno in turnos:
            if (turno["libre_desde"] + BUFFER_CAMBIO_SALA <= tarea["inicio"]
                    and tarea["fin"] <= turno["salida"]):
                elegido = turno
                break
        if elegido is None:
            entrada = tarea["inicio"]
            elegido = {
                "entrada": entrada,
                "salida": entrada + duracion_turno,
                "libre_desde": entrada,
                "tareas": [],
            }
            turnos.append(elegido)
        elegido["tareas"].append(tarea)
        elegido["libre_desde"] = tarea["fin"]

    return sorted(turnos, key=lambda t: t["entrada"])


# ============================================================
# ---- Calculo de headcount ----
# ============================================================
def personal_pico(tareas):
    """Maximo de personas requeridas al mismo tiempo (pico instantaneo)."""
    if not tareas:
        return 0
    tareas = sorted(tareas, key=lambda t: t[0])
    libres = []
    for inicio, fin in tareas:
        if libres and libres[0] <= inicio:
            heapq.heappop(libres)
        heapq.heappush(libres, fin + BUFFER_CAMBIO_SALA)
    return len(libres)


def personal_por_turnos(tareas, duracion_turno: timedelta = DURACION_TURNO, detalle: bool = False):
    """
    Numero minimo de personas (turnos de duracion fija, hora de entrada
    libre/escalonable) necesarias para cubrir todas las tareas del dia,
    incluyendo relevos cuando la demanda se sostiene mas tiempo del que
    dura un turno.
    """
    if not tareas:
        return (0, []) if detalle else 0

    delta_demanda = {}
    for inicio, fin in tareas:
        fin_real_ocupacion = fin + BUFFER_CAMBIO_SALA
        delta_demanda[inicio] = delta_demanda.get(inicio, 0) + 1
        delta_demanda[fin_real_ocupacion] = delta_demanda.get(fin_real_ocupacion, 0) - 1

    por_revisar = list(delta_demanda.keys())
    heapq.heapify(por_revisar)

    vencimientos_turnos = []
    demanda_actual = 0
    turnos_contratados = 0
    turnos_lista = []
    ya_revisado = set()

    while por_revisar:
        t = heapq.heappop(por_revisar)
        if t in ya_revisado:
            continue
        ya_revisado.add(t)

        if t in delta_demanda:
            demanda_actual += delta_demanda[t]

        while vencimientos_turnos and vencimientos_turnos[0] <= t:
            heapq.heappop(vencimientos_turnos)

        activos = len(vencimientos_turnos)
        if activos < demanda_actual:
            faltan = demanda_actual - activos
            for _ in range(faltan):
                vence = t + duracion_turno
                heapq.heappush(vencimientos_turnos, vence)
                heapq.heappush(por_revisar, vence)
                turnos_lista.append((t, vence))
            turnos_contratados += faltan

    if detalle:
        return turnos_contratados, sorted(turnos_lista)
    return turnos_contratados


def calcular_por_dia_y_estacion(malla: pd.DataFrame):
    vuelos = expandir_vuelos(malla)

    resultados = []
    for (fecha, estacion), grupo in vuelos.groupby(["fecha", "estacion"]):
        tareas_checkin = generar_tareas(grupo, "checkin")
        tareas_gate = generar_tareas(grupo, "gate")

        pico_checkin = personal_pico(tareas_checkin)
        pico_gate = personal_pico(tareas_gate)
        turnos_checkin = personal_por_turnos(tareas_checkin)
        turnos_gate = personal_por_turnos(tareas_gate)

        resultados.append({
            "fecha": fecha,
            "estacion": estacion,
            "vuelos": len(grupo),
            "pico_checkin": pico_checkin,
            "pico_gate": pico_gate,
            COL_CHECKIN: turnos_checkin,
            COL_GATE: turnos_gate,
            COL_TOTAL: turnos_checkin + turnos_gate,
        })

    df_resultados = pd.DataFrame(resultados).sort_values(["fecha", "estacion"])
    return df_resultados, vuelos


# ============================================================
# ---- Heatmap semanal (con fix de medianoche) ----
# ============================================================
def matriz_calor_semanal(vuelos_estacion: pd.DataFrame, servicio: str):
    """
    Matriz 7x24 con la demanda PROMEDIO simultanea por dia de la semana / hora.
    """
    acum = np.zeros((7, 24))
    conteo_dias = np.zeros(7)

    fechas_unicas = pd.Series(vuelos_estacion["fecha"].unique())
    for f in fechas_unicas:
        conteo_dias[f.weekday()] += 1

    for fecha, grupo in vuelos_estacion.groupby("fecha"):
        dow_nominal = fecha.weekday()
        for ini, fin in generar_tareas(grupo, servicio):
            cur = ini
            while cur < fin:
                offset_dias = (cur.date() - fecha).days
                dow_real = (dow_nominal + offset_dias) % 7
                acum[dow_real, cur.hour] += 1
                cur = (cur.replace(minute=0, second=0, microsecond=0)
                       + timedelta(hours=1))

    with np.errstate(divide="ignore", invalid="ignore"):
        prom = np.where(conteo_dias[:, None] > 0, acum / conteo_dias[:, None], 0.0).round(1)
    return prom


def turnos_para_fecha(vuelos_estacion: pd.DataFrame, fecha):
    """
    Para UNA estacion (ya filtrada) y UNA fecha puntual, calcula los
    turnos de check-in y de gate necesarios ese dia, dejando registrado
    que vuelos cubre cada turno (trazabilidad) y los datos crudos que
    necesita el Gantt del dia (ventanas de cada vuelo + bloques de turno).
    """
    grupo = vuelos_estacion[vuelos_estacion["fecha"] == fecha]

    tareas_checkin = generar_tareas_detalle(grupo, "checkin")
    tareas_gate = generar_tareas_detalle(grupo, "gate")
    turnos_checkin = asignar_turnos(tareas_checkin)
    turnos_gate = asignar_turnos(tareas_gate)

    def utilizacion(turno):
        cubierto = sum((t["fin"] - t["inicio"]).total_seconds() for t in turno["tareas"])
        return round(cubierto / DURACION_TURNO.total_seconds() * 100)

    def fmt_html(turnos):
        if not turnos:
            return "<div class='empty-state'>Sin turnos requeridos para este servicio.</div>"
        filas = []
        for i, turno in enumerate(turnos, start=1):
            offset_ini = "<span class='tag-offset'>día ant.</span>" if turno["entrada"].date() < fecha else ""
            offset_fin = "<span class='tag-offset'>día sig.</span>" if turno["salida"].date() > fecha else ""
            etiqueta_ini = f"<span class='hora-mono'>{turno['entrada'].strftime('%H:%M')}</span> {offset_ini}"
            etiqueta_fin = f"<span class='hora-mono'>{turno['salida'].strftime('%H:%M')}</span> {offset_fin}"

            vuelos_lista = sorted({t["vuelo"] for t in turno["tareas"]})
            vuelos_chips = "".join(f"<span class='chip-vuelo'>{v}</span>" for v in vuelos_lista)

            pct = utilizacion(turno)
            clase_pct = "utilizacion-baja" if pct < 30 else ("utilizacion-media" if pct < 60 else "utilizacion-alta")
            filas.append(
                f"<tr>"
                f"<td class='col-turno'><span class='turno-badge'>T{i}</span></td>"
                f"<td>{etiqueta_ini}</td>"
                f"<td>{etiqueta_fin}</td>"
                f"<td class='vuelos-cubiertos'>{vuelos_chips}</td>"
                f"<td>"
                f"  <div class='util-wrapper {clase_pct}'>"
                f"    <div class='util-bar'><div class='util-fill' style='width:{min(pct, 100)}%'></div></div>"
                f"    <span class='util-text'>{pct}%</span>"
                f"  </div>"
                f"</td>"
                f"</tr>"
            )
        return (
            "<div class='tabla-turnos-wrap'><table class='tabla-turnos'>"
            "<thead><tr><th>Turno</th><th>Entrada</th><th>Salida</th>"
            "<th>Vuelos Asignados</th><th>Ocupación</th></tr></thead><tbody>"
            + "".join(filas) + "</tbody></table></div>"
        )

    def fmt_raw(turnos):
        return [
            {
                "entrada": t["entrada"].isoformat(),
                "salida": t["salida"].isoformat(),
                "vuelos": sorted({x["vuelo"] for x in t["tareas"]}),
                "utilizacion": utilizacion(t),
            }
            for t in turnos
        ]

    return {
        "fecha": fecha.strftime("%d-%b-%Y"),
        "fecha_iso": fecha.isoformat(),
        "dow": DIAS_SEMANA[fecha.weekday()],
        "vuelos": len(grupo),
        "checkin": len(turnos_checkin),
        "gate": len(turnos_gate),
        "total": len(turnos_checkin) + len(turnos_gate),
        "turnos_checkin_html": fmt_html(turnos_checkin),
        "turnos_gate_html": fmt_html(turnos_gate),
        "ventanas_checkin": ventanas_servicio(grupo, "checkin"),
        "ventanas_gate": ventanas_servicio(grupo, "gate"),
        "turnos_checkin_raw": fmt_raw(turnos_checkin),
        "turnos_gate_raw": fmt_raw(turnos_gate),
    }


# ============================================================
# ---- Dashboard HTML interactivo ----
# ============================================================
def generar_dashboard_html(df_resumen: pd.DataFrame, vuelos: pd.DataFrame,
                            nombre_html: str = "dashboard_dimensionamiento.html"):
    print(f"\nGenerando dashboard interactivo: {nombre_html}...")

    estaciones = sorted(vuelos["estacion"].unique())

    # ---- KPIs de red (para las tarjetas del encabezado) ----
    red_por_dia = df_resumen.groupby("fecha")[COL_TOTAL].sum()
    kpi_headcount_prom = math.ceil(red_por_dia.mean())
    kpi_headcount_pico = int(red_por_dia.max())
    kpi_num_estaciones = len(estaciones)
    kpi_fecha_min = df_resumen["fecha"].min()
    kpi_fecha_max = df_resumen["fecha"].max()

    # ---- Tabla resumen por estacion (todas a la vista, sin usar el selector) ----
    resumen_por_estacion = (
        df_resumen.groupby("estacion")
        .agg(
            vuelos_prom=("vuelos", "mean"),
            checkin_prom=(COL_CHECKIN, "mean"),
            gate_prom=(COL_GATE, "mean"),
            total_prom=(COL_TOTAL, "mean"),
            total_pico=(COL_TOTAL, "max"),
        )
        .reset_index()
        .sort_values("estacion")
    )
    filas_tabla = "\n".join(
        f"<tr onclick=\"seleccionarEstacionDesdeTabla('{r.estacion}')\" style='cursor:pointer' title='Clic para inspeccionar {r.estacion}'>"
        f"<td><span class='iata-pill'>{r.estacion}</span></td>"
        f"<td>{r.vuelos_prom:.1f}</td>"
        f"<td>{math.ceil(r.checkin_prom)}</td>"
        f"<td>{math.ceil(r.gate_prom)}</td>"
        f"<td><span class='badge-total'>{math.ceil(r.checkin_prom) + math.ceil(r.gate_prom)}</span></td>"
        f"<td><span class='badge-pico'>{int(r.total_pico)}</span></td></tr>"
        for r in resumen_por_estacion.itertuples()
    )

    # ---- Escala de color COMUN para poder comparar estaciones entre si ----
    matrices = {}
    for est in estaciones:
        v_est = vuelos[vuelos["estacion"] == est]
        matrices[est] = {
            "checkin": matriz_calor_semanal(v_est, "checkin"),
            "gate": matriz_calor_semanal(v_est, "gate"),
        }
    max_checkin = max(m["checkin"].max() for m in matrices.values()) or 1
    max_gate = max(m["gate"].max() for m in matrices.values()) or 1

    horas_cols = [f"{h:02d}:00" for h in range(24)]

    fig = make_subplots(
        rows=3, cols=1,
        specs=[[{"type": "heatmap"}], [{"type": "heatmap"}], [{"type": "xy"}]],
        subplot_titles=(
            "Demanda promedio simultánea — CHECK-IN (personas por hora, promedio semanal)",
            "Demanda promedio simultánea — GATE / SALAS (personas por hora, promedio semanal)",
            "Headcount total requerido por día (Check-in + Gate, turnos de 8h)",
        ),
        vertical_spacing=0.11,
        row_heights=[0.33, 0.33, 0.34],
    )

    for i, est in enumerate(estaciones):
        es_primera = i == 0
        m = matrices[est]

        fig.add_trace(
            go.Heatmap(
                z=m["checkin"], x=horas_cols, y=DIAS_SEMANA,
                zmin=0, zmax=max_checkin, colorscale="Blues",
                text=m["checkin"], texttemplate="%{text:.1f}", textfont={"size": 10},
                xgap=2, ygap=2,
                showscale=True, colorbar=dict(title="Pers.", x=1.01, y=0.86, len=0.28, thickness=12),
                hovertemplate=f"<b>{est} · Check-in</b><br>%{{y}} — %{{x}}<br>Demanda prom: <b>%{{z:.1f}} personas</b><extra></extra>",
                visible=es_primera,
            ),
            row=1, col=1,
        )
        fig.add_trace(
            go.Heatmap(
                z=m["gate"], x=horas_cols, y=DIAS_SEMANA,
                zmin=0, zmax=max_gate, colorscale="YlOrBr",
                text=m["gate"], texttemplate="%{text:.1f}", textfont={"size": 10},
                xgap=2, ygap=2,
                showscale=True, colorbar=dict(title="Pers.", x=1.01, y=0.50, len=0.28, thickness=12),
                hovertemplate=f"<b>{est} · Gate</b><br>%{{y}} — %{{x}}<br>Demanda prom: <b>%{{z:.1f}} personas</b><extra></extra>",
                visible=es_primera,
            ),
            row=2, col=1,
        )

        serie = df_resumen[df_resumen["estacion"] == est].sort_values("fecha")
        fig.add_trace(
            go.Scatter(
                x=serie["fecha"], y=serie[COL_TOTAL],
                mode="lines+markers",
                line=dict(color="#2563eb", width=2.5, shape="linear"),
                fill="tozeroy",
                fillcolor="rgba(37, 99, 235, 0.08)",
                marker=dict(size=6, color="#1d4ed8", line=dict(width=1.5, color="#ffffff")),
                name=est,
                hovertemplate=f"<b>{est}</b><br>Fecha: %{{x|%d-%b-%Y}}<br>Headcount total: <b>%{{y}} turnos</b><extra></extra>",
                visible=es_primera, showlegend=False,
            ),
            row=3, col=1,
        )

    fig.update_layout(
        height=980,
        margin=dict(l=70, r=50, t=55, b=50),
        font=dict(family="Inter, Segoe UI, system-ui, sans-serif", size=12, color="#334155"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        hoverlabel=dict(bgcolor="#0f172a", font_size=12, font_family="Inter, Segoe UI, sans-serif"),
    )
    fig.update_xaxes(tickangle=0, row=1, col=1, showgrid=False)
    fig.update_xaxes(tickangle=0, row=2, col=1, showgrid=False)
    fig.update_xaxes(showgrid=True, gridcolor="rgba(148,163,184,0.18)", row=3, col=1)
    fig.update_yaxes(autorange="reversed", row=1, col=1, showgrid=False)
    fig.update_yaxes(autorange="reversed", row=2, col=1, showgrid=False)
    fig.update_yaxes(title_text="Turnos (8h)", showgrid=True, gridcolor="rgba(148,163,184,0.18)", row=3, col=1)

    grafico_html = fig.to_html(
        full_html=False, include_plotlyjs="cdn", config={"displaylogo": False, "responsive": True}, div_id="mainPlot"
    )

    # ---- Datos por estacion para el selector (KPIs, tabla de turnos) ----
    trazas_por_estacion = 3
    datos_estaciones = {}
    for i, est in enumerate(estaciones):
        res_est = df_resumen[df_resumen["estacion"] == est]
        visibilidad = [False] * (len(estaciones) * trazas_por_estacion)
        visibilidad[i * trazas_por_estacion] = True
        visibilidad[i * trazas_por_estacion + 1] = True
        visibilidad[i * trazas_por_estacion + 2] = True

        fila_critica = res_est.loc[res_est[COL_TOTAL].idxmax()]
        fecha_critica_iso = fila_critica["fecha"].isoformat()

        v_est = vuelos[vuelos["estacion"] == est]
        fechas_est = sorted(v_est["fecha"].unique())
        diario = {}
        for f in fechas_est:
            info = turnos_para_fecha(v_est, f)
            diario[info["fecha_iso"]] = info

        chk_ceil = math.ceil(res_est[COL_CHECKIN].mean())
        gate_ceil = math.ceil(res_est[COL_GATE].mean())

        datos_estaciones[est] = {
            "visible": visibilidad,
            "vuelos_prom": f"{res_est['vuelos'].mean():.1f}",
            "checkin_prom": chk_ceil,
            "gate_prom": gate_ceil,
            "total_prom": chk_ceil + gate_ceil,
            "total_pico": int(res_est[COL_TOTAL].max()),
            "fecha_min": fechas_est[0].isoformat(),
            "fecha_max": fechas_est[-1].isoformat(),
            "fecha_critica": fecha_critica_iso,
            "diario": diario,
        }

    datos_json = json.dumps(datos_estaciones, ensure_ascii=False)
    opciones_select = "\n".join(f'<option value="{est}">Estación {est}</option>' for est in estaciones)
    est_ini = estaciones[0]
    d0 = datos_estaciones[est_ini]

    # ---- Pagina HTML final ----
    html = f"""<!DOCTYPE html>
<html lang="es" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dimensionamiento Operativo de Personal — Aeropuertos</title>
<style>
  :root {{
    --bg: #f1f5f9;
    --surface: #ffffff;
    --surface-2: #f8fafc;
    --surface-hover: #f1f5f9;
    --ink: #0f172a;
    --ink-soft: #334155;
    --muted: #64748b;
    --line: #e2e8f0;
    --navy: #0b192c;
    --blue: #2563eb;
    --blue-soft: rgba(37, 99, 235, 0.1);
    --cyan: #0d9488;
    --orange: #d97706;
    --orange-soft: rgba(217, 119, 6, 0.12);
    --danger: #dc2626;
    --danger-soft: rgba(220, 38, 38, 0.1);
    --success: #059669;
    --shadow-sm: 0 1px 3px rgba(15, 23, 42, 0.06);
    --shadow: 0 10px 30px -5px rgba(15, 23, 42, 0.07);
    --radius: 16px;
  }}

  [data-theme="dark"] {{
    --bg: #090d16;
    --surface: #111827;
    --surface-2: #1e293b;
    --surface-hover: #263349;
    --ink: #f8fafc;
    --ink-soft: #cbd5e1;
    --muted: #94a3b8;
    --line: #1f2937;
    --navy: #060b13;
    --blue: #3b82f6;
    --blue-soft: rgba(59, 130, 246, 0.18);
    --cyan: #14b8a6;
    --orange: #f59e0b;
    --orange-soft: rgba(245, 158, 11, 0.18);
    --danger: #f87171;
    --danger-soft: rgba(248, 113, 113, 0.15);
    --success: #34d399;
    --shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.3);
    --shadow: 0 12px 32px -4px rgba(0, 0, 0, 0.45);
  }}

  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    font-family: 'Inter', ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: var(--bg);
    color: var(--ink);
    line-height: 1.5;
    transition: background 0.2s ease, color 0.2s ease;
  }}

  header {{
    background: linear-gradient(135deg, #091526 0%, #102542 60%, #17365d 100%);
    color: #fff;
    padding: 28px clamp(20px, 4vw, 64px) 92px;
    position: relative;
    overflow: hidden;
    border-bottom: 1px solid rgba(255,255,255,0.08);
  }}
  header::before {{
    content: '';
    position: absolute;
    inset: 0;
    background:
      radial-gradient(circle at 85% 20%, rgba(56, 189, 248, 0.22), transparent 35%),
      radial-gradient(circle at 15% 80%, rgba(37, 99, 235, 0.25), transparent 40%);
    pointer-events: none;
  }}
  .header-inner {{
    max-width: 1380px;
    margin: 0 auto;
    position: relative;
    z-index: 1;
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 20px;
    flex-wrap: wrap;
  }}
  .brand-badge {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: rgba(255, 255, 255, 0.1);
    backdrop-filter: blur(8px);
    border: 1px solid rgba(255, 255, 255, 0.15);
    padding: 5px 12px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: #93c5fd;
  }}
  header h1 {{
    margin: 12px 0 6px;
    font-size: clamp(24px, 2.8vw, 36px);
    font-weight: 800;
    letter-spacing: -0.03em;
    line-height: 1.15;
  }}
  header p {{
    margin: 0;
    color: #cbd5e1;
    font-size: 13.5px;
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
  }}
  .header-actions {{
    display: flex;
    align-items: center;
    gap: 10px;
  }}
  .btn-theme {{
    background: rgba(255, 255, 255, 0.1);
    color: #fff;
    border: 1px solid rgba(255, 255, 255, 0.18);
    border-radius: 10px;
    padding: 9px 14px;
    font: inherit;
    font-size: 12.5px;
    font-weight: 600;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 7px;
    transition: 0.2s;
  }}
  .btn-theme:hover {{ background: rgba(255, 255, 255, 0.2); }}

  .kpis {{
    max-width: 1380px;
    margin: -58px auto 24px;
    padding: 0 24px;
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 18px;
    position: relative;
    z-index: 2;
  }}
  .kpi-card {{
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: var(--radius);
    padding: 20px 24px;
    box-shadow: var(--shadow);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
  }}
  .kpi-card .valor {{
    font-size: 32px;
    line-height: 1.05;
    font-weight: 800;
    color: var(--ink);
    letter-spacing: -0.03em;
  }}
  .kpi-card .etiqueta {{
    margin-top: 6px;
    font-size: 11px;
    font-weight: 700;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.07em;
  }}
  .kpi-icon {{
    width: 48px;
    height: 48px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
  }}
  .kpi-icon.blue {{ background: var(--blue-soft); color: var(--blue); }}
  .kpi-icon.orange {{ background: var(--orange-soft); color: var(--orange); }}
  .kpi-icon.cyan {{ background: rgba(13, 148, 136, 0.12); color: var(--cyan); }}

  .nav-tabs-wrap {{
    max-width: 1380px;
    margin: 0 auto 20px;
    padding: 0 24px;
  }}
  .nav-tabs {{
    display: inline-flex;
    background: var(--surface);
    border: 1px solid var(--line);
    padding: 5px;
    border-radius: 12px;
    gap: 4px;
    box-shadow: var(--shadow-sm);
    flex-wrap: wrap;
  }}
  .nav-tab {{
    border: 0;
    background: transparent;
    color: var(--muted);
    font: inherit;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 8px;
    cursor: pointer;
    transition: 0.15s;
  }}
  .nav-tab:hover {{ color: var(--ink); background: var(--surface-2); }}
  .nav-tab.activo {{ background: var(--blue); color: #fff; box-shadow: 0 2px 8px rgba(37, 99, 235, 0.25); }}

  section {{
    max-width: 1380px;
    margin: 0 auto;
    padding: 0 24px 56px;
  }}
  .panel {{
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: var(--radius);
    padding: 26px 28px;
    box-shadow: var(--shadow-sm);
    margin-bottom: 24px;
  }}
  .panel-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 14px;
    margin-bottom: 18px;
    padding-bottom: 14px;
    border-bottom: 1px solid var(--line);
  }}
  .panel-title {{
    display: flex;
    align-items: center;
    gap: 10px;
    margin: 0;
    font-size: 18px;
    font-weight: 700;
    letter-spacing: -0.02em;
    color: var(--ink);
  }}
  .panel-title::before {{
    content: '';
    width: 4px;
    height: 20px;
    border-radius: 99px;
    background: linear-gradient(180deg, var(--blue), var(--cyan));
  }}
  .notas {{
    color: var(--muted);
    font-size: 12.5px;
    line-height: 1.6;
    margin: 0 0 18px;
  }}

  .barra-controles {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    flex-wrap: wrap;
    padding: 14px 16px;
    border: 1px solid var(--line);
    border-radius: 12px;
    background: var(--surface-2);
    margin-bottom: 20px;
  }}
  .controles-grupo {{
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
  }}
  select#selectorEstacion, input#inputFecha, .input-busqueda {{
    font: inherit;
    font-size: 13.5px;
    font-weight: 600;
    color: var(--ink);
    background: var(--surface);
    padding: 9px 13px;
    border: 1px solid var(--line);
    border-radius: 9px;
    outline: none;
    transition: 0.15s;
  }}
  select#selectorEstacion {{ min-width: 180px; cursor: pointer; }}
  select#selectorEstacion:focus, input#inputFecha:focus, .input-busqueda:focus {{
    border-color: var(--blue);
    box-shadow: 0 0 0 3px var(--blue-soft);
  }}
  .toggle-modo {{
    display: inline-flex;
    border: 1px solid var(--line);
    border-radius: 9px;
    overflow: hidden;
    background: var(--surface);
  }}
  .toggle-modo button {{
    border: 0;
    border-right: 1px solid var(--line);
    background: transparent;
    color: var(--muted);
    padding: 9px 14px;
    cursor: pointer;
    font: inherit;
    font-size: 12.5px;
    font-weight: 600;
    transition: 0.15s;
  }}
  .toggle-modo button:last-child {{ border-right: 0; }}
  .toggle-modo button:hover {{ background: var(--surface-hover); color: var(--ink); }}
  .toggle-modo button.activo {{ background: var(--blue); color: #fff; }}

  .btn-accion {{
    font: inherit;
    font-size: 12.5px;
    font-weight: 600;
    padding: 9px 13px;
    border-radius: 9px;
    border: 1px solid var(--line);
    background: var(--surface);
    color: var(--ink);
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    transition: 0.15s;
  }}
  .btn-accion:hover {{
    border-color: var(--blue);
    color: var(--blue);
    background: var(--surface-hover);
  }}
  .btn-accion.pico {{
    background: var(--orange-soft);
    color: var(--orange);
    border-color: rgba(217, 119, 6, 0.3);
  }}
  .btn-accion.pico:hover {{
    background: var(--orange);
    color: #fff;
  }}
  .dia-semana-pill {{
    font-size: 12px;
    font-weight: 700;
    padding: 6px 11px;
    border-radius: 8px;
    background: var(--blue-soft);
    color: var(--blue);
  }}

  .kpis-estacion {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 12px;
    margin-bottom: 22px;
  }}
  .kpis-estacion .kpi-card {{
    padding: 15px 18px;
    border-radius: 12px;
    box-shadow: none;
    background: var(--surface-2);
    display: block;
  }}
  .kpis-estacion .valor {{ font-size: 24px; }}

  .tabla-scroll, .tabla-turnos-wrap {{
    max-height: 440px;
    overflow: auto;
    border: 1px solid var(--line);
    border-radius: 12px;
    background: var(--surface);
  }}
  table.resumen, table.tabla-turnos, table.semana, table.consolidado {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 13px;
  }}
  table th {{
    background: var(--surface-2);
    color: var(--muted);
    text-align: left;
    font-size: 10.5px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    padding: 12px 14px;
    border-bottom: 1px solid var(--line);
    position: sticky;
    top: 0;
    z-index: 1;
  }}
  table td {{
    padding: 11px 14px;
    border-bottom: 1px solid var(--line);
    vertical-align: middle;
    color: var(--ink-soft);
  }}
  table tr:last-child td {{ border-bottom: 0; }}
  table tbody tr:hover td {{ background: var(--surface-hover); }}

  .iata-pill {{
    display: inline-block;
    font-family: ui-monospace, SFMono-Regular, Consolas, monospace;
    font-weight: 700;
    font-size: 12px;
    padding: 3px 9px;
    border-radius: 6px;
    background: var(--blue-soft);
    color: var(--blue);
  }}
  .badge-total {{
    font-weight: 800;
    color: var(--ink);
    background: var(--surface-2);
    padding: 3px 8px;
    border-radius: 6px;
    border: 1px solid var(--line);
  }}
  .badge-pico {{
    font-weight: 700;
    color: var(--orange);
    background: var(--orange-soft);
    padding: 3px 8px;
    border-radius: 6px;
  }}
  .turno-badge {{
    display: inline-block;
    font-weight: 700;
    font-size: 11px;
    padding: 2px 8px;
    border-radius: 6px;
    background: var(--surface-2);
    border: 1px solid var(--line);
    color: var(--ink);
  }}
  .hora-mono {{
    font-family: ui-monospace, SFMono-Regular, Consolas, monospace;
    font-weight: 600;
    font-size: 12.5px;
    color: var(--ink);
  }}
  .tag-offset {{
    font-size: 10px;
    font-weight: 600;
    padding: 1px 6px;
    border-radius: 4px;
    background: var(--orange-soft);
    color: var(--orange);
    margin-left: 4px;
  }}
  .chip-vuelo {{
    display: inline-block;
    font-family: ui-monospace, SFMono-Regular, Consolas, monospace;
    font-size: 11px;
    font-weight: 600;
    padding: 2px 7px;
    margin: 2px 3px 2px 0;
    border-radius: 5px;
    background: var(--surface-2);
    border: 1px solid var(--line);
    color: var(--ink-soft);
  }}

  .util-wrapper {{
    display: flex;
    align-items: center;
    gap: 8px;
    min-width: 110px;
  }}
  .util-bar {{
    flex: 1;
    height: 7px;
    background: var(--surface-2);
    border-radius: 99px;
    overflow: hidden;
    border: 1px solid var(--line);
  }}
  .util-fill {{ height: 100%; border-radius: 99px; }}
  .util-text {{ font-size: 11.5px; font-weight: 700; min-width: 34px; text-align: right; }}
  .utilizacion-alta .util-fill {{ background: var(--success); }}
  .utilizacion-alta .util-text {{ color: var(--success); }}
  .utilizacion-media .util-fill {{ background: var(--blue); }}
  .utilizacion-media .util-text {{ color: var(--blue); }}
  .utilizacion-baja .util-fill {{ background: var(--danger); }}
  .utilizacion-baja .util-text {{ color: var(--danger); }}

  .turnos-grid {{
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 20px;
    margin-top: 16px;
  }}
  .servicio-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: 0 0 10px;
    font-size: 14px;
    font-weight: 700;
    color: var(--ink);
  }}
  .badge {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; }}
  .badge-checkin {{ background: var(--blue); }}
  .badge-gate {{ background: var(--orange); }}
  .empty-state {{
    padding: 20px;
    text-align: center;
    color: var(--muted);
    font-style: italic;
    background: var(--surface-2);
    border-radius: 10px;
    border: 1px dashed var(--line);
  }}

  .gantt-card {{
    background: var(--surface-2);
    border: 1px solid var(--line);
    border-radius: 12px;
    padding: 16px 18px;
    margin: 16px 0 22px;
  }}
  .gantt-top {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
    margin-bottom: 12px;
  }}
  .gantt-title {{ font-size: 13px; font-weight: 700; color: var(--ink); }}
  .gantt-legend {{
    display: flex;
    gap: 14px;
    flex-wrap: wrap;
    font-size: 11.5px;
    color: var(--muted);
  }}
  .gantt-legend span {{ display: inline-flex; align-items: center; gap: 6px; }}
  .leg-box {{ width: 12px; height: 12px; border-radius: 3px; display: inline-block; }}
  .gantt {{ font-size: 11.5px; overflow-x: auto; padding-bottom: 6px; }}
  .gantt-horas {{
    position: relative;
    height: 22px;
    border-bottom: 1px solid var(--line);
    margin-left: 145px;
    min-width: 680px;
  }}
  .gantt-hora {{
    position: absolute;
    transform: translateX(-50%);
    color: var(--muted);
    font-size: 10px;
    font-weight: 600;
  }}
  .gantt-separador {{
    font-size: 10.5px;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--muted);
    margin: 10px 0 6px;
    padding-top: 6px;
    border-top: 1px dashed var(--line);
  }}
  .gantt-fila {{ display: flex; align-items: center; margin-bottom: 5px; min-width: 825px; }}
  .gantt-etiqueta {{
    width: 145px;
    flex-shrink: 0;
    color: var(--ink-soft);
    font-weight: 600;
    font-size: 11px;
    padding-right: 10px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }}
  .gantt-pista {{
    position: relative;
    flex: 1;
    height: 24px;
    background: repeating-linear-gradient(
      90deg,
      var(--surface) 0,
      var(--surface) calc(8.333% - 1px),
      var(--line) 8.333%
    );
    border-radius: 6px;
    border: 1px solid var(--line);
  }}
  .gantt-barra {{
    position: absolute;
    top: 2px;
    height: 18px;
    border-radius: 5px;
    color: #fff;
    font-size: 10px;
    font-weight: 600;
    line-height: 18px;
    padding: 0 6px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    cursor: pointer;
    transition: transform 0.1s, filter 0.1s;
  }}
  .gantt-barra:hover {{ filter: brightness(1.1); z-index: 5; }}
  .barra-vuelo-checkin {{ background: #60a5fa; color: #0f172a; }}
  .barra-turno-checkin {{ background: #1d4ed8; height: 14px; top: 4px; box-shadow: 0 2px 4px rgba(29,78,216,0.25); }}
  .barra-vuelo-gate {{ background: #fbbf24; color: #0f172a; }}
  .barra-turno-gate {{ background: #b45309; height: 14px; top: 4px; box-shadow: 0 2px 4px rgba(180,83,9,0.25); }}

  .gantt-tooltip {{
    position: fixed;
    display: none;
    background: #0f172a;
    color: #f8fafc;
    font-size: 12px;
    padding: 8px 12px;
    border-radius: 8px;
    pointer-events: none;
    z-index: 999;
    box-shadow: 0 10px 25px rgba(0,0,0,0.35);
    border: 1px solid rgba(255,255,255,0.12);
    white-space: nowrap;
  }}

  table.semana tr.sin-vuelos td {{ color: var(--muted); font-style: italic; }}
  details.detalle-turnos summary {{ cursor: pointer; color: var(--blue); font-size: 12px; font-weight: 700; }}
  details.detalle-turnos .turnos-grid {{ margin-top: 10px; }}
  tr.fila-total td {{
    background: var(--blue-soft) !important;
    border-top: 2px solid var(--blue);
    font-weight: 800;
    color: var(--ink);
  }}

  @media (max-width: 900px) {{
    .kpis {{ grid-template-columns: 1fr; margin-top: -36px; }}
    .turnos-grid {{ grid-template-columns: 1fr; }}
    .panel {{ padding: 18px; }}
    section {{ padding: 0 14px 40px; }}
  }}
  @media print {{
    .nav-tabs-wrap, .barra-controles, .btn-theme, .notas {{ display: none !important; }}
    header {{ padding: 16px 24px; background: #fff !important; color: #000 !important; border-bottom: 2px solid #000; }}
    .kpis {{ margin-top: 12px; }}
    .panel {{ box-shadow: none; border: 1px solid #ccc; page-break-inside: avoid; }}
    body {{ background: #fff; color: #000; }}
  }}
</style>
</head>
<body>

<div id="ganttTooltip" class="gantt-tooltip"></div>

<header>
  <div class="header-inner">
    <div>
      <span class="brand-badge">● Planificación Operativa Aeroportuaria</span>
      <h1>Dimensionamiento de Personal — Check-in &amp; Gate</h1>
      <p>
        <span>📅 Periodo analizado: <b>{kpi_fecha_min:%d-%b-%Y}</b> a <b>{kpi_fecha_max:%d-%b-%Y}</b></span>
        <span>•</span>
        <span>⏱️ Turnos de 8h (buffer de 15 min incluido)</span>
      </p>
    </div>
    <div class="header-actions">
      <button class="btn-theme" onclick="alternarTema()" id="btnTema">🌙 Modo Oscuro</button>
    </div>
  </div>
</header>

<div class="kpis">
  <div class="kpi-card">
    <div>
      <div class="valor">{kpi_headcount_prom}</div>
      <div class="etiqueta">Headcount Promedio / Día (Red)</div>
    </div>
    <div class="kpi-icon blue">
      <svg width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
    </div>
  </div>
  <div class="kpi-card">
    <div>
      <div class="valor">{kpi_headcount_pico}</div>
      <div class="etiqueta">Headcount Pico (Día Más Crítico)</div>
    </div>
    <div class="kpi-icon orange">
      <svg width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>
    </div>
  </div>
  <div class="kpi-card">
    <div>
      <div class="valor">{kpi_num_estaciones}</div>
      <div class="etiqueta">Estaciones Analizadas</div>
    </div>
    <div class="kpi-icon cyan">
      <svg width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M17.8 19.2 16 11l3.5-3.5C21 6 21.5 4 21 3c-1-.5-3 0-4.5 1.5L13 8 4.8 6.2c-.5-.1-.9.1-1.1.5l-.3.5c-.2.5-.1 1 .3 1.3L9 12l-2 3H4l-1 1 3 2 2 3 1-1v-3l3-2 3.5 5.3c.3.4.8.5 1.3.3l.5-.2c.4-.3.6-.7.5-1.2z"/></svg>
    </div>
  </div>
</div>

<div class="nav-tabs-wrap">
  <div class="nav-tabs">
    <button class="nav-tab activo" onclick="cambiarSeccion('todo', this)">Vista Completa</button>
    <button class="nav-tab" onclick="cambiarSeccion('operacion', this)">Programación de Turnos &amp; Gantt</button>
    <button class="nav-tab" onclick="cambiarSeccion('analitica', this)">Mapas de Calor &amp; Curva Diaria</button>
    <button class="nav-tab" onclick="cambiarSeccion('resumen', this)">Resumen Consolidado de Red</button>
  </div>
</div>

<section>
  <!-- PANEL 1: PROGRAMACION DE TURNOS POR FECHA -->
  <div class="panel" id="panelCalendario">
    <div class="panel-header">
      <h2 class="panel-title">Programación de Turnos y Cobertura por Fecha</h2>
      <span id="etiquetaDiaSemana" class="dia-semana-pill"></span>
    </div>
    <p class="notas">
      Filtra por día puntual, por semana completa, o compara todas las estaciones en una misma fecha.
      <b>Ocupación</b> muestra el porcentaje del turno de 8h cubierto efectivamente por vuelos (<span style="color:var(--danger);font-weight:700">&lt;30% baja ocupación</span>).
    </p>

    <div class="barra-controles">
      <div class="controles-grupo">
        <select id="selectorEstacion" onchange="cambiarEstacion(this.value)">
          {opciones_select}
        </select>
        <div class="toggle-modo">
          <button id="btnModoDia" class="activo" onclick="setModoFecha('dia')">Día</button>
          <button id="btnModoSemana" onclick="setModoFecha('semana')">Semana</button>
          <button id="btnModoTodas" onclick="setModoFecha('todas')">Todas las estaciones</button>
        </div>
      </div>

      <div class="controles-grupo">
        <button class="btn-accion" onclick="desplazarFecha(-1)" title="Día anterior">◀ Ant.</button>
        <input type="date" id="inputFecha" onchange="actualizarVistaFecha()">
        <button class="btn-accion" onclick="desplazarFecha(1)" title="Día siguiente">Sig. ▶</button>
        <button class="btn-accion pico" onclick="irADiaCritico()" title="Saltar al día de mayor demanda de esta estación">🔥 Día Pico</button>
      </div>

      <div class="controles-grupo">
        <button class="btn-accion" onclick="descargarCSV()">⬇ Exportar CSV</button>
        <button class="btn-accion" onclick="window.print()">🖨️ Imprimir</button>
      </div>
    </div>

    <div id="resultadoFecha"></div>
  </div>

  <!-- PANEL 2: MAPAS DE CALOR Y DEMANDA POR ESTACION -->
  <div class="panel" id="panelAnalitica">
    <div class="panel-header">
      <h2 class="panel-title">Planeación — Mapas de Calor y Demanda por Estación</h2>
    </div>
    <div class="kpis-estacion" id="kpisEstacion">
      <div class="kpi-card"><div class="valor" id="kpiVuelos">{d0['vuelos_prom']}</div><div class="etiqueta">Vuelos / Día (Prom.)</div></div>
      <div class="kpi-card"><div class="valor" id="kpiCheckin">{d0['checkin_prom']}</div><div class="etiqueta">Personal Check-in (Prom.)</div></div>
      <div class="kpi-card"><div class="valor" id="kpiGate">{d0['gate_prom']}</div><div class="etiqueta">Personal Gate (Prom.)</div></div>
      <div class="kpi-card"><div class="valor" id="kpiTotal">{d0['total_prom']}</div><div class="etiqueta">Total Personal (Prom.)</div></div>
      <div class="kpi-card"><div class="valor" id="kpiPico">{d0['total_pico']}</div><div class="etiqueta">Total Personal (Día Pico)</div></div>
    </div>

    {grafico_html}
  </div>

  <!-- PANEL 3: RESUMEN DE RED -->
  <div class="panel" id="panelResumen">
    <div class="panel-header">
      <h2 class="panel-title">Resumen General por Estación (Toda la Red)</h2>
      <input type="text" id="buscadorEstacion" class="input-busqueda" placeholder="🔍 Filtrar estación (ej. BOG, MTR)..." oninput="filtrarTablaResumen(this.value)">
    </div>
    <div class="tabla-scroll">
      <table class="resumen" id="tablaResumenRed">
        <thead>
          <tr>
            <th>Estación</th>
            <th>Vuelos / Día (Prom.)</th>
            <th>Check-in (Prom.)</th>
            <th>Gate (Prom.)</th>
            <th>Total (Prom.)</th>
            <th>Total (Día Pico)</th>
          </tr>
        </thead>
        <tbody>
          {filas_tabla}
        </tbody>
      </table>
    </div>
    <p class="notas" style="margin-top:12px;margin-bottom:0">
      💡 <b>Tip:</b> Haz clic sobre cualquier fila de la tabla para cargar automáticamente esa estación en el programador de turnos y en las gráficas. Los promedios están redondeados hacia arriba.
    </p>
  </div>
</section>

<script>
const datosEstaciones = {datos_json};
const DIAS_SEMANA_JS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
let modoFecha = 'dia';
let datosCSVActual = {{ encabezados: [], filas: [] }};

// ---- Helpers de formato (solo afectan como se escribe la fecha/CSV) ----
function fmtFecha(iso) {{
  const [y, m, d] = iso.split('-');
  return `${{d}}/${{m}}/${{y}}`;
}}
function fmtFechaHora(iso) {{
  const d = new Date(iso);
  const p = n => String(n).padStart(2, '0');
  return `${{p(d.getDate())}}/${{p(d.getMonth() + 1)}}/${{d.getFullYear()}} ${{p(d.getHours())}}:${{p(d.getMinutes())}}`;
}}
function fmtISOLocal(d) {{
  const p = n => String(n).padStart(2, '0');
  return `${{d.getFullYear()}}-${{p(d.getMonth() + 1)}}-${{p(d.getDate())}}`;
}}
function celdaCSV(v) {{
  return '"' + String(v ?? '').replace(/"/g, '""') + '"';
}}

function alternarTema() {{
  const root = document.documentElement;
  const actual = root.getAttribute('data-theme') || 'light';
  const nuevo = actual === 'light' ? 'dark' : 'light';
  root.setAttribute('data-theme', nuevo);
  document.getElementById('btnTema').textContent = nuevo === 'dark' ? '☀️ Modo Claro' : '🌙 Modo Oscuro';
  localStorage.setItem('temaDashboardTurnos', nuevo);

  const colorTexto = nuevo === 'dark' ? '#cbd5e1' : '#334155';
  const colorGrid = nuevo === 'dark' ? 'rgba(148,163,184,0.12)' : 'rgba(148,163,184,0.18)';
  if (window.Plotly) {{
    Plotly.relayout('mainPlot', {{
      'font.color': colorTexto,
      'xaxis3.gridcolor': colorGrid,
      'yaxis3.gridcolor': colorGrid
    }});
  }}
}}

function cambiarSeccion(seccion, btn) {{
  document.querySelectorAll('.nav-tab').forEach(b => b.classList.remove('activo'));
  btn.classList.add('activo');

  const pCal = document.getElementById('panelCalendario');
  const pAna = document.getElementById('panelAnalitica');
  const pRes = document.getElementById('panelResumen');

  pCal.style.display = (seccion === 'todo' || seccion === 'operacion') ? 'block' : 'none';
  pAna.style.display = (seccion === 'todo' || seccion === 'analitica') ? 'block' : 'none';
  pRes.style.display = (seccion === 'todo' || seccion === 'resumen') ? 'block' : 'none';

  if ((seccion === 'todo' || seccion === 'analitica') && window.Plotly) {{
    setTimeout(() => Plotly.Plots.resize('mainPlot'), 60);
  }}
}}

function seleccionarEstacionDesdeTabla(est) {{
  const sel = document.getElementById('selectorEstacion');
  sel.value = est;
  cambiarEstacion(est);
  window.scrollTo({{ top: 0, behavior: 'smooth' }});
}}

function filtrarTablaResumen(texto) {{
  const filtro = texto.trim().toUpperCase();
  const filas = document.querySelectorAll('#tablaResumenRed tbody tr');
  filas.forEach(tr => {{
    const est = tr.cells[0].textContent.toUpperCase();
    tr.style.display = est.includes(filtro) ? '' : 'none';
  }});
}}

function cambiarEstacion(est) {{
  const d = datosEstaciones[est];
  if (window.Plotly) {{
    Plotly.restyle('mainPlot', {{visible: d.visible}});
  }}
  document.getElementById('kpiVuelos').textContent = d.vuelos_prom;
  document.getElementById('kpiCheckin').textContent = d.checkin_prom;
  document.getElementById('kpiGate').textContent = d.gate_prom;
  document.getElementById('kpiTotal').textContent = d.total_prom;
  document.getElementById('kpiPico').textContent = d.total_pico;

  const inputFecha = document.getElementById('inputFecha');
  inputFecha.min = d.fecha_min;
  inputFecha.max = d.fecha_max;
  inputFecha.value = d.fecha_critica;
  actualizarVistaFecha();
}}

function desplazarFecha(dias) {{
  const inputFecha = document.getElementById('inputFecha');
  if (!inputFecha.value) return;
  const f = new Date(inputFecha.value + 'T00:00:00');
  f.setDate(f.getDate() + dias);
  const iso = fmtISOLocal(f);
  if (inputFecha.min && iso < inputFecha.min) return;
  if (inputFecha.max && iso > inputFecha.max) return;
  inputFecha.value = iso;
  actualizarVistaFecha();
}}

function irADiaCritico() {{
  const est = document.getElementById('selectorEstacion').value;
  const d = datosEstaciones[est];
  document.getElementById('inputFecha').value = d.fecha_critica;
  actualizarVistaFecha();
}}

function setModoFecha(modo) {{
  modoFecha = modo;
  document.getElementById('btnModoDia').classList.toggle('activo', modo === 'dia');
  document.getElementById('btnModoSemana').classList.toggle('activo', modo === 'semana');
  document.getElementById('btnModoTodas').classList.toggle('activo', modo === 'todas');
  document.getElementById('selectorEstacion').disabled = (modo === 'todas');
  actualizarVistaFecha();
}}

function actualizarVistaFecha() {{
  const est = document.getElementById('selectorEstacion').value;
  const d = datosEstaciones[est];
  const fechaISO = document.getElementById('inputFecha').value;
  const cont = document.getElementById('resultadoFecha');
  const lblDia = document.getElementById('etiquetaDiaSemana');

  if (!fechaISO) {{
    cont.innerHTML = '';
    lblDia.textContent = '';
    datosCSVActual = {{encabezados: [], filas: []}};
    return;
  }}

  const fObj = new Date(fechaISO + 'T00:00:00');
  const idxDia = (fObj.getDay() + 6) % 7;
  const esCritico = (fechaISO === d.fecha_critica && modoFecha === 'dia');
  lblDia.innerHTML = `${{DIAS_SEMANA_JS[idxDia]}} · ${{fechaISO}} ${{esCritico ? '🔥 (Día Pico)' : ''}}`;

  if (modoFecha === 'dia') {{
    const info = d.diario[fechaISO];
    if (!info) {{
      cont.innerHTML = '<div class="empty-state">No hay vuelos programados para esa fecha en esta estación.</div>';
      datosCSVActual = {{encabezados: [], filas: []}};
      return;
    }}
    cont.innerHTML = `
      <div class="kpis-estacion">
        <div class="kpi-card"><div class="valor">${{info.vuelos}}</div><div class="etiqueta">Vuelos del Día</div></div>
        <div class="kpi-card"><div class="valor" style="color:var(--blue)">${{info.checkin}}</div><div class="etiqueta">Turnos Check-in (8h)</div></div>
        <div class="kpi-card"><div class="valor" style="color:var(--orange)">${{info.gate}}</div><div class="etiqueta">Turnos Gate (8h)</div></div>
        <div class="kpi-card"><div class="valor">${{info.total}}</div><div class="etiqueta">Total Personal Requerido</div></div>
      </div>
      <div id="ganttDia"></div>
      <div class="turnos-grid">
        <div>
          <div class="servicio-header">
            <span><span class="badge badge-checkin"></span>Detalle de Turnos — Check-in</span>
            <span class="turno-badge">${{info.checkin}} turnos</span>
          </div>
          ${{info.turnos_checkin_html}}
        </div>
        <div>
          <div class="servicio-header">
            <span><span class="badge badge-gate"></span>Detalle de Turnos — Gate / Salas</span>
            <span class="turno-badge">${{info.gate}} turnos</span>
          </div>
          ${{info.turnos_gate_html}}
        </div>
      </div>`;
    dibujarGantt('ganttDia', info);

    const filasCSV = [];
    info.turnos_checkin_raw.forEach((t, i) => filasCSV.push([
      est, fmtFecha(info.fecha_iso), 'Check-in', 'Turno ' + (i + 1),
      fmtFechaHora(t.entrada), fmtFechaHora(t.salida), t.vuelos.join(' '), t.utilizacion + '%']));
    info.turnos_gate_raw.forEach((t, i) => filasCSV.push([
      est, fmtFecha(info.fecha_iso), 'Gate', 'Turno ' + (i + 1),
      fmtFechaHora(t.entrada), fmtFechaHora(t.salida), t.vuelos.join(' '), t.utilizacion + '%']));
    datosCSVActual = {{
      encabezados: ['Estacion', 'Fecha', 'Servicio', 'Turno', 'Entrada', 'Salida', 'Vuelos que cubre', 'Utilizacion'],
      filas: filasCSV,
    }};
  }} else if (modoFecha === 'semana') {{
    const f0 = new Date(fechaISO + 'T00:00:00');
    const offsetLunes = (f0.getDay() + 6) % 7;
    const lunes = new Date(f0);
    lunes.setDate(f0.getDate() - offsetLunes);

    let filas = '';
    let totalSemana = {{vuelos: 0, checkin: 0, gate: 0, total: 0}};
    const filasCSV = [];
    for (let i = 0; i < 7; i++) {{
      const f = new Date(lunes);
      f.setDate(lunes.getDate() + i);
      const iso = fmtISOLocal(f);
      const info = d.diario[iso];

      if (info) {{
        totalSemana.vuelos += info.vuelos;
        totalSemana.checkin += info.checkin;
        totalSemana.gate += info.gate;
        totalSemana.total += info.total;
        filasCSV.push([est, fmtFecha(iso), DIAS_SEMANA_JS[i], info.vuelos, info.checkin, info.gate, info.total]);
        filas += `<tr>
          <td><b>${{DIAS_SEMANA_JS[i]}}</b></td>
          <td class="hora-mono">${{iso}}</td>
          <td>${{info.vuelos}}</td>
          <td>${{info.checkin}}</td>
          <td>${{info.gate}}</td>
          <td><span class="badge-total">${{info.total}}</span></td>
          <td><details class="detalle-turnos"><summary>Ver horarios de turnos</summary>
              <div class="turnos-grid"><div>${{info.turnos_checkin_html}}</div><div>${{info.turnos_gate_html}}</div></div>
              </details></td>
        </tr>`;
      }} else {{
        filas += `<tr class="sin-vuelos"><td>${{DIAS_SEMANA_JS[i]}}</td><td class="hora-mono">${{iso}}</td><td colspan="5">Sin vuelos programados</td></tr>`;
      }}
    }}

    cont.innerHTML = `<div class="tabla-scroll"><table class="semana">
      <thead><tr><th>Día</th><th>Fecha</th><th>Vuelos</th><th>Check-in</th><th>Gate</th><th>Total</th><th>Detalle</th></tr></thead>
      <tbody>${{filas}}
      <tr class="fila-total"><td colspan="2">Total Semana (${{est}})</td>
        <td>${{totalSemana.vuelos}}</td><td>${{totalSemana.checkin}}</td>
        <td>${{totalSemana.gate}}</td><td>${{totalSemana.total}}</td><td></td></tr>
      </tbody></table></div>`;
    datosCSVActual = {{
      encabezados: ['Estacion', 'Fecha', 'Dia', 'Vuelos', 'Check-in', 'Gate', 'Total'],
      filas: filasCSV,
    }};
  }} else {{
    const estaciones = Object.keys(datosEstaciones).sort();
    let filas = '';
    let totalRed = {{vuelos: 0, checkin: 0, gate: 0, total: 0}};
    const filasCSV = [];
    estaciones.forEach(nombreEst => {{
      const info = datosEstaciones[nombreEst].diario[fechaISO];
      if (info) {{
        totalRed.vuelos += info.vuelos;
        totalRed.checkin += info.checkin;
        totalRed.gate += info.gate;
        totalRed.total += info.total;
        filasCSV.push([nombreEst, fmtFecha(fechaISO), info.vuelos, info.checkin, info.gate, info.total]);
        filas += `<tr onclick="seleccionarEstacionDesdeTabla('${{nombreEst}}'); setModoFecha('dia');" style="cursor:pointer">
          <td><span class="iata-pill">${{nombreEst}}</span></td>
          <td>${{info.vuelos}}</td>
          <td>${{info.checkin}}</td>
          <td>${{info.gate}}</td>
          <td><span class="badge-total">${{info.total}}</span></td>
        </tr>`;
      }} else {{
        filas += `<tr class="sin-vuelos"><td><span class="iata-pill">${{nombreEst}}</span></td><td colspan="4">Sin vuelos programados</td></tr>`;
      }}
    }});

    cont.innerHTML = `<div class="tabla-scroll"><table class="consolidado">
      <thead><tr><th>Estación</th><th>Vuelos</th><th>Check-in</th><th>Gate</th><th>Total Turnos</th></tr></thead>
      <tbody>${{filas}}
      <tr class="fila-total"><td>Total Red (${{fechaISO}})</td><td>${{totalRed.vuelos}}</td>
        <td>${{totalRed.checkin}}</td><td>${{totalRed.gate}}</td><td>${{totalRed.total}}</td></tr>
      </tbody></table></div>`;
    datosCSVActual = {{
      encabezados: ['Estacion', 'Fecha', 'Vuelos', 'Check-in', 'Gate', 'Total'],
      filas: filasCSV,
    }};
  }}
}}

function descargarCSV() {{
  if (!datosCSVActual.filas.length) return;
  let csv = datosCSVActual.encabezados.map(celdaCSV).join(';') + '\\r\\n';
  csv += datosCSVActual.filas.map(f => f.map(celdaCSV).join(';')).join('\\r\\n');
  const blob = new Blob(['\\ufeff' + csv], {{type: 'text/csv;charset=utf-8;'}});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const fechaISO = document.getElementById('inputFecha').value || 'export';
  a.href = url;
  a.download = `turnos_${{modoFecha}}_${{fechaISO}}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}}

function dibujarGantt(contenedorId, info) {{
  const filas = [];
  info.ventanas_checkin.forEach(v => filas.push({{...v, seccion: 'CHECK-IN', grupo: '✈ Vuelo ' + v.vuelo, clase: 'barra-vuelo-checkin'}}));
  info.turnos_checkin_raw.forEach((t, i) => filas.push({{
    inicio: t.entrada, fin: t.salida, vuelo: 'T' + (i + 1) + ' (' + t.utilizacion + '%): ' + t.vuelos.join(', '),
    seccion: 'CHECK-IN', grupo: '👤 Turno Check-in ' + (i + 1), clase: 'barra-turno-checkin'}}));
  info.ventanas_gate.forEach(v => filas.push({{...v, seccion: 'GATE / SALAS', grupo: '✈ Vuelo ' + v.vuelo, clase: 'barra-vuelo-gate'}}));
  info.turnos_gate_raw.forEach((t, i) => filas.push({{
    inicio: t.entrada, fin: t.salida, vuelo: 'T' + (i + 1) + ' (' + t.utilizacion + '%): ' + t.vuelos.join(', '),
    seccion: 'GATE / SALAS', grupo: '👤 Turno Gate ' + (i + 1), clase: 'barra-turno-gate'}}));

  const cont = document.getElementById(contenedorId);
  if (filas.length === 0) {{ cont.innerHTML = ''; return; }}

  const tiempos = filas.flatMap(f => [new Date(f.inicio).getTime(), new Date(f.fin).getTime()]);
  const minT = Math.min(...tiempos) - 30 * 60000;
  const maxT = Math.max(...tiempos) + 30 * 60000;
  const rango = maxT - minT;

  const horas = [];
  let t = new Date(Math.ceil(minT / 3600000) * 3600000);
  while (t.getTime() <= maxT) {{
    horas.push(`<div class="gantt-hora" style="left:${{(t.getTime() - minT) / rango * 100}}%">${{t.getHours().toString().padStart(2,'0')}}:00</div>`);
    t = new Date(t.getTime() + 3600000);
  }}

  const formatearHora = (d) => d.toTimeString().slice(0, 5) +
    (d.toDateString() !== new Date(info.fecha_iso + 'T00:00:00').toDateString()
      ? ' (' + d.toLocaleDateString('es-CO', {{day: '2-digit', month: 'short'}}) + ')' : '');

  let seccionActual = '';
  const filasHtml = filas.map((f, idx) => {{
    let sepHtml = '';
    if (f.seccion !== seccionActual) {{
      seccionActual = f.seccion;
      sepHtml = `<div class="gantt-separador">${{seccionActual}}</div>`;
    }}
    const iniDate = new Date(f.inicio);
    const finDate = new Date(f.fin);
    const left = (iniDate.getTime() - minT) / rango * 100;
    const width = Math.max((finDate.getTime() - iniDate.getTime()) / rango * 100, 0.5);
    return `${{sepHtml}}<div class="gantt-fila">
      <div class="gantt-etiqueta" title="${{f.grupo}}">${{f.grupo}}</div>
      <div class="gantt-pista">
        <div class="gantt-barra ${{f.clase}}" data-idx="${{idx}}" style="left:${{left}}%;width:${{width}}%">${{f.vuelo}}</div>
      </div>
    </div>`;
  }}).join('');

  cont.innerHTML = `
    <div class="gantt-card">
      <div class="gantt-top">
        <div class="gantt-title">Línea de Tiempo Operativa (Vuelos vs. Turnos de 8h)</div>
        <div class="gantt-legend">
          <span><i class="leg-box" style="background:#60a5fa"></i>Ventana Vuelo Check-in</span>
          <span><i class="leg-box" style="background:#1d4ed8"></i>Turno Check-in (8h)</span>
          <span><i class="leg-box" style="background:#fbbf24"></i>Ventana Vuelo Gate</span>
          <span><i class="leg-box" style="background:#b45309"></i>Turno Gate (8h)</span>
        </div>
      </div>
      <div class="gantt"><div class="gantt-horas">${{horas.join('')}}</div>${{filasHtml}}</div>
    </div>`;

  const tips = filas.map(f => {{
    const iniDate = new Date(f.inicio);
    const finDate = new Date(f.fin);
    return `${{f.grupo}} — ${{f.vuelo}} (${{formatearHora(iniDate)}} a ${{formatearHora(finDate)}})`;
  }});
  const tooltip = document.getElementById('ganttTooltip');
  cont.querySelectorAll('.gantt-barra').forEach(el => {{
    el.addEventListener('mouseenter', () => {{
      tooltip.textContent = tips[Number(el.dataset.idx)];
      tooltip.style.display = 'block';
    }});
    el.addEventListener('mousemove', (e) => {{
      tooltip.style.left = (e.clientX + 14) + 'px';
      tooltip.style.top = (e.clientY + 14) + 'px';
    }});
    el.addEventListener('mouseleave', () => {{ tooltip.style.display = 'none'; }});
  }});
}}

(function inicializarDashboard() {{
  const temaGuardado = localStorage.getItem('temaDashboardTurnos');
  if (temaGuardado === 'dark') {{
    alternarTema();
  }}
  const primerEst = document.getElementById('selectorEstacion').value;
  const d = datosEstaciones[primerEst];
  const inputFecha = document.getElementById('inputFecha');
  inputFecha.min = d.fecha_min;
  inputFecha.max = d.fecha_max;
  inputFecha.value = d.fecha_critica;
  actualizarVistaFecha();
}})();
</script>

</body>
</html>"""

    with open(nombre_html, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Dashboard guardado en '{nombre_html}'.")

# ============================================================
# ---- Interfaz grafica (para que cualquier persona lo use sin saber
#      Python ni consola).
# ============================================================

COLUMNAS_REQUERIDAS = [
    "EFFECTIVE FROM", "EFFECTIVE TO", "DOW", "ETD",
    "DEPARTURESTATION", "ARRIVALSTATION", "FLIGHTNUMBER",
]


class AppPersonal:
    def __init__(self, root, archivo_inicial=None):
        self.root = root
        root.title("Calculadora de Personal - Check-in y Gate")
        root.geometry("560x400")
        root.resizable(False, False)
        root.configure(bg="#f4f6f8")

        tk.Label(
            root, text="Calculadora de Personal Requerido",
            font=("Segoe UI", 16, "bold"), bg="#f4f6f8", fg="#1a1a1a",
        ).pack(pady=(24, 4))

        tk.Label(
            root,
            text="Arrastra aqui tu Excel de itinerario, o usa el boton de abajo.",
            font=("Segoe UI", 10), bg="#f4f6f8", fg="#555555", justify="center",
        ).pack(pady=(0, 14))

        self.zona_drop = tk.Label(
            root, text="Suelta aqui el archivo .xlsx",
            font=("Segoe UI", 11), bg="#e8edf3", fg="#7a8695",
            relief="ridge", bd=2, width=48, height=5,
        )
        self.zona_drop.pack(pady=(0, 16))

        if _DND_DISPONIBLE:
            self.zona_drop.drop_target_register(DND_FILES)
            self.zona_drop.dnd_bind("<<Drop>>", self._al_soltar_archivo)
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._al_soltar_archivo)
        else:
            self.zona_drop.config(text="(arrastrar no disponible en este empaquetado,\nusa el boton de abajo)")

        self.boton = tk.Button(
            root, text="Seleccionar archivo Excel...", font=("Segoe UI", 11),
            command=self.seleccionar_archivo, bg="#2563eb", fg="white",
            activebackground="#1d4ed8", activeforeground="white",
            relief="flat", padx=18, pady=10, cursor="hand2",
        )
        self.boton.pack(pady=6)

        self.estado = tk.Label(
            root, text="", font=("Segoe UI", 10), bg="#f4f6f8", fg="#333333",
            wraplength=480, justify="center",
        )
        self.estado.pack(pady=(18, 6))

        self.barra = tk.Label(root, text="v1.1", font=("Segoe UI", 9), bg="#f4f6f8", fg="#888888")
        self.barra.pack(side="bottom", pady=10)

        if archivo_inicial:
            root.after(300, lambda: self.procesar(archivo_inicial))

    def _al_soltar_archivo(self, event):
        rutas = self.root.tk.splitlist(event.data)
        if rutas:
            self.procesar(rutas[0])

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el Excel de itinerario",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
        )
        if ruta:
            self.procesar(ruta)

    def set_estado(self, texto, color="#333333"):
        self.estado.config(text=texto, fg=color)

    def _estado_seguro(self, texto, color="#333333"):
        self.root.after(0, lambda: self.set_estado(texto, color))

    def procesar(self, ruta_entrada):
        if not str(ruta_entrada).lower().endswith((".xlsx", ".xls")):
            messagebox.showerror(
                "Archivo no valido",
                "Eso no parece un archivo Excel (.xlsx o .xls). Intenta de nuevo.",
            )
            return

        self.boton.config(state="disabled")
        self.set_estado("Procesando... esto puede tardar unos segundos, la ventana sigue activa.")
        hilo = threading.Thread(target=self._trabajo, args=(ruta_entrada,), daemon=True)
        hilo.start()

    def _trabajo(self, ruta_entrada):
        try:
            if not os.path.isfile(ruta_entrada):
                raise FileNotFoundError(f"No se encontro el archivo:\n{ruta_entrada}")

            self._estado_seguro("Cargando componentes (solo la primera vez tarda un poco)...")
            _cargar_dependencias_pesadas()

            self._estado_seguro("Leyendo el archivo Excel...")
            malla = pd.read_excel(ruta_entrada)

            faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in malla.columns]
            if faltantes:
                raise ValueError(
                    "El Excel no tiene el formato esperado.\n"
                    "Faltan estas columnas: " + ", ".join(faltantes)
                )

            self._estado_seguro("Calculando el personal necesario por dia y estacion...")
            resultado, vuelos = calcular_por_dia_y_estacion(malla)

            carpeta_entrada = os.path.dirname(os.path.abspath(ruta_entrada))
            carpeta_salida = os.path.join(carpeta_entrada, "Resultados")
            os.makedirs(carpeta_salida, exist_ok=True)

            ruta_excel_salida = os.path.join(carpeta_salida, "personal_requerido.xlsx")
            ruta_html_salida = os.path.join(carpeta_salida, "dashboard_dimensionamiento.html")

            self._estado_seguro("Guardando el archivo Excel de resultados...")
            resultado.to_excel(ruta_excel_salida, index=False)

            self._estado_seguro("Generando el dashboard interactivo...")
            cwd_original = os.getcwd()
            try:
                os.chdir(carpeta_salida)
                generar_dashboard_html(resultado, vuelos)
            finally:
                os.chdir(cwd_original)

            self.root.after(0, lambda: self._terminar_ok(carpeta_salida, ruta_html_salida))

        except Exception as exc:
            traceback.print_exc()
            self.root.after(0, lambda exc=exc: self._terminar_error(exc))

    def _terminar_ok(self, carpeta_salida, ruta_html_salida):
        self.set_estado(
            "Listo. Se guardaron los resultados en la carpeta 'Resultados',\n"
            "junto al archivo que seleccionaste. Abriendo el dashboard...",
            color="#166534",
        )
        webbrowser.open("file://" + os.path.abspath(ruta_html_salida))
        messagebox.showinfo(
            "Proceso terminado",
            "El calculo termino correctamente.\n\n"
            f"Los archivos quedaron en:\n{carpeta_salida}\n\n"
            "El dashboard se abrio en tu navegador.",
        )
        self.boton.config(state="normal")

    def _terminar_error(self, exc):
        self.set_estado("Ocurrio un error. Revisa el mensaje.", color="#b91c1c")
        messagebox.showerror(
            "No se pudo procesar el archivo",
            f"{exc}\n\n"
            "Verifica que sea el Excel de itinerario correcto e intentalo de nuevo.",
        )
        self.boton.config(state="normal")


# ============================================================
# ---- Main ----
# ============================================================
if __name__ == "__main__":
    archivo_arrastrado = sys.argv[1] if len(sys.argv) > 1 else None
    root = TkinterDnD.Tk() if _DND_DISPONIBLE else tk.Tk()
    app = AppPersonal(root, archivo_inicial=archivo_arrastrado)
    root.mainloop()