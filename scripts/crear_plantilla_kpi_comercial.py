"""
Genera la plantilla Excel de KPI comercial: cumplimiento y performance por
canal de ventas (Paneles / Otros) y por consultor de ventas (RSM).

Este script se ejecuta UNA sola vez (o cuando se quiera rediseñar la
plantilla). El archivo resultante
(plantilla/Plantilla_KPI_Comercial_RSM.xlsx) es el que el equipo comercial
actualiza cada periodo (mensual/YTD); la hoja "Dashboard" recalcula tablas,
semaforos y graficos automaticamente a partir de la hoja "Datos KPI".

Uso:
    python scripts/crear_plantilla_kpi_comercial.py
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

OUT_PATH = "plantilla/Plantilla_KPI_Comercial_RSM.xlsx"

FONT_NAME = "Arial"

COLOR_HEADER = "1E2761"    # navy
COLOR_HEADER_TEXT = "FFFFFF"
COLOR_INPUT = "FFF2CC"     # amarillo suave -> celdas para llenar
COLOR_SECTION = "CADCFC"   # azul hielo
COLOR_CALC = "F2F2F2"      # gris claro -> columnas calculadas (no editar)
COLOR_GREEN = "C6EFCE"
COLOR_YELLOW = "FFEB9C"
COLOR_RED = "FFC7CE"

header_font = Font(name=FONT_NAME, size=10, bold=True, color=COLOR_HEADER_TEXT)
header_fill = PatternFill("solid", fgColor=COLOR_HEADER)
section_font = Font(name=FONT_NAME, size=13, bold=True, color=COLOR_HEADER)
section_fill = PatternFill("solid", fgColor=COLOR_SECTION)
input_fill = PatternFill("solid", fgColor=COLOR_INPUT)
calc_fill = PatternFill("solid", fgColor=COLOR_CALC)
normal_font = Font(name=FONT_NAME, size=10)
input_font = Font(name=FONT_NAME, size=10, color="0000FF")
calc_font = Font(name=FONT_NAME, size=10)
title_font = Font(name=FONT_NAME, size=16, bold=True, color=COLOR_HEADER)
total_font = Font(name=FONT_NAME, size=10, bold=True)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
center_nowrap = Alignment(horizontal="center", vertical="center")

wb = Workbook()


def autofit(ws, widths, start_col=1):
    for i, w in enumerate(widths, start=start_col):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# 1. INSTRUCCIONES
# ---------------------------------------------------------------------------
ws = wb.active
ws.title = "Instrucciones"
ws.sheet_view.showGridLines = False
ws["B2"] = "Plantilla KPI Comercial - Cumplimiento por Canal y Consultor (RSM)"
ws["B2"].font = title_font
ws.merge_cells("B2:H2")

ws["B4"] = "Como usar esta plantilla"
ws["B4"].font = section_font

instrucciones = [
    "1. Abra la hoja 'Datos KPI' y llene UNICAMENTE las celdas en AMARILLO (texto azul): "
    "Consultor (RSM), Zona, Order Intake YTD, Ventas (Sales) YTD, Presupuesto (Budget) YTD "
    "Paneles y Backlog actual, para cada canal (Paneles / Otros).",
    "2. Las columnas en GRIS (% Cumplimiento, Order Intake Total, Ventas Total, Backlog "
    "Total, Book-to-Bill y Estado) se calculan solas con formulas: no las edite a mano.",
    "3. Ya se incluyen 5 filas de ejemplo con datos reales de un periodo YTD y 3 filas en "
    "blanco listas para agregar nuevos consultores; si necesita mas filas, inserte filas "
    "nuevas ARRIBA de la fila 'TOTAL GENERAL' y copie el formato/formulas de la fila anterior.",
    "4. La fila 'TOTAL GENERAL' suma automaticamente toda la tabla y calcula el % de "
    "cumplimiento y el Book-to-Bill consolidados.",
    "5. El 'Estado' de cada consultor se calcula asi sobre el % de cumplimiento de "
    "presupuesto (Paneles): CUMPLE >= 100%, EN RIESGO entre 80% y 99%, BAJO < 80%. Los "
    "colores (verde/amarillo/rojo) se aplican solos con formato condicional.",
    "6. La hoja 'Dashboard' muestra el resumen ejecutivo (tarjetas KPI) y 5 graficos que se "
    "actualizan automaticamente al llenar 'Datos KPI': Order Intake vs Ventas vs Presupuesto, "
    "% Cumplimiento por consultor (con meta de 100%), participacion de Ventas por Zona, "
    "Backlog por consultor y Book-to-Bill (con meta de 1.0x).",
    "7. Si agrega consultores en filas nuevas y quiere que aparezcan en los graficos, amplie "
    "el rango de datos de cada grafico en la hoja 'Dashboard' (clic derecho sobre el grafico "
    "> Seleccionar datos).",
    "8. Guarde el archivo con Excel o LibreOffice para que las formulas queden calculadas.",
]
r = 6
for line in instrucciones:
    ws.cell(row=r, column=2, value=line).font = normal_font
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30
    r += 2

ws.merge_cells("B2:H2")
for rr in range(6, r, 2):
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=8)

ws.cell(row=r + 1, column=2, value="Convenciones de color").font = section_font
ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=8)
leyenda = [
    (COLOR_INPUT, "Celda de entrada: dato que usted debe llenar cada periodo."),
    (COLOR_CALC, "Celda calculada: contiene formula, no editar."),
    (COLOR_GREEN, "Cumple la meta (>= 100% del presupuesto)."),
    (COLOR_YELLOW, "En riesgo (entre 80% y 99% del presupuesto)."),
    (COLOR_RED, "Bajo cumplimiento (< 80% del presupuesto)."),
]
rr = r + 3
for color, texto in leyenda:
    c = ws.cell(row=rr, column=2)
    c.fill = PatternFill("solid", fgColor=color)
    c.border = border
    ws.cell(row=rr, column=3, value=texto).font = normal_font
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=8)
    rr += 1

autofit(ws, [3, 16, 14, 14, 14, 14, 14, 14])

# ---------------------------------------------------------------------------
# 2. DATOS KPI
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Datos KPI")
ws.sheet_view.showGridLines = False

ws["B2"] = "Cumplimiento y Performance Comercial por Canal y Consultor (RSM)"
ws["B2"].font = title_font
ws.merge_cells("B2:P2")

ws["B3"] = "Periodo:"
ws["B3"].font = Font(name=FONT_NAME, bold=True)
ws["D3"] = "YTD 2024"
ws["D3"].font = input_font
ws["D3"].fill = input_fill
ws["D3"].border = border
ws.merge_cells("D3:F3")

GROUP_ROW = 5
SUB_ROW = 6
FIRST_DATA_ROW = 7
N_REAL_ROWS = 5
N_BLANK_ROWS = 3
LAST_DATA_ROW = FIRST_DATA_ROW + N_REAL_ROWS + N_BLANK_ROWS - 1  # 14
TOTAL_ROW = LAST_DATA_ROW + 1  # 15

# Columnas: B..P
COL = {
    "rsm": 2, "zona": 3,
    "oi_pan": 4, "oi_otr": 5,
    "vt_pan": 6, "vt_otr": 7,
    "bud_pan": 8, "pct_bud": 9,
    "bl_pan": 10, "bl_otr": 11,
    "oi_tot": 12, "vt_tot": 13, "bl_tot": 14,
    "btb": 15, "estado": 16,
}


def merge_vertical(col, header):
    c = ws.cell(row=GROUP_ROW, column=col, value=header)
    c.font = header_font
    c.fill = header_fill
    c.alignment = center
    c.border = border
    ws.merge_cells(start_row=GROUP_ROW, start_column=col, end_row=SUB_ROW, end_column=col)
    ws.cell(row=SUB_ROW, column=col).border = border
    ws.cell(row=SUB_ROW, column=col).fill = header_fill


def merge_pair(col_start, header, sub_a, sub_b):
    c = ws.cell(row=GROUP_ROW, column=col_start, value=header)
    c.font = header_font
    c.fill = header_fill
    c.alignment = center
    ws.merge_cells(start_row=GROUP_ROW, start_column=col_start, end_row=GROUP_ROW, end_column=col_start + 1)
    for off, sub in enumerate((sub_a, sub_b)):
        sc = ws.cell(row=SUB_ROW, column=col_start + off, value=sub)
        sc.font = header_font
        sc.fill = header_fill
        sc.alignment = center
        sc.border = border
    ws.cell(row=GROUP_ROW, column=col_start).border = border
    ws.cell(row=GROUP_ROW, column=col_start + 1).border = border


merge_vertical(COL["rsm"], "Consultor de Ventas (RSM)")
merge_vertical(COL["zona"], "Zona")
merge_pair(COL["oi_pan"], "Order Intake YTD", "Paneles", "Otros")
merge_pair(COL["vt_pan"], "Ventas (Sales) YTD", "Paneles", "Otros")
merge_vertical(COL["bud_pan"], "Presupuesto (Budget) YTD (Paneles)")
merge_vertical(COL["pct_bud"], "% Cumplimiento Presupuesto (Paneles)")
merge_pair(COL["bl_pan"], "Backlog Actual", "Paneles", "Otros")
merge_vertical(COL["oi_tot"], "Order Intake Total")
merge_vertical(COL["vt_tot"], "Ventas Total")
merge_vertical(COL["bl_tot"], "Backlog Total")
merge_vertical(COL["btb"], "Book-to-Bill\n(Order Intake / Ventas)")
merge_vertical(COL["estado"], "Estado")

ws.row_dimensions[GROUP_ROW].height = 24
ws.row_dimensions[SUB_ROW].height = 26

# --- datos de ejemplo (reales, tomados de un reporte YTD 2024) ---
ejemplo = [
    ("LOZANO TENORIO MARIA CAROLINA", "FRIO", 31115, 13253, 50135, 24389, 50000, 2360, 3815),
    ("URIBE JIMENEZ SANDRA MILENA", "OCCIDENTE", 40564, 4875, 26224, 4800, 33000, 20570, 687),
    ("GUERRERO ESPINOSA JULIO CESAR", "CENTRO", 14375, 1427, 10524, 420, 17000, 5233, 487),
    ("RUGE PEREZ DANIEL FELIPE", "NORTE", 9896, 18701, 6761, 20094, 12000, 3030, 90),
    ("CALDERON CHANTRE DIEGO ALFONSO", "SUR", 25951, 12665, 8664, 7747, 15000, 17174, 5445),
]

INPUT_KEYS = ("rsm", "zona", "oi_pan", "oi_otr", "vt_pan", "vt_otr", "bud_pan", "bl_pan", "bl_otr")
CALC_KEYS = ("pct_bud", "oi_tot", "vt_tot", "bl_tot", "btb", "estado")


def write_row_formulas(row):
    L = {k: f"{get_column_letter(v)}{row}" for k, v in COL.items()}
    ws[L["pct_bud"]] = f'=IFERROR({L["vt_pan"]}/{L["bud_pan"]},0)'
    ws[L["oi_tot"]] = f'={L["oi_pan"]}+{L["oi_otr"]}'
    ws[L["vt_tot"]] = f'={L["vt_pan"]}+{L["vt_otr"]}'
    ws[L["bl_tot"]] = f'={L["bl_pan"]}+{L["bl_otr"]}'
    ws[L["btb"]] = f'=IFERROR({L["oi_tot"]}/{L["vt_tot"]},0)'
    ws[L["estado"]] = (
        f'=IF({L["rsm"]}="","",IF({L["pct_bud"]}>=1,"CUMPLE",'
        f'IF({L["pct_bud"]}>=0.8,"EN RIESGO","BAJO")))'
    )


for i in range(N_REAL_ROWS + N_BLANK_ROWS):
    row = FIRST_DATA_ROW + i
    if i < N_REAL_ROWS:
        rsm, zona, oi_p, oi_o, vt_p, vt_o, bud_p, bl_p, bl_o = ejemplo[i]
        ws.cell(row=row, column=COL["rsm"], value=rsm)
        ws.cell(row=row, column=COL["zona"], value=zona)
        ws.cell(row=row, column=COL["oi_pan"], value=oi_p)
        ws.cell(row=row, column=COL["oi_otr"], value=oi_o)
        ws.cell(row=row, column=COL["vt_pan"], value=vt_p)
        ws.cell(row=row, column=COL["vt_otr"], value=vt_o)
        ws.cell(row=row, column=COL["bud_pan"], value=bud_p)
        ws.cell(row=row, column=COL["bl_pan"], value=bl_p)
        ws.cell(row=row, column=COL["bl_otr"], value=bl_o)
    write_row_formulas(row)

    for key in INPUT_KEYS:
        col = COL[key]
        cell = ws.cell(row=row, column=col)
        cell.font = input_font
        cell.fill = input_fill
        cell.border = border
        cell.alignment = center_nowrap if key != "rsm" else Alignment(vertical="center", wrap_text=True)
    for key in CALC_KEYS:
        col = COL[key]
        cell = ws.cell(row=row, column=col)
        cell.font = calc_font
        cell.fill = calc_fill
        cell.border = border
        cell.alignment = center_nowrap

    for key in ("oi_pan", "oi_otr", "vt_pan", "vt_otr", "bud_pan", "bl_pan", "bl_otr",
                "oi_tot", "vt_tot", "bl_tot"):
        ws.cell(row=row, column=COL[key]).number_format = "#,##0"
    ws.cell(row=row, column=COL["pct_bud"]).number_format = "0%"
    ws.cell(row=row, column=COL["btb"]).number_format = '0.00"x"'

# --- fila de totales ---
def col_range(key):
    c = get_column_letter(COL[key])
    return f"{c}{FIRST_DATA_ROW}:{c}{LAST_DATA_ROW}"


ws.cell(row=TOTAL_ROW, column=COL["rsm"], value="TOTAL GENERAL").font = total_font
ws.merge_cells(start_row=TOTAL_ROW, start_column=COL["rsm"], end_row=TOTAL_ROW, end_column=COL["zona"])
for key in ("oi_pan", "oi_otr", "vt_pan", "vt_otr", "bud_pan", "bl_pan", "bl_otr",
            "oi_tot", "vt_tot", "bl_tot"):
    letter = get_column_letter(COL[key])
    ws.cell(row=TOTAL_ROW, column=COL[key], value=f"=SUM({col_range(key)})")
    ws.cell(row=TOTAL_ROW, column=COL[key]).number_format = "#,##0"
ws.cell(row=TOTAL_ROW, column=COL["pct_bud"],
        value=f'=IFERROR({get_column_letter(COL["vt_pan"])}{TOTAL_ROW}/'
              f'{get_column_letter(COL["bud_pan"])}{TOTAL_ROW},0)')
ws.cell(row=TOTAL_ROW, column=COL["pct_bud"]).number_format = "0%"
ws.cell(row=TOTAL_ROW, column=COL["btb"],
        value=f'=IFERROR({get_column_letter(COL["oi_tot"])}{TOTAL_ROW}/'
              f'{get_column_letter(COL["vt_tot"])}{TOTAL_ROW},0)')
ws.cell(row=TOTAL_ROW, column=COL["btb"]).number_format = '0.00"x"'

for col in range(COL["rsm"], COL["estado"] + 1):
    cell = ws.cell(row=TOTAL_ROW, column=col)
    cell.font = total_font
    cell.fill = section_fill
    cell.border = border
    cell.alignment = center_nowrap

# --- formato condicional (semaforo) sobre % cumplimiento y estado ---
rng_pct = f'{get_column_letter(COL["pct_bud"])}{FIRST_DATA_ROW}:{get_column_letter(COL["pct_bud"])}{TOTAL_ROW}'
rng_estado = f'{get_column_letter(COL["estado"])}{FIRST_DATA_ROW}:{get_column_letter(COL["estado"])}{TOTAL_ROW}'
anchor_pct = f'${get_column_letter(COL["pct_bud"])}{FIRST_DATA_ROW}'

green_fill = PatternFill("solid", fgColor=COLOR_GREEN)
yellow_fill = PatternFill("solid", fgColor=COLOR_YELLOW)
red_fill = PatternFill("solid", fgColor=COLOR_RED)

for rng in (rng_pct, rng_estado):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{anchor_pct}>=1'], fill=green_fill))
    ws.conditional_formatting.add(
        rng, FormulaRule(formula=[f'AND({anchor_pct}>=0.8,{anchor_pct}<1)'], fill=yellow_fill))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{anchor_pct}<0.8'], fill=red_fill))

autofit(ws, [3, 26, 12, 11, 11, 11, 11, 13, 12, 11, 11, 13, 12, 12, 12, 12], start_col=1)
ws.freeze_panes = "D7"

# ---------------------------------------------------------------------------
# 3. DASHBOARD
# ---------------------------------------------------------------------------
ws2 = wb.create_sheet("Dashboard")
ws2.sheet_view.showGridLines = False
ws2["B2"] = "Dashboard KPI Comercial - Cumplimiento y Performance"
ws2["B2"].font = title_font
ws2.merge_cells("B2:M2")
ws2["B3"] = "='Datos KPI'!D3"
ws2["B3"].font = Font(name=FONT_NAME, italic=True, color="808080")
ws2.merge_cells("B3:M3")

# --- tarjetas KPI ejecutivas ---
tarjetas = [
    ("Ventas Total YTD", f"='Datos KPI'!{get_column_letter(COL['vt_tot'])}{TOTAL_ROW}", "#,##0"),
    ("Presupuesto Total (Paneles)", f"='Datos KPI'!{get_column_letter(COL['bud_pan'])}{TOTAL_ROW}", "#,##0"),
    ("% Cumplimiento Global", f"='Datos KPI'!{get_column_letter(COL['pct_bud'])}{TOTAL_ROW}", "0%"),
    ("Backlog Total", f"='Datos KPI'!{get_column_letter(COL['bl_tot'])}{TOTAL_ROW}", "#,##0"),
    ("Book-to-Bill Global", f"='Datos KPI'!{get_column_letter(COL['btb'])}{TOTAL_ROW}", '0.00"x"'),
]
card_col = 2
for titulo, formula, fmt in tarjetas:
    c_lbl = ws2.cell(row=5, column=card_col, value=titulo)
    c_lbl.font = Font(name=FONT_NAME, size=9, bold=True, color=COLOR_HEADER_TEXT)
    c_lbl.fill = header_fill
    c_lbl.alignment = center
    c_val = ws2.cell(row=6, column=card_col, value=formula)
    c_val.font = Font(name=FONT_NAME, size=16, bold=True, color=COLOR_HEADER)
    c_val.number_format = fmt
    c_val.alignment = center_nowrap
    c_val.fill = section_fill
    for r_ in (5, 6):
        ws2.cell(row=r_, column=card_col).border = border
        ws2.cell(row=r_, column=card_col + 1).border = border
    ws2.merge_cells(start_row=5, start_column=card_col, end_row=5, end_column=card_col + 1)
    ws2.merge_cells(start_row=6, start_column=card_col, end_row=6, end_column=card_col + 1)
    card_col += 2

ws2.row_dimensions[5].height = 20
ws2.row_dimensions[6].height = 28
autofit(ws2, [3] + [13] * 12, start_col=1)

CHART_CAT_FIRST = FIRST_DATA_ROW
CHART_CAT_LAST = FIRST_DATA_ROW + N_REAL_ROWS - 1  # solo filas reales (evita barras vacias)

data_sheet_name = "Datos KPI"


def ref(key, first=CHART_CAT_FIRST, last=CHART_CAT_LAST, header_row=SUB_ROW):
    return Reference(ws, min_col=COL[key], min_row=header_row, max_row=last)


cats_rsm = Reference(ws, min_col=COL["rsm"], min_row=CHART_CAT_FIRST, max_row=CHART_CAT_LAST)
cats_zona = Reference(ws, min_col=COL["zona"], min_row=CHART_CAT_FIRST, max_row=CHART_CAT_LAST)

# --- Grafico 1: Order Intake vs Ventas vs Presupuesto (Paneles) por consultor ---
chart1 = BarChart()
chart1.type = "col"
chart1.grouping = "clustered"
chart1.title = "Order Intake vs Ventas vs Presupuesto (Paneles) por Consultor"
chart1.y_axis.title = "Valor"
chart1.x_axis.title = "Consultor (RSM)"
chart1.height, chart1.width = 9, 18
for key, name in ((COL["oi_pan"], "Order Intake"), (COL["vt_pan"], "Ventas"), (COL["bud_pan"], "Presupuesto")):
    data = Reference(ws, min_col=key, min_row=SUB_ROW, max_row=CHART_CAT_LAST)
    chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats_rsm)
ws2.add_chart(chart1, "B9")

# --- Grafico 2: % Cumplimiento por consultor con meta de 100% ---
bar2 = BarChart()
bar2.type = "col"
bar2.title = "% Cumplimiento de Presupuesto (Paneles) por Consultor"
bar2.y_axis.title = "% Cumplimiento"
bar2.x_axis.title = "Consultor (RSM)"
bar2.y_axis.numFmt = "0%"
bar2.height, bar2.width = 9, 18
data_pct = Reference(ws, min_col=COL["pct_bud"], min_row=SUB_ROW, max_row=CHART_CAT_LAST)
bar2.add_data(data_pct, titles_from_data=True)
bar2.set_categories(cats_rsm)
bar2.series[0].dLbls = DataLabelList()
bar2.series[0].dLbls.showVal = True
bar2.series[0].dLbls.numFmt = "0%"

META_PCT_COL = COL["estado"] + 2
line2 = LineChart()
meta_header_cell = ws.cell(row=SUB_ROW, column=META_PCT_COL, value="Meta 100%")
for i in range(N_REAL_ROWS):
    ws.cell(row=CHART_CAT_FIRST + i, column=META_PCT_COL, value=1)
    ws.cell(row=CHART_CAT_FIRST + i, column=META_PCT_COL).number_format = "0%"
data_meta = Reference(ws, min_col=META_PCT_COL, min_row=SUB_ROW, max_row=CHART_CAT_LAST)
line2.add_data(data_meta, titles_from_data=True)
line2.set_categories(cats_rsm)
bar2 += line2
ws2.add_chart(bar2, "B27")

# --- Grafico 3: participacion de Ventas Total por Zona ---
pie3 = PieChart()
pie3.title = "Participacion en Ventas Total YTD por Zona"
pie3.height, pie3.width = 9, 12
data_vt = Reference(ws, min_col=COL["vt_tot"], min_row=SUB_ROW, max_row=CHART_CAT_LAST)
pie3.add_data(data_vt, titles_from_data=True)
pie3.set_categories(cats_zona)
pie3.dataLabels = DataLabelList()
pie3.dataLabels.showPercent = True
ws2.add_chart(pie3, "L9")

# --- Grafico 4: Backlog por consultor (Paneles + Otros, apilado) ---
chart4 = BarChart()
chart4.type = "col"
chart4.grouping = "stacked"
chart4.overlap = 100
chart4.title = "Backlog Actual por Consultor (Paneles + Otros)"
chart4.y_axis.title = "Backlog"
chart4.x_axis.title = "Consultor (RSM)"
chart4.height, chart4.width = 9, 18
for key, name in ((COL["bl_pan"], "Paneles"), (COL["bl_otr"], "Otros")):
    data = Reference(ws, min_col=key, min_row=SUB_ROW, max_row=CHART_CAT_LAST)
    chart4.add_data(data, titles_from_data=True)
chart4.set_categories(cats_rsm)
ws2.add_chart(chart4, "B45")

# --- Grafico 5: Book-to-Bill por consultor con meta de 1.0x ---
bar5 = BarChart()
bar5.type = "col"
bar5.title = "Book-to-Bill (Order Intake / Ventas) por Consultor"
bar5.y_axis.title = "Book-to-Bill"
bar5.x_axis.title = "Consultor (RSM)"
bar5.height, bar5.width = 9, 12
data_btb = Reference(ws, min_col=COL["btb"], min_row=SUB_ROW, max_row=CHART_CAT_LAST)
bar5.add_data(data_btb, titles_from_data=True)
bar5.set_categories(cats_rsm)

line5 = LineChart()
META_BTB_COL = COL["estado"] + 3
ws.cell(row=SUB_ROW, column=META_BTB_COL, value="Meta 1.0x")
for i in range(N_REAL_ROWS):
    ws.cell(row=CHART_CAT_FIRST + i, column=META_BTB_COL, value=1)
    ws.cell(row=CHART_CAT_FIRST + i, column=META_BTB_COL).number_format = '0.00"x"'
data_meta_btb = Reference(ws, min_col=META_BTB_COL, min_row=SUB_ROW, max_row=CHART_CAT_LAST)
line5.add_data(data_meta_btb, titles_from_data=True)
line5.set_categories(cats_rsm)
bar5 += line5
ws2.add_chart(bar5, "L27")

# columnas auxiliares (metas fijas usadas solo por los graficos): se ocultan
ws.cell(row=GROUP_ROW, column=META_PCT_COL, value="Ref. graficos (no editar)").font = Font(
    name=FONT_NAME, size=8, italic=True, color="808080")
ws.column_dimensions[get_column_letter(META_PCT_COL)].hidden = True
ws.column_dimensions[get_column_letter(META_BTB_COL)].hidden = True

wb.save(OUT_PATH)
print(f"Plantilla creada en {OUT_PATH}")
