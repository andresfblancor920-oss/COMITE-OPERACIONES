"""
Crea la plantilla Excel de Ventas / Revenue / Order Intake para el Comité de
Operaciones (vista Customer Excellence), precargada con los datos de
Septiembre (corte W4).

    python scripts/crear_plantilla_ventas.py [ruta_salida.xlsx]

Por defecto escribe plantilla/Plantilla_Ventas_OrderIntake.xlsx.
OJO: sobrescribe el archivo; no lo ejecute sobre una plantilla ya diligenciada.

Convención de colores:
    amarillo = dato de entrada (se diligencia cada semana)
    gris     = fórmula (se calcula sola, no editar)
"""

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NAVY = "0B2545"
INPUT = PatternFill("solid", fgColor="FFF2CC")
CALC = PatternFill("solid", fgColor="EDEFF2")
HEAD = PatternFill("solid", fgColor=NAVY)
SUB = PatternFill("solid", fgColor="DCE6F2")
THIN = Side(style="thin", color="BFC5CC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_HEAD = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
F_TITLE = Font(name="Calibri", bold=True, color=NAVY, size=14)
F_NOTE = Font(name="Calibri", italic=True, color="666666", size=9)
F_BODY = Font(name="Calibri", size=10)
F_BOLD = Font(name="Calibri", size=10, bold=True)

NUM = "#,##0"
PCT = "0%"
COP = '"$" #,##0'

GREEN_F = PatternFill("solid", fgColor="C8E6C9")
AMBER_F = PatternFill("solid", fgColor="FFE8A3")
RED_F = PatternFill("solid", fgColor="F8C9C9")


def titulo(ws, texto, nota):
    ws["A1"] = texto
    ws["A1"].font = F_TITLE
    ws["A2"] = nota
    ws["A2"].font = F_NOTE
    ws.sheet_view.showGridLines = False


def encabezados(ws, row, headers, widths=None):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill, c.font, c.border = HEAD, F_HEAD, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def celda(ws, row, col, value, fill=None, fmt=None, bold=False):
    c = ws.cell(row=row, column=col, value=value)
    c.border = BORDER
    c.font = F_BOLD if bold else F_BODY
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


def semaforo(ws, rango, verde="Config!$B$8", amarillo="Config!$B$9"):
    ws.conditional_formatting.add(rango, CellIsRule(operator="equal", formula=['""'], fill=PatternFill()))
    ws.conditional_formatting.add(rango, CellIsRule(operator="greaterThanOrEqual", formula=[verde], fill=GREEN_F))
    ws.conditional_formatting.add(rango, CellIsRule(operator="greaterThanOrEqual", formula=[amarillo], fill=AMBER_F))
    ws.conditional_formatting.add(rango, CellIsRule(operator="lessThan", formula=[amarillo], fill=RED_F))


# ---------------------------------------------------------------------------
def hoja_instrucciones(wb):
    ws = wb.active
    ws.title = "Instrucciones"
    titulo(ws, "Comité de Operaciones · Ventas, Revenue y Order Intake",
           "Customer Excellence — plantilla semanal que alimenta las 3 diapositivas del comité")
    pasos = [
        "1. Config: actualice mes, semana de corte, fecha del comité y umbrales del semáforo.",
        "2. Ventas_m2: por línea, F'cast inicial/actual, Budget, F'cast semanal (W1-W4), Real/Proyección semanal y Real Sales facturado.",
        "3. Revenue: cantidad F'cast (m², und) y valor en COP por línea; Budget en m² o en COP cuando aplique.",
        "4. Order_Intake: Budget y m² ganados (adjudicados) por semana; Estado = Cerrada / En curso / Proyectada.",
        "5. Pipeline: negocios por ganar con m², probabilidad y semana/mes de cierre esperado.",
        "6. Comentarios: mensajes propios para cada diapositiva (opcional; se suman a las alertas automáticas).",
        "",
        "Celdas AMARILLAS = datos de entrada.  Celdas GRISES = fórmulas (no editar).",
        "Subtotales Premium / Estándar, % de cumplimiento, gaps y precios promedio se calculan solos.",
        "",
        "Generar la presentación:  python scripts/generar_ppt_ventas.py plantilla/Plantilla_Ventas_OrderIntake.xlsx",
        "Cada gráfico del PowerPoint es nativo: clic derecho > Editar datos abre su Excel embebido",
        "(celdas amarillas editables) y el gráfico se actualiza sin macros ni vínculos externos.",
    ]
    for i, p in enumerate(pasos, start=4):
        ws.cell(row=i, column=1, value=p).font = F_BODY
    ws.column_dimensions["A"].width = 120


def hoja_config(wb):
    ws = wb.create_sheet("Config")
    titulo(ws, "Configuración del periodo", "Estos valores alimentan títulos, semáforos y alertas")
    filas = [
        ("Mes", "Septiembre", None),
        ("Año", 2026, "0"),
        ("Semana de corte (1-5)", 4, "0"),
        ("Fecha del comité", "30/09/2026", None),
        ("Área que presenta", "Customer Excellence", None),
        ("Umbral verde (≥)", 0.95, PCT),
        ("Umbral amarillo (≥)", 0.85, PCT),
        ("Alcance Order Intake", "Premium (Env + Refri)", None),
        ("Budget Order Intake mes (m²)", 147719, NUM),
    ]
    for r, (k, v, fmt) in enumerate(filas, start=3):
        celda(ws, r, 1, k, bold=True)
        celda(ws, r, 2, v, INPUT, fmt)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 26


VENTAS = [
    # línea, grupo, fc_ini, fc_act, budget, fc W1-W4, real/proy W1-W4, real sales
    ("Envolventes", "Premium", 103379, 83329, 91536, (25832, 16666, 19999, 20832), (28862, 19733, 17184, 17550), 80178),
    ("Refrigeración", "Premium", 35600, 41293, 43141, (12801, 8259, 9910, 10323), (12610, 5429, 8229, 15025), 33426),
    ("KingCore", "Estándar", 29779, 25931, 27974, (3690, 5186, 10572, 6483), (4460, 1760, 9860, 9129), 25209),
    ("Single Skin", "Estándar", 23851, 4594, 20387, (721, 2800, 0, 1073), (2209, 1042, 0, 1343), 3251),
    ("Ducto", "Estándar", 520, 1044, 1385, (396, 0, 520, 0), (1044, 0, 0, 0), 1044),
    ("Línea Arquitectónica", "Arquitectónica", 718, 718, 3182, (0, 0, 0, 718), (0, 264, 0, 0), 264),
]
VENTAS_FILA0 = 5
VENTAS_MAX = 14  # filas de entrada disponibles (5..18)


def hoja_ventas(wb):
    ws = wb.create_sheet("Ventas_m2")
    titulo(ws, "Volumen de ventas (m²) por línea",
           "F'cast semanal = plan de facturación por semana · Real/Proy = facturado en semanas cerradas y proyección en la semana en curso · Real Sales = facturado acumulado")
    headers = ["Línea", "Grupo", "F'cast Inicial", "F'cast Actual", "Budget",
               "F'cast W1", "F'cast W2", "F'cast W3", "F'cast W4",
               "Real/Proy W1", "Real/Proy W2", "Real/Proy W3", "Real/Proy W4", "Real Sales (facturado)",
               "Cierre proyectado", "% F'cast vs Budget", "% Real vs Budget", "Gap Real vs Budget",
               "Cumpl. W1", "Cumpl. W2", "Cumpl. W3", "Cumpl. W4"]
    widths = [22, 14] + [11] * 12 + [12, 11, 11, 12, 9, 9, 9, 9]
    encabezados(ws, VENTAS_FILA0 - 1, headers, widths)
    dv = DataValidation(type="list", formula1='"Premium,Estándar,Arquitectónica"', allow_blank=True)
    ws.add_data_validation(dv)
    last = VENTAS_FILA0 + VENTAS_MAX - 1
    for i in range(VENTAS_MAX):
        r = VENTAS_FILA0 + i
        d = VENTAS[i] if i < len(VENTAS) else None
        vals = ([d[0], d[1], d[2], d[3], d[4], *d[5], *d[6], d[7]] if d else [None] * 14)
        for c, v in enumerate(vals, start=1):
            celda(ws, r, c, v, INPUT, NUM if c > 2 else None)
        dv.add(ws.cell(row=r, column=2))
        f = {
            15: f'=IF(A{r}="","",SUM(J{r}:M{r}))',
            16: f'=IF(OR(A{r}="",E{r}=0),"",D{r}/E{r})',
            17: f'=IF(OR(A{r}="",E{r}=0),"",N{r}/E{r})',
            18: f'=IF(A{r}="","",N{r}-E{r})',
        }
        for k, col in enumerate("FGHI"):
            real = "JKLM"[k]
            f[19 + k] = f'=IF(OR(A{r}="",{col}{r}=0),"",{real}{r}/{col}{r})'
        for c, formula in f.items():
            celda(ws, r, c, formula, CALC, PCT if c in (16, 17) or c >= 19 else NUM)

    # Subtotales por grupo
    r0 = last + 2
    for j, grupo in enumerate(["Premium", "Estándar", "Arquitectónica", "TOTAL"]):
        r = r0 + j
        celda(ws, r, 1, grupo if grupo == "TOTAL" else f"Subtotal {grupo}", SUB, bold=True)
        celda(ws, r, 2, grupo, SUB, bold=True)
        for c in range(3, 16):
            L = get_column_letter(c)
            rango = f"{L}${VENTAS_FILA0}:{L}${last}"
            formula = f"=SUM({rango})" if grupo == "TOTAL" else f'=SUMIF($B${VENTAS_FILA0}:$B${last},$B{r},{rango})'
            celda(ws, r, c, formula, SUB, NUM, bold=True)
        celda(ws, r, 16, f'=IF(E{r}=0,"",D{r}/E{r})', SUB, PCT, bold=True)
        celda(ws, r, 17, f'=IF(E{r}=0,"",N{r}/E{r})', SUB, PCT, bold=True)
        celda(ws, r, 18, f"=N{r}-E{r}", SUB, NUM, bold=True)
        for k, col in enumerate("FGHI"):
            real = "JKLM"[k]
            celda(ws, r, 19 + k, f'=IF({col}{r}=0,"",{real}{r}/{col}{r})', SUB, PCT, bold=True)
    semaforo(ws, f"P{VENTAS_FILA0}:V{r0 + 3}")
    ws.freeze_panes = ws.cell(row=VENTAS_FILA0, column=3)


REVENUE = [
    # línea, grupo, unidad, cantidad, COP, budget cantidad, budget COP
    ("Envolventes", "Premium", "m²", 83329, 9169606489, 91536, None),
    ("Refrigeración", "Premium", "m²", 41293, 4478225850, 43141, None),
    ("KingCore", "Estándar", "m²", 25931, 836015895, 27974, None),
    ("Single Skin", "Estándar", "m²", 4594, 103998996, 20387, None),
    ("Ducto", "Estándar", "m²", 1044, 52372800, 1385, None),
    ("Traslúcida", "Estándar", "m²", 2155, 573520000, None, None),
    ("Línea Arquitectónica", "Arquitectónica", "m²", 264, 20500000, 3182, None),
    ("Proyectos", "Proyectos", "—", None, 328000000, None, 533000000),
    ("Fijaciones y Accesorios", "Accesorios", "—", None, 1007000000, None, None),
    ("Consumibles", "Accesorios", "—", None, 0, None, None),
    ("Puertas", "Puertas", "und", 17, 83050000, None, None),
    ("Facturación ICO", "ICO", "—", None, 1251189846, None, None),
]
REV_FILA0 = 5
REV_MAX = 16


def hoja_revenue(wb):
    ws = wb.create_sheet("Revenue")
    titulo(ws, "Revenue (COP) por línea — F'cast del mes",
           "Precio promedio = Revenue / cantidad · % Performance contra Budget en cantidad o, si no hay, contra Budget en COP")
    headers = ["Línea", "Grupo", "Unidad", "Cantidad F'cast", "Revenue F'cast (COP)",
               "Budget cantidad", "Budget Revenue (COP)", "Precio promedio", "% Performance", "% Mix Revenue"]
    encabezados(ws, REV_FILA0 - 1, headers, [24, 16, 8, 14, 20, 14, 20, 14, 13, 12])
    dv = DataValidation(type="list", formula1='"Premium,Estándar,Arquitectónica,Proyectos,Accesorios,Puertas,ICO,Otros"', allow_blank=True)
    ws.add_data_validation(dv)
    last = REV_FILA0 + REV_MAX - 1
    for i in range(REV_MAX):
        r = REV_FILA0 + i
        d = REVENUE[i] if i < len(REVENUE) else (None,) * 7
        for c, v in enumerate(d, start=1):
            celda(ws, r, c, v, INPUT, COP if c in (5, 7) else (NUM if c in (4, 6) else None))
        dv.add(ws.cell(row=r, column=2))
        celda(ws, r, 8, f'=IF(OR(A{r}="",N(D{r})=0),"",E{r}/D{r})', CALC, COP)
        celda(ws, r, 9, f'=IF(A{r}="","",IF(N(F{r})>0,D{r}/F{r},IF(N(G{r})>0,E{r}/G{r},"")))', CALC, PCT)
        celda(ws, r, 10, f'=IF(A{r}="","",E{r}/$E${last + 2 + 7})', CALC, PCT)
    r0 = last + 2
    grupos = ["Premium", "Estándar", "Arquitectónica", "Proyectos", "Accesorios", "Puertas", "ICO"]
    for j, g in enumerate(grupos + ["TOTAL"]):
        r = r0 + j
        celda(ws, r, 1, "TOTAL F'cast (incl. ICO)" if g == "TOTAL" else f"Subtotal {g}", SUB, bold=True)
        celda(ws, r, 2, g, SUB, bold=True)
        celda(ws, r, 3, "", SUB)
        for c in (4, 5, 6, 7):
            L = get_column_letter(c)
            rango = f"{L}${REV_FILA0}:{L}${last}"
            formula = f"=SUM({rango})" if g == "TOTAL" else f'=SUMIF($B${REV_FILA0}:$B${last},$B{r},{rango})'
            if g == "TOTAL" and c in (4, 6):
                formula = ""  # no se suman m² con unidades
            celda(ws, r, c, formula, SUB, COP if c in (5, 7) else NUM, bold=True)
        celda(ws, r, 8, "" if g == "TOTAL" else f'=IF(D{r}=0,"",E{r}/D{r})', SUB, COP, bold=True)
        # % performance del subtotal: solo sobre las líneas que tienen budget en cantidad
        celda(ws, r, 9, (f'=IF(F{r}>0,SUMIFS($D${REV_FILA0}:$D${last},$B${REV_FILA0}:$B${last},$B{r},$F${REV_FILA0}:$F${last},">0")/F{r},'
                         f'IF(G{r}>0,E{r}/G{r},""))') if g != "TOTAL" else "", SUB, PCT, bold=True)
        celda(ws, r, 10, f"=E{r}/$E${r0 + len(grupos)}", SUB, PCT, bold=True)
    semaforo(ws, f"I{REV_FILA0}:I{r0 + 6}")
    ws.freeze_panes = ws.cell(row=REV_FILA0, column=2)


OI = [("W1", 36930, 25901, "Cerrada"), ("W2", 36930, 41491, "Cerrada"),
      ("W3", 36930, 45187, "Cerrada"), ("W4", 36929, 27588, "En curso"), ("W5", None, None, None)]


def hoja_order_intake(wb):
    ws = wb.create_sheet("Order_Intake")
    titulo(ws, "Order Intake semanal — Budget vs negocios ganados (m² adjudicados)",
           "Budget semanal sugerido = Budget mes / N° semanas (Config). Estado: Cerrada / En curso / Proyectada")
    headers = ["Semana", "Budget (m²)", "Real ganado (m²)", "Estado", "Gap (m²)", "Cumplimiento",
               "Budget acumulado", "Real acumulado", "Cumpl. acumulado"]
    encabezados(ws, 4, headers, [12, 14, 16, 13, 12, 13, 16, 15, 15])
    dv = DataValidation(type="list", formula1='"Cerrada,En curso,Proyectada"', allow_blank=True)
    ws.add_data_validation(dv)
    for i, (w, b, real, est) in enumerate(OI):
        r = 5 + i
        celda(ws, r, 1, w, INPUT)
        celda(ws, r, 2, b, INPUT, NUM)
        celda(ws, r, 3, real, INPUT, NUM)
        celda(ws, r, 4, est, INPUT)
        dv.add(ws.cell(row=r, column=4))
        celda(ws, r, 5, f'=IF(OR(B{r}="",C{r}=""),"",C{r}-B{r})', CALC, NUM)
        celda(ws, r, 6, f'=IF(OR(N(B{r})=0,C{r}=""),"",C{r}/B{r})', CALC, PCT)
        celda(ws, r, 7, f'=IF(B{r}="","",SUM($B$5:B{r}))', CALC, NUM)
        celda(ws, r, 8, f'=IF(C{r}="","",SUM($C$5:C{r}))', CALC, NUM)
        celda(ws, r, 9, f'=IF(OR(G{r}="",H{r}=""),"",H{r}/G{r})', CALC, PCT)
    r = 10
    celda(ws, r, 1, "TOTAL", SUB, bold=True)
    celda(ws, r, 2, "=SUM(B5:B9)", SUB, NUM, bold=True)
    celda(ws, r, 3, "=SUM(C5:C9)", SUB, NUM, bold=True)
    celda(ws, r, 4, "", SUB)
    celda(ws, r, 5, "=C10-B10", SUB, NUM, bold=True)
    celda(ws, r, 6, '=IF(B10=0,"",C10/B10)', SUB, PCT, bold=True)
    semaforo(ws, "F5:F10")
    semaforo(ws, "I5:I9")
    ws["A12"] = "Control: Budget del mes en Config"
    ws["A12"].font = F_NOTE
    ws["C12"] = "=Config!B11"
    ws["C12"].number_format = NUM
    ws["D12"] = '=IF(B10=Config!B11,"OK","Revise: la suma semanal difiere del budget del mes")'
    ws["D12"].font = F_NOTE


PIPELINE = [
    ("Oportunidad A (ejemplo)", "Cliente por definir", "Envolventes", 6500, 0.7, "W4", "Comercial"),
    ("Oportunidad B (ejemplo)", "Cliente por definir", "Refrigeración", 4200, 0.5, "W4", "Comercial"),
    ("Oportunidad C (ejemplo)", "Cliente por definir", "Envolventes", 12000, 0.6, "Mes siguiente", "Comercial"),
    ("Oportunidad D (ejemplo)", "Cliente por definir", "Refrigeración", 8000, 0.4, "Mes siguiente", "Comercial"),
]


def hoja_pipeline(wb):
    ws = wb.create_sheet("Pipeline")
    titulo(ws, "Pipeline — negocios por ganar",
           "Filas de EJEMPLO: reemplácelas por las oportunidades reales. Ponderado = m² × probabilidad. "
           "Cierre esperado: W1..W5 del mes en curso o 'Mes siguiente'")
    headers = ["Oportunidad / Proyecto", "Cliente", "Línea", "m²", "Probabilidad", "Cierre esperado",
               "Responsable", "m² ponderados"]
    encabezados(ws, 4, headers, [28, 24, 16, 12, 12, 16, 16, 15])
    dv = DataValidation(type="list", formula1='"W1,W2,W3,W4,W5,Mes siguiente"', allow_blank=True)
    ws.add_data_validation(dv)
    for i in range(25):
        r = 5 + i
        d = PIPELINE[i] if i < len(PIPELINE) else (None,) * 7
        for c, v in enumerate(d, start=1):
            celda(ws, r, c, v, INPUT, NUM if c == 4 else (PCT if c == 5 else None))
        dv.add(ws.cell(row=r, column=6))
        celda(ws, r, 8, f'=IF(A{r}="","",D{r}*E{r})', CALC, NUM)
    celda(ws, 31, 1, "TOTAL", SUB, bold=True)
    celda(ws, 31, 4, "=SUM(D5:D29)", SUB, NUM, bold=True)
    celda(ws, 31, 8, "=SUM(H5:H29)", SUB, NUM, bold=True)


def hoja_comentarios(wb):
    ws = wb.create_sheet("Comentarios")
    titulo(ws, "Comentarios para cada diapositiva (opcional)",
           "Hasta 3 mensajes por diapositiva; aparecen antes de las alertas automáticas")
    encabezados(ws, 4, ["Diapositiva", "Comentario"], [18, 110])
    for i, s in enumerate(["Ventas m2"] * 3 + ["Revenue"] * 3 + ["Order Intake"] * 3):
        celda(ws, 5 + i, 1, s, CALC, bold=True)
        celda(ws, 5 + i, 2, None, INPUT)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "plantilla" / "Plantilla_Ventas_OrderIntake.xlsx"
    wb = Workbook()
    hoja_instrucciones(wb)
    hoja_config(wb)
    hoja_ventas(wb)
    hoja_revenue(wb)
    hoja_order_intake(wb)
    hoja_pipeline(wb)
    hoja_comentarios(wb)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Plantilla creada: {out}")


if __name__ == "__main__":
    main()
