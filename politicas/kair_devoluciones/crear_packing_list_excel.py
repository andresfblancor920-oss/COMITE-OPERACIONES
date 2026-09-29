"""Genera la plantilla editable de la Lista Detallada de Empaque (Packing List) KAIR.

Sigue el "Formato Producto Terminado Despachado" de logística e incluye la hoja
de verificación que cruza la lista detallada contra la remisión del despacho.

Uso: python politicas/kair_devoluciones/crear_packing_list_excel.py
Crea Plantilla_Packing_List_KAIR.xlsx junto a este script (lo sobrescribe).

Hojas:
  Lista Detallada / Verificación                   -> en blanco, para cada despacho
  Ejemplo - Lista Detallada / Ejemplo - Verificación -> módulo habitacional muestra
                                                      (remisión 22390, sin datos personales)
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

NAVY, SAGE, SAGE_LIGHT, INPUT, GREY = "0E2F56", "8FAF8C", "DDE8DA", "FFF6D5", "E6EAEE"
thin = Side(style="thin", color="C5D0DC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
PIE = "Política de Devolución de Producto KAIR - Versión 1 - 29-09-2026"
N_LINEAS = 40
N_REFS = 20

# Ejemplo: módulo habitacional muestra (remisión 22390 / lista detallada del 22-09-2026)
EJ_ENCABEZADO = {
    "Fecha": "22/09/2026", "N° Remisión": "22390", "P.V. Nro.": "PV-00012991",
    "Cliente": "KINGSPAN PANELES AISLADOS SAS", "Dirección de entrega": "Hotel Hilton, Carrera 1 #62 · 07 KM",
    "Ciudad": "Cartagena", "Contacto en obra": "(dato personal omitido)", "Teléfono": "(dato personal omitido)",
    "Transportadora": "PASAMAR", "Placa": "TAL758", "Conductor": "(dato personal omitido)", "Tipo de vehículo": "",
}
EJ_LINEAS = [  # código, descripción, U.M., paquete, longitud mm, ancho útil m, cantidad
    ("3110", "KINGFRIGO PIR100 CAL28-9002/CAL28-9002", "M2", "8003", 3000, 1, 8),
    ("3110", "KINGFRIGO PIR100 CAL28-9002/CAL28-9002", "M2", "8003", 2700, 1, 2),
    ("3110", "KINGFRIGO PIR100 CAL28-9002/CAL28-9002", "M2", "8003", 2730, 1, 2),
    ("6099", "KINGROOF PIR50 CAL26-7043/CAL28-MADCLARO", "M2", "", 4750, 1, 3),
    ("315", "ESQUINERO EXTERIOR", "M", "", 3000, None, 4),
    ("316", "ESQUINERO INTERIOR", "M", "", 3000, None, 12),
    ("308", "REMATE FRONTAL 18/30", "M", "", 3000, None, 2),
    ("4215", "REMATE DISEÑO ESPECIAL - 7043", "M", "", 3000, None, 3),
    ("320", "REMATE DISEÑO ESPECIAL - 9002", "M", "", 2400, None, 16),
    ("5626", "KIT DE ANCLAJE FACHADA 100", "UND", "", None, None, 50),
]
EJ_REMISION = [  # código, descripción, U.M., lote, cantidad despachada
    ("3110", "KINGFRIGO PIR100 CAL28-9002/CAL28-9002", "M2", "WOP18757-260922", 34.86),
    ("6099", "KINGROOF PIR50 CAL26-7043/CAL28-MADCLARO", "M2", "WOP17099-260205", 14.25),
    ("315", "ESQUINERO EXTERIOR", "M", "WAC15445-260922", 12),
    ("316", "ESQUINERO INTERIOR", "M", "WAC15445-260922", 36),
    ("308", "REMATE FRONTAL 18/30", "M", "WAC15357-260702", 6),
    ("4215", "REMATE DISEÑO ESPECIAL - 7043", "M", "WAC15444-260922", 9),
    ("320", "REMATE DISEÑO ESPECIAL - 9002", "M", "WAC15446-260922", 38.4),
    ("5626", "KIT DE ANCLAJE FACHADA 100", "UND", "", 50),
]
CHEQUEO = [
    "La placa y el conductor coinciden con el correo de transporte",
    "Se recibieron todos los paquetes relacionados en la lista detallada",
    "El material no presenta golpes, rayones, humedad ni deformaciones",
    "Se recibió el registro fotográfico del despacho",
    "Remisión y lista detallada firmadas como recibido (nombre, cédula, teléfono)",
]


def titulo(ws, texto, ultima_col):
    ws.merge_cells(f"A1:{ultima_col}1")
    ws["A1"] = texto
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    ws.merge_cells(f"A2:{ultima_col}2")
    ws["A2"] = PIE + " · Celdas amarillas: diligenciar"
    ws["A2"].font = Font(italic=True, size=9, color="50627A")


def cabecera_tabla(ws, fila, cabeceras):
    for c, t in enumerate(cabeceras, start=1):
        cell = ws.cell(row=fila, column=c, value=t)
        cell.font = Font(bold=True, color="FFFFFF", size=9)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[fila].height = 30


def celda_input(cell):
    cell.fill = PatternFill("solid", fgColor=INPUT)
    cell.border = BORDER
    cell.font = Font(size=9)


def hoja_lista(wb, nombre, encabezado=None, lineas=()):
    ws = wb.create_sheet(nombre)
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJ", [9, 42, 7, 10, 12, 10, 10, 11, 11, 28]):
        ws.column_dimensions[col].width = w
    titulo(ws, "KAIR by Kingspan · LISTA DETALLADA DE EMPAQUE (PACKING LIST)", "J")

    campos = list(EJ_ENCABEZADO)  # mismos campos, en el mismo orden
    for i, campo in enumerate(campos):
        r = 4 + i // 2
        lab, val_ini, val_fin = ("A", "B", "B") if i % 2 == 0 else ("D", "E", "J")
        if lab == "D":
            ws.merge_cells(f"C{r}:D{r}")
            lab = "C"
        ws[f"{lab}{r}"] = campo
        ws[f"{lab}{r}"].font = Font(bold=True, size=9, color=NAVY)
        if val_ini != val_fin:
            ws.merge_cells(f"{val_ini}{r}:{val_fin}{r}")
        celda_input(ws[f"{val_ini}{r}"])
        if encabezado:
            ws[f"{val_ini}{r}"] = encabezado.get(campo, "")
        ws.row_dimensions[r].height = 18

    h = 11
    cabecera_tabla(ws, h, ["Ítem (código)", "Descripción", "U.M.", "N° Paquete", "Longitud (mm)",
                           "Ancho útil (m)", "Cantidad (und)", "Total", "Recibido (Sí/No)", "Observación"])
    dv_um = DataValidation(type="list", formula1='"M2,M,UND"', allow_blank=True)
    dv_si = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
    ws.add_data_validation(dv_um)
    ws.add_data_validation(dv_si)
    primera, ultima = h + 1, h + N_LINEAS
    for i in range(N_LINEAS):
        r = primera + i
        for c in range(1, 11):
            if c != 8:
                celda_input(ws.cell(row=r, column=c))
        if i < len(lineas):
            for c, v in enumerate(lineas[i], start=1):
                if v is not None:
                    ws.cell(row=r, column=c, value=v)
        # Total: m² = longitud × ancho útil (1 m por defecto) × cantidad; m = longitud × cantidad; und = cantidad
        ws[f"H{r}"] = (f'=IF(G{r}="","",IF(C{r}="UND",G{r},IF(C{r}="M2",E{r}/1000*IF(F{r}="",1,F{r})*G{r},'
                       f'E{r}/1000*G{r})))')
        ws[f"H{r}"].number_format = "#,##0.00"
        ws[f"H{r}"].border = BORDER
        ws[f"H{r}"].font = Font(bold=True, size=9, color=NAVY)
        ws[f"H{r}"].fill = PatternFill("solid", fgColor=GREY)
        dv_um.add(f"C{r}")
        dv_si.add(f"I{r}")

    r = ultima + 2
    ws[f"A{r}"] = "OBSERVACIONES"
    ws[f"A{r}"].font = Font(bold=True, size=9, color=NAVY)
    ws.merge_cells(f"A{r + 1}:J{r + 3}")
    celda_input(ws[f"A{r + 1}"])
    ws[f"A{r + 1}"].alignment = Alignment(wrap_text=True, vertical="top")

    r += 5
    for col_lab, col_val, filas in [("A", "B", ["Elaborado por", "Cargo", "Firma y sello"]),
                                    ("E", "G", ["Recibido por", "Cédula", "Teléfono", "Firma y sello"])]:
        for k, t in enumerate(filas):
            ws[f"{col_lab}{r + k}"] = t
            ws[f"{col_lab}{r + k}"].font = Font(bold=True, size=9, color=NAVY)
            fin = "D" if col_val == "B" else "J"
            ws.merge_cells(f"{col_val}{r + k}:{fin}{r + k}")
            celda_input(ws[f"{col_val}{r + k}"])
    if encabezado:
        ws[f"B{r + 1}"] = "Supervisor de logística"

    ws.freeze_panes = f"A{primera}"
    ws.print_title_rows = f"{h}:{h}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws, primera, ultima


def hoja_verificacion(wb, nombre, hoja_lista_nombre, primera, ultima, remision=()):
    ws = wb.create_sheet(nombre)
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", [9, 42, 7, 18, 14, 14, 12, 12, 30]):
        ws.column_dimensions[col].width = w
    titulo(ws, "KAIR by Kingspan · VERIFICACIÓN EN LA ENTREGA (REMISIÓN vs. LISTA DETALLADA)", "I")
    ws.merge_cells("A3:I3")
    ws["A3"] = ("Registre cada referencia de la remisión con su cantidad despachada. El total de la lista detallada "
                "se calcula solo; si la diferencia no es cero, el estado muestra NOVEDAD (numerales 5.5 y 4.1).")
    ws["A3"].font = Font(size=9, color=NAVY)
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 28

    h = 5
    cabecera_tabla(ws, h, ["Código", "Descripción", "U.M.", "Lote (remisión)", "Cant. despachada (remisión)",
                           "Total lista detallada", "Diferencia", "Estado", "Observación"])
    ref = f"'{hoja_lista_nombre}'"
    p0, p1 = h + 1, h + N_REFS
    for i in range(N_REFS):
        r = p0 + i
        for c in (1, 2, 3, 4, 5, 9):
            celda_input(ws.cell(row=r, column=c))
        if i < len(remision):
            for c, v in zip((1, 2, 3, 4, 5), remision[i]):
                ws.cell(row=r, column=c, value=v)
        ws[f"F{r}"] = f'=IF(A{r}="","",SUMIF({ref}!$A${primera}:$A${ultima},A{r},{ref}!$H${primera}:$H${ultima}))'
        ws[f"G{r}"] = f'=IF(OR(A{r}="",E{r}=""),"",ROUND(F{r}-E{r},3))'
        ws[f"H{r}"] = f'=IF(G{r}="","",IF(ABS(G{r})<0.005,"OK","NOVEDAD"))'
        for c in "EFG":
            ws[f"{c}{r}"].number_format = "#,##0.00"
        for c in "FGH":
            ws[f"{c}{r}"].border = BORDER
            ws[f"{c}{r}"].font = Font(bold=True, size=9, color=NAVY)
            ws[f"{c}{r}"].fill = PatternFill("solid", fgColor=GREY)
        ws[f"H{r}"].alignment = Alignment(horizontal="center")
    ws.conditional_formatting.add(f"H{p0}:H{p1}", CellIsRule(operator="equal", formula=['"OK"'],
                                  fill=PatternFill("solid", fgColor="CFE3CC"), font=Font(bold=True, color="2F6B2C")))
    ws.conditional_formatting.add(f"H{p0}:H{p1}", CellIsRule(operator="equal", formula=['"NOVEDAD"'],
                                  fill=PatternFill("solid", fgColor="F6D3CC"), font=Font(bold=True, color="A33A24")))

    t = p1 + 1
    ws.merge_cells(f"A{t}:D{t}")
    ws[f"A{t}"] = "TOTAL"
    ws[f"A{t}"].alignment = Alignment(horizontal="right")
    ws[f"E{t}"] = f"=SUM(E{p0}:E{p1})"
    ws[f"F{t}"] = f"=SUM(F{p0}:F{p1})"
    ws[f"H{t}"] = f'=COUNTIF(H{p0}:H{p1},"NOVEDAD")&" novedad(es)"'
    for c in range(1, 10):
        cell = ws.cell(row=t, column=c)
        cell.fill = PatternFill("solid", fgColor=SAGE_LIGHT)
        cell.font = Font(bold=True, size=9, color=NAVY)
        cell.border = BORDER
    for c in "EF":
        ws[f"{c}{t}"].number_format = "#,##0.00"

    r = t + 2
    ws.merge_cells(f"A{r}:I{r}")
    ws[f"A{r}"] = "LISTA DE CHEQUEO EN LA ENTREGA"
    ws[f"A{r}"].font = Font(bold=True, color="FFFFFF")
    ws[f"A{r}"].fill = PatternFill("solid", fgColor=SAGE)
    dv_si = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
    ws.add_data_validation(dv_si)
    for k, texto in enumerate(CHEQUEO, start=1):
        ws.merge_cells(f"A{r + k}:G{r + k}")
        ws[f"A{r + k}"] = f"{k}. {texto}"
        ws[f"A{r + k}"].font = Font(size=9, color=NAVY)
        celda_input(ws[f"H{r + k}"])
        dv_si.add(f"H{r + k}")
    r += len(CHEQUEO) + 2
    ws.merge_cells(f"A{r}:I{r + 1}")
    ws[f"A{r}"] = ("Reporte faltantes, productos incorrectos o novedades a devolucioneskair@kingspan.com dentro de los "
                   "5 días hábiles siguientes a la recepción, indicando N° de remisión, código, lote, N° de paquete y "
                   "fotos (numerales 1.1 y 4.1).")
    ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="center")
    ws[f"A{r}"].font = Font(size=9, color=NAVY)
    ws[f"A{r}"].fill = PatternFill("solid", fgColor=SAGE_LIGHT)
    ws.freeze_panes = f"A{p0}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)
    _, a, b = hoja_lista(wb, "Lista Detallada")
    hoja_verificacion(wb, "Verificación", "Lista Detallada", a, b)
    _, a, b = hoja_lista(wb, "Ejemplo - Lista Detallada", EJ_ENCABEZADO, EJ_LINEAS)
    hoja_verificacion(wb, "Ejemplo - Verificación", "Ejemplo - Lista Detallada", a, b, EJ_REMISION)
    for ws in wb.worksheets[2:]:
        ws.sheet_properties.tabColor = SAGE
    salida = Path(__file__).with_name("Plantilla_Packing_List_KAIR.xlsx")
    wb.save(salida)
    print(f"Plantilla generada: {salida.name}")


if __name__ == "__main__":
    main()
