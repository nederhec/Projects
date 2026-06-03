"""
Dashboard Generator - Cargos GBS
Gera Dashboard_Cargos_GBS.xlsx com fórmulas em PT-BR.
Aba "Dados" contém os dados brutos; todas as outras abas referenciam via fórmulas.
"""

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

SOURCE = '/root/.claude/uploads/39518091-0fb9-4665-a41a-ed8e752c52fd/88b7cf84-Cargos_GBS.xlsx'
OUTPUT = '/home/user/Projects/Dashboard_Cargos_GBS.xlsx'

# Mapeamento de colunas na aba Dados (letra)
# A=EMPRESA, B=EMPRESA(SAP), C=CÓDIGO CARGO, D=CARGO, E=GRADE, F=SITUAÇÃO,
# G=GRUPO, H=FAMÍLIA GBS, I=SUB FAMÍLIA, J=GRUPO HIERARQUICO, K=NÍVEL,
# L=CBO, M=REGRA CONTROLE JORNADA, N=CONTROLE JORNADA, O=COTA APRENDIZ,
# P=TIPO CONTRATO, Q=SINDICATO, R=MERCADO SELECIONADO, S=DESCRIÇÃO SUMÁRIA,
# T=Cod. Mercer 2025, U=Cod. Mercer 2026
LAST_ROW = 195  # linha 2 a 195 = 194 registros

# ── Paleta ───────────────────────────────────────────────────────────────────
DARK_BLUE  = "1F3864"
MID_BLUE   = "2E75B6"
LIGHT_BLUE = "BDD7EE"
ALT_ROW    = "DEEAF1"
WHITE      = "FFFFFF"
DARK_GREY  = "595959"
GREEN_FILL = "C6EFCE"
YELLOW_FILL= "FFEB9C"

def solid(h):
    return PatternFill("solid", fgColor=h)

def border():
    s = Side(style='thin', color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def hfont(size=11):
    return Font(bold=True, color=WHITE, size=size)

def nfont(bold=False, size=10):
    return Font(bold=bold, size=size, color="000000")

def center():
    return Alignment(horizontal='center', vertical='center')

def left():
    return Alignment(horizontal='left', vertical='center')

def auto_width(ws, min_w=8, max_w=50):
    for col in ws.columns:
        ml = max((len(str(c.value)) for c in col if c.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_w, max(min_w, ml + 2))

def style_row(ws, row, ncols, bg, bold=False, font_color=None):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = solid(bg)
        cell.font = Font(bold=bold, size=10,
                         color=font_color if font_color else ("FFFFFF" if bg == DARK_BLUE else "000000"))
        cell.alignment = center()
        cell.border = border()

def title_row(ws, text, ncols, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1)
    c.value = text
    c.fill = solid(DARK_BLUE)
    c.font = Font(bold=True, color=WHITE, size=14)
    c.alignment = center()
    ws.row_dimensions[row].height = 30

def header_cells(ws, row, headers, bg=MID_BLUE):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c)
        cell.value = h
        cell.fill = solid(bg)
        cell.font = hfont(11)
        cell.alignment = center()
        cell.border = border()

# ── Carregar dados para copiar na aba Dados ───────────────────────────────────
df = pd.read_excel(SOURCE)

wb = Workbook()
wb.remove(wb.active)

# ════════════════════════════════════════════════════════════════════════════
# ABA 0 – Dados (dados brutos, referenciados pelas fórmulas)
# ════════════════════════════════════════════════════════════════════════════
ws0 = wb.create_sheet("Dados")

# Cabeçalho
for c, col in enumerate(df.columns, 1):
    cell = ws0.cell(row=1, column=c)
    cell.value = col
    cell.fill = solid(DARK_BLUE)
    cell.font = hfont(10)
    cell.alignment = center()
    cell.border = border()

# Dados
for r_i, row_data in df.iterrows():
    for c_i, val in enumerate(row_data, 1):
        cell = ws0.cell(row=r_i + 2, column=c_i)
        cell.value = val
        cell.border = border()
        cell.font = nfont(size=9)
        cell.alignment = left()

auto_width(ws0, max_w=40)


# ════════════════════════════════════════════════════════════════════════════
# ABA 1 – KPIs Executivos  (fórmulas PT-BR)
# ════════════════════════════════════════════════════════════════════════════
ws1 = wb.create_sheet("KPIs Executivos")
ws1.sheet_view.showGridLines = False

ws1.merge_cells('A1:F1')
c = ws1['A1']
c.value = "DASHBOARD – CARGOS GBS"
c.fill = solid(DARK_BLUE)
c.font = Font(bold=True, color=WHITE, size=16)
c.alignment = center()
ws1.row_dimensions[1].height = 36

ws1.merge_cells('A2:F2')
c = ws1['A2']
c.value = "Visão Executiva de KPIs"
c.fill = solid(MID_BLUE)
c.font = Font(bold=True, color=WHITE, size=12)
c.alignment = center()
ws1.row_dimensions[2].height = 24

def kpi_box(ws, row, col, label, formula_value, formula_pct=None, bg=LIGHT_BLUE):
    ws.merge_cells(start_row=row,   start_column=col, end_row=row,   end_column=col+1)
    ws.merge_cells(start_row=row+1, start_column=col, end_row=row+1, end_column=col+1)
    lbl = ws.cell(row=row, column=col)
    lbl.value = label
    lbl.fill = solid(DARK_BLUE)
    lbl.font = Font(bold=True, color=WHITE, size=10)
    lbl.alignment = center()
    val = ws.cell(row=row+1, column=col)
    val.value = formula_value
    val.fill = solid(bg)
    val.font = Font(bold=True, color=DARK_BLUE, size=14)
    val.alignment = center()
    if formula_pct is not None:
        ws.merge_cells(start_row=row+2, start_column=col, end_row=row+2, end_column=col+1)
        sv = ws.cell(row=row+2, column=col)
        sv.value = formula_pct
        sv.fill = solid(bg)
        sv.font = Font(color=DARK_GREY, size=9)
        sv.alignment = center()
        sv.number_format = '0,0%'

# Linha 4-6
kpi_box(ws1, 4, 1, "Total de Cargos",
        f'=CONT.VALORES(Dados!A2:A{LAST_ROW})',
        None, LIGHT_BLUE)

kpi_box(ws1, 4, 3, "Cargos Ativos",
        f'=CONT.SE(Dados!F2:F{LAST_ROW},"ATIVO")',
        f'=CONT.SE(Dados!F2:F{LAST_ROW},"ATIVO")/CONT.VALORES(Dados!A2:A{LAST_ROW})',
        GREEN_FILL)

kpi_box(ws1, 4, 5, "Cargos Inativos",
        f'=CONT.SE(Dados!F2:F{LAST_ROW},"INATIVO")',
        f'=CONT.SE(Dados!F2:F{LAST_ROW},"INATIVO")/CONT.VALORES(Dados!A2:A{LAST_ROW})',
        YELLOW_FILL)

# Linha 8-10
kpi_box(ws1, 8, 1, "Total de Empresas",
        f'=CONT.VALORES(ÚNICO(Dados!A2:A{LAST_ROW}))',
        None, LIGHT_BLUE)

kpi_box(ws1, 8, 3, "Famílias GBS",
        f'=CONT.VALORES(ÚNICO(Dados!H2:H{LAST_ROW}))',
        None, LIGHT_BLUE)

kpi_box(ws1, 8, 5, "Grupos Hierárquicos",
        f'=CONT.VALORES(ÚNICO(Dados!J2:J{LAST_ROW}))',
        None, LIGHT_BLUE)

# Linha 12-14
kpi_box(ws1, 12, 1, "Sem Cod. Mercer 2026",
        f'=CONT.SE(Dados!U2:U{LAST_ROW},"")',
        f'=CONT.SE(Dados!U2:U{LAST_ROW},"")/CONT.VALORES(Dados!A2:A{LAST_ROW})',
        YELLOW_FILL)

kpi_box(ws1, 12, 3, "Corporativo",
        f'=CONT.SE(Dados!G2:G{LAST_ROW},"Corporativo")',
        f'=CONT.SE(Dados!G2:G{LAST_ROW},"Corporativo")/CONT.VALORES(Dados!A2:A{LAST_ROW})',
        LIGHT_BLUE)

kpi_box(ws1, 12, 5, "Negócios & Operações",
        f'=CONT.SE(Dados!G2:G{LAST_ROW},"Negócios & operações")',
        f'=CONT.SE(Dados!G2:G{LAST_ROW},"Negócios & operações")/CONT.VALORES(Dados!A2:A{LAST_ROW})',
        LIGHT_BLUE)

for r in [3, 7, 11, 15]:
    ws1.row_dimensions[r].height = 8
for r in [4, 5, 6, 8, 9, 10, 12, 13, 14]:
    ws1.row_dimensions[r].height = 24
for col in range(1, 7):
    ws1.column_dimensions[get_column_letter(col)].width = 20


# ════════════════════════════════════════════════════════════════════════════
# ABA 2 – Distribuição por Grade
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Distribuição por Grade")
ws2.sheet_view.showGridLines = False

grades = [2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
title_row(ws2, "Distribuição de Cargos por Grade", 5)
header_cells(ws2, 2, ["Grade", "Ativos", "Inativos", "Total", "% do Total"])

for i, g in enumerate(grades):
    r = 3 + i
    bg = ALT_ROW if i % 2 == 0 else WHITE
    # Escreve o valor do grade como número
    ws2.cell(row=r, column=1).value = g
    # Fórmulas CONT.SES
    ws2.cell(row=r, column=2).value = (
        f'=CONT.SES(Dados!E$2:E${LAST_ROW},A{r},Dados!F$2:F${LAST_ROW},"ATIVO")')
    ws2.cell(row=r, column=3).value = (
        f'=CONT.SES(Dados!E$2:E${LAST_ROW},A{r},Dados!F$2:F${LAST_ROW},"INATIVO")')
    ws2.cell(row=r, column=4).value = f'=B{r}+C{r}'
    ws2.cell(row=r, column=5).value = f'=D{r}/CONT.VALORES(Dados!A$2:A${LAST_ROW})'
    ws2.cell(row=r, column=5).number_format = '0,0%'
    style_row(ws2, r, 5, bg)

# Linha de total
tr = 3 + len(grades)
ws2.cell(row=tr, column=1).value = "TOTAL"
ws2.cell(row=tr, column=2).value = f'=SOMA(B3:B{tr-1})'
ws2.cell(row=tr, column=3).value = f'=SOMA(C3:C{tr-1})'
ws2.cell(row=tr, column=4).value = f'=SOMA(D3:D{tr-1})'
ws2.cell(row=tr, column=5).value = f'=SOMA(E3:E{tr-1})'
ws2.cell(row=tr, column=5).number_format = '0,0%'
style_row(ws2, tr, 5, DARK_BLUE, bold=True)

# Gráfico
chart = BarChart()
chart.type = "col"
chart.title = "Cargos Ativos por Grade"
chart.y_axis.title = "Quantidade"
chart.x_axis.title = "Grade"
chart.style = 10
chart.height = 12
chart.width = 18
chart.add_data(Reference(ws2, min_col=2, max_col=2, min_row=2, max_row=2+len(grades)),
               titles_from_data=True)
chart.set_categories(Reference(ws2, min_col=1, min_row=3, max_row=2+len(grades)))
chart.series[0].graphicalProperties.solidFill = MID_BLUE
ws2.add_chart(chart, "G2")
auto_width(ws2)


# ════════════════════════════════════════════════════════════════════════════
# ABA 3 – Análise por Família GBS
# ════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Análise por Família GBS")
ws3.sheet_view.showGridLines = False

families = ['Administração & Serviços', 'Arquitetura & Engenharia',
            'Atuarial, modelagem e dados', 'Privacidade de dados']
subfamilies_map = {
    'Administração & Serviços':    ['Facilities e Serviços Gerais','Secretaria',
                                    'Serviços administrativos','Serviços operacionais'],
    'Arquitetura & Engenharia':    ['Arquitetura','Engenharia'],
    'Atuarial, modelagem e dados': ['Atuarial','Ciência de Dados','Informações Gerenciais'],
    'Privacidade de dados':        ['Privacidade de dados'],
}
hier = ['GA','GS','Coordenador','Supervisor','Especialista','Analista','Assistente','Auxiliar']
ncols3 = 2 + len(hier) + 1  # Família + SubFamília + 8 hierarquias + Total

title_row(ws3, "Análise por Família GBS × Grupo Hierárquico", ncols3)
headers3 = ["Família GBS", "Sub Família"] + hier + ["TOTAL"]
header_cells(ws3, 2, headers3)

r = 3
for fi, fam in enumerate(families):
    subs = subfamilies_map[fam]
    bg_fam = "EBF3FB" if fi % 2 == 0 else WHITE
    for sf in subs:
        ws3.cell(row=r, column=1).value = fam if sf == subs[0] else ""
        ws3.cell(row=r, column=2).value = sf
        col_start = 3
        for hi, h in enumerate(hier):
            col = col_start + hi
            col_letter = get_column_letter(col)
            ws3.cell(row=r, column=col).value = (
                f'=CONT.SES(Dados!H$2:H${LAST_ROW},"{fam}",'
                f'Dados!I$2:I${LAST_ROW},"{sf}",'
                f'Dados!J$2:J${LAST_ROW},"{h}")')
        tot_col = col_start + len(hier)
        ws3.cell(row=r, column=tot_col).value = (
            f'=SOMA({get_column_letter(col_start)}{r}:{get_column_letter(col_start+len(hier)-1)}{r})')
        style_row(ws3, r, ncols3, bg_fam)
        ws3.cell(row=r, column=1).alignment = left()
        ws3.cell(row=r, column=2).alignment = left()
        r += 1
    # Subtotal da família
    ws3.cell(row=r, column=1).value = f"Subtotal – {fam}"
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    sub_start_row = r - len(subs)
    for hi in range(len(hier)):
        col = 3 + hi
        cl = get_column_letter(col)
        ws3.cell(row=r, column=col).value = (
            f'=SOMA({cl}{sub_start_row}:{cl}{r-1})')
    tot_col = 3 + len(hier)
    ws3.cell(row=r, column=tot_col).value = (
        f'=SOMA({get_column_letter(3)}{r}:{get_column_letter(2+len(hier))}{r})')
    style_row(ws3, r, ncols3, LIGHT_BLUE, bold=True, font_color=DARK_BLUE)
    ws3.cell(row=r, column=1).alignment = left()
    r += 1

# Total geral
ws3.cell(row=r, column=1).value = "TOTAL GERAL"
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
for hi in range(len(hier)):
    col = 3 + hi
    ws3.cell(row=r, column=col).value = (
        f'=CONT.SES(Dados!J$2:J${LAST_ROW},"{hier[hi]}")')
tot_col = 3 + len(hier)
ws3.cell(row=r, column=tot_col).value = (
    f'=SOMA({get_column_letter(3)}{r}:{get_column_letter(2+len(hier))}{r})')
style_row(ws3, r, ncols3, DARK_BLUE, bold=True)
ws3.cell(row=r, column=1).alignment = left()
auto_width(ws3, max_w=40)


# ════════════════════════════════════════════════════════════════════════════
# ABA 4 – Análise por Empresa
# ════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Análise por Empresa")
ws4.sheet_view.showGridLines = False

empresas = ['AFFINITY','BARE','BSP','BVP','CAP','HOLDING','MEDSERVICE','NOVAMED','SAUDE','SAUDE OPERADORA']
title_row(ws4, "Análise por Empresa", 5)
header_cells(ws4, 2, ["Empresa", "Total Cargos", "Ativos", "Inativos", "% Ativo"])

for i, emp in enumerate(empresas):
    r = 3 + i
    bg = ALT_ROW if i % 2 == 0 else WHITE
    ws4.cell(row=r, column=1).value = emp
    ws4.cell(row=r, column=2).value = f'=CONT.SE(Dados!A$2:A${LAST_ROW},"{emp}")'
    ws4.cell(row=r, column=3).value = (
        f'=CONT.SES(Dados!A$2:A${LAST_ROW},"{emp}",Dados!F$2:F${LAST_ROW},"ATIVO")')
    ws4.cell(row=r, column=4).value = (
        f'=CONT.SES(Dados!A$2:A${LAST_ROW},"{emp}",Dados!F$2:F${LAST_ROW},"INATIVO")')
    ws4.cell(row=r, column=5).value = f'=SE(B{r}>0,C{r}/B{r},0)'
    ws4.cell(row=r, column=5).number_format = '0,0%'
    style_row(ws4, r, 5, bg)
    ws4.cell(row=r, column=1).alignment = left()

tr4 = 3 + len(empresas)
ws4.cell(row=tr4, column=1).value = "TOTAL"
ws4.cell(row=tr4, column=2).value = f'=SOMA(B3:B{tr4-1})'
ws4.cell(row=tr4, column=3).value = f'=SOMA(C3:C{tr4-1})'
ws4.cell(row=tr4, column=4).value = f'=SOMA(D3:D{tr4-1})'
ws4.cell(row=tr4, column=5).value = f'=SE(B{tr4}>0,C{tr4}/B{tr4},0)'
ws4.cell(row=tr4, column=5).number_format = '0,0%'
style_row(ws4, tr4, 5, DARK_BLUE, bold=True)
ws4.cell(row=tr4, column=1).alignment = left()
auto_width(ws4)


# ════════════════════════════════════════════════════════════════════════════
# ABA 5 – Benchmarking Mercer
# ════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Benchmarking Mercer")
ws5.sheet_view.showGridLines = False

title_row(ws5, "Benchmarking Mercer – Cobertura 2025 vs 2026", 4)
header_cells(ws5, 2, ["Indicador", "Quantidade", "% Total"])

mercer_rows = [
    ("Cargos com Mercer 2025",
     f'=CONT.SE(Dados!T$2:T${LAST_ROW},"<>"&"")',
     f'=B3/CONT.VALORES(Dados!A$2:A${LAST_ROW})'),
    ("Cargos com Mercer 2026",
     f'=CONT.SE(Dados!U$2:U${LAST_ROW},"<>"&"")',
     f'=B4/CONT.VALORES(Dados!A$2:A${LAST_ROW})'),
    ("Cargos SEM Mercer 2026",
     f'=CONT.SE(Dados!U$2:U${LAST_ROW},"")',
     f'=B5/CONT.VALORES(Dados!A$2:A${LAST_ROW})'),
    ("% Cobertura Mercer 2026",
     "",
     f'=B4/CONT.VALORES(Dados!A$2:A${LAST_ROW})'),
]
bg_map = {3: GREEN_FILL, 4: GREEN_FILL, 5: YELLOW_FILL, 6: LIGHT_BLUE}
for ri, (label, qty_f, pct_f) in enumerate(mercer_rows):
    r = 3 + ri
    ws5.cell(row=r, column=1).value = label
    ws5.cell(row=r, column=2).value = qty_f if qty_f else ""
    ws5.cell(row=r, column=3).value = pct_f
    ws5.cell(row=r, column=3).number_format = '0,0%'
    style_row(ws5, r, 3, bg_map.get(r, WHITE))
    ws5.cell(row=r, column=1).alignment = left()

# Lista de cargos sem Mercer 2026 — copiados diretamente (dados estáticos da análise)
import pandas as pd
df2 = pd.read_excel(SOURCE)
df2['EMPRESA'] = df2['EMPRESA'].str.strip()
df2['FAMÍLIA GBS'] = df2['FAMÍLIA GBS'].str.strip()
sem_m26 = df2[df2['Cod. Mercer 2026'].isna()][
    ['CARGO','EMPRESA','GRADE','FAMÍLIA GBS','SITUAÇÃO']].sort_values(['FAMÍLIA GBS','GRADE'])

row_offset = 9
ws5.merge_cells(f'A{row_offset}:E{row_offset}')
hdr = ws5.cell(row=row_offset, column=1)
hdr.value = f"Cargos SEM Cod. Mercer 2026  —  total: {len(sem_m26)}"
hdr.fill = solid(DARK_BLUE)
hdr.font = Font(bold=True, color=WHITE, size=12)
hdr.alignment = center()
ws5.row_dimensions[row_offset].height = 24

header_cells(ws5, row_offset + 1, ["Cargo","Empresa","Grade","Família GBS","Situação"])
for i, (_, row_d) in enumerate(sem_m26.iterrows()):
    r = row_offset + 2 + i
    bg = ALT_ROW if i % 2 == 0 else WHITE
    vals = [row_d['CARGO'], row_d['EMPRESA'], row_d['GRADE'],
            row_d['FAMÍLIA GBS'], row_d['SITUAÇÃO']]
    for c, v in enumerate(vals, 1):
        cell = ws5.cell(row=r, column=c)
        cell.value = v
        cell.fill = solid(bg)
        cell.font = nfont(size=10)
        cell.alignment = left() if c in [1, 4] else center()
        cell.border = border()
auto_width(ws5, max_w=55)


# ════════════════════════════════════════════════════════════════════════════
# ABA 6 – Pirâmide Hierárquica
# ════════════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Pirâmide Hierárquica")
ws6.sheet_view.showGridLines = False

hier_order = ['GA','GS','Coordenador','Supervisor','Especialista','Analista','Assistente','Auxiliar']
title_row(ws6, "Pirâmide Hierárquica – Cargos Ativos", 4)
header_cells(ws6, 2, ["Grupo Hierárquico","Cargos Ativos","Cargos Inativos","Total"])

for i, grp in enumerate(hier_order):
    r = 3 + i
    bg = ALT_ROW if i % 2 == 0 else WHITE
    ws6.cell(row=r, column=1).value = grp
    ws6.cell(row=r, column=2).value = (
        f'=CONT.SES(Dados!J$2:J${LAST_ROW},"{grp}",Dados!F$2:F${LAST_ROW},"ATIVO")')
    ws6.cell(row=r, column=3).value = (
        f'=CONT.SES(Dados!J$2:J${LAST_ROW},"{grp}",Dados!F$2:F${LAST_ROW},"INATIVO")')
    ws6.cell(row=r, column=4).value = f'=B{r}+C{r}'
    style_row(ws6, r, 4, bg)
    ws6.cell(row=r, column=1).alignment = left()

tr6 = 3 + len(hier_order)
ws6.cell(row=tr6, column=1).value = "TOTAL"
ws6.cell(row=tr6, column=2).value = f'=SOMA(B3:B{tr6-1})'
ws6.cell(row=tr6, column=3).value = f'=SOMA(C3:C{tr6-1})'
ws6.cell(row=tr6, column=4).value = f'=SOMA(D3:D{tr6-1})'
style_row(ws6, tr6, 4, DARK_BLUE, bold=True)
ws6.cell(row=tr6, column=1).alignment = left()

# Gráfico de barras horizontal
chart6 = BarChart()
chart6.type = "bar"
chart6.title = "Cargos Ativos por Grupo Hierárquico"
chart6.x_axis.title = "Quantidade"
chart6.style = 10
chart6.height = 14
chart6.width = 20
chart6.add_data(Reference(ws6, min_col=2, max_col=2, min_row=2, max_row=2+len(hier_order)),
                titles_from_data=True)
chart6.set_categories(Reference(ws6, min_col=1, min_row=3, max_row=2+len(hier_order)))
chart6.series[0].graphicalProperties.solidFill = MID_BLUE
ws6.add_chart(chart6, "F2")
auto_width(ws6)


# ── Salvar ────────────────────────────────────────────────────────────────────
wb.save(OUTPUT)
print(f"Salvo: {OUTPUT}")
