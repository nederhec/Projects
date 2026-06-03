"""
Dashboard Generator - Cargos GBS
Generates Dashboard_Cargos_GBS.xlsx with 6 sheets.
"""

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

SOURCE = '/root/.claude/uploads/39518091-0fb9-4665-a41a-ed8e752c52fd/88b7cf84-Cargos_GBS.xlsx'
OUTPUT = '/home/user/Projects/Dashboard_Cargos_GBS.xlsx'

# ── colour palette ──────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "BDD7EE"
LIGHT_GREY  = "F2F2F2"
DARK_GREY   = "595959"
WHITE       = "FFFFFF"
ALT_ROW     = "DEEAF1"

def solid(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style='thin', color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def header_font(white=True, size=11):
    return Font(bold=True, color=WHITE if white else DARK_BLUE, size=size)

def normal_font(bold=False, size=10):
    return Font(bold=bold, size=size)

def center():
    return Alignment(horizontal='center', vertical='center', wrap_text=False)

def auto_width(ws, min_w=8, max_w=50):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_w, max(min_w, max_len + 2))

def style_header_row(ws, row, cols, bg=DARK_BLUE, font_size=11):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = solid(bg)
        cell.font = header_font(white=True, size=font_size)
        cell.alignment = center()
        cell.border = thin_border()


# ── Load data ────────────────────────────────────────────────────────────────
df = pd.read_excel(SOURCE)
df['SITUAÇÃO'] = df['SITUAÇÃO'].str.strip()
df['GRUPO'] = df['GRUPO'].str.strip()
df['FAMÍLIA GBS'] = df['FAMÍLIA GBS'].str.strip()
df['GRUPO HIERARQUICO'] = df['GRUPO HIERARQUICO'].str.strip()
df['EMPRESA'] = df['EMPRESA'].str.strip()

ativos   = df[df['SITUAÇÃO'] == 'ATIVO']
inativos = df[df['SITUAÇÃO'] == 'INATIVO']

total     = len(df)
n_ativo   = len(ativos)
n_inativo = len(inativos)
n_emp     = df['EMPRESA'].nunique()
n_fam     = df['FAMÍLIA GBS'].nunique()
n_grp_h   = df['GRUPO HIERARQUICO'].nunique()
n_sem_m26 = df['Cod. Mercer 2026'].isna().sum()
n_corp    = (df['GRUPO'] == 'Corporativo').sum()
n_neg     = (df['GRUPO'] == 'Negócios & operações').sum()

wb = Workbook()
wb.remove(wb.active)   # remove default sheet


# ════════════════════════════════════════════════════════════════════════════
# SHEET 1 – KPIs Executivos
# ════════════════════════════════════════════════════════════════════════════
ws1 = wb.create_sheet("KPIs Executivos")
ws1.sheet_view.showGridLines = False

# Title
ws1.merge_cells('A1:F1')
title = ws1['A1']
title.value = "DASHBOARD – CARGOS GBS"
title.fill  = solid(DARK_BLUE)
title.font  = Font(bold=True, color=WHITE, size=16)
title.alignment = center()
ws1.row_dimensions[1].height = 36

ws1.merge_cells('A2:F2')
sub = ws1['A2']
sub.value = "Visão Executiva de KPIs"
sub.fill  = solid(MID_BLUE)
sub.font  = Font(bold=True, color=WHITE, size=12)
sub.alignment = center()
ws1.row_dimensions[2].height = 24

def kpi_box(ws, row, col, label, value, sub_val=None, bg=LIGHT_BLUE):
    ws.merge_cells(start_row=row,   start_column=col,
                   end_row=row,     end_column=col + 1)
    ws.merge_cells(start_row=row+1, start_column=col,
                   end_row=row+1,   end_column=col + 1)
    lbl = ws.cell(row=row, column=col)
    lbl.value = label
    lbl.fill  = solid(DARK_BLUE)
    lbl.font  = Font(bold=True, color=WHITE, size=10)
    lbl.alignment = center()
    val = ws.cell(row=row+1, column=col)
    val.value = value
    val.fill  = solid(bg)
    val.font  = Font(bold=True, color=DARK_BLUE, size=14)
    val.alignment = center()
    if sub_val is not None:
        ws.merge_cells(start_row=row+2, start_column=col,
                       end_row=row+2,   end_column=col + 1)
        sv = ws.cell(row=row+2, column=col)
        sv.value = sub_val
        sv.fill  = solid(bg)
        sv.font  = Font(color=DARK_GREY, size=9)
        sv.alignment = center()

# Row 4-6
kpi_box(ws1, 4, 1, "Total de Cargos",       total,        None,                             LIGHT_BLUE)
kpi_box(ws1, 4, 3, "Cargos Ativos",         f"{n_ativo}", f"{n_ativo/total*100:.1f}% do total",  "C6EFCE")
kpi_box(ws1, 4, 5, "Cargos Inativos",       f"{n_inativo}",f"{n_inativo/total*100:.1f}% do total","FFEB9C")

# Row 8-10
kpi_box(ws1, 8, 1, "Total de Empresas",     n_emp,  None, LIGHT_BLUE)
kpi_box(ws1, 8, 3, "Famílias GBS",          n_fam,  None, LIGHT_BLUE)
kpi_box(ws1, 8, 5, "Grupos Hierárquicos",   n_grp_h,None, LIGHT_BLUE)

# Row 12-14
kpi_box(ws1, 12, 1, "Sem Cod. Mercer 2026",  f"{n_sem_m26}", f"{n_sem_m26/total*100:.1f}% do total","FFEB9C")
kpi_box(ws1, 12, 3, "Corporativo",           n_corp, f"{n_corp/total*100:.1f}% do total", LIGHT_BLUE)
kpi_box(ws1, 12, 5, "Negócios & Operações",  n_neg,  f"{n_neg/total*100:.1f}% do total",  LIGHT_BLUE)

for r in [3, 7, 11, 15]:
    ws1.row_dimensions[r].height = 8
for r in [4, 5, 6, 8, 9, 10, 12, 13, 14]:
    ws1.row_dimensions[r].height = 22
for c in range(1, 7):
    ws1.column_dimensions[get_column_letter(c)].width = 18


# ════════════════════════════════════════════════════════════════════════════
# SHEET 2 – Distribuição por Grade
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Distribuição por Grade")
ws2.sheet_view.showGridLines = False

all_grades = sorted(df['GRADE'].unique())

ws2.merge_cells('A1:E1')
t = ws2['A1']
t.value = "Distribuição de Cargos por Grade"
t.fill  = solid(DARK_BLUE)
t.font  = Font(bold=True, color=WHITE, size=14)
t.alignment = center()
ws2.row_dimensions[1].height = 30

for c, h in enumerate(["Grade", "Ativos", "Inativos", "Total", "% do Total"], 1):
    cell = ws2.cell(row=2, column=c)
    cell.value = h
    cell.fill  = solid(MID_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()

data_start = 3
for i, grade in enumerate(all_grades):
    r = data_start + i
    na = len(df[(df['GRADE'] == grade) & (df['SITUAÇÃO'] == 'ATIVO')])
    ni = len(df[(df['GRADE'] == grade) & (df['SITUAÇÃO'] == 'INATIVO')])
    tot2 = na + ni
    bg = ALT_ROW if i % 2 == 0 else WHITE
    for c, v in enumerate([grade, na, ni, tot2, tot2/total], 1):
        cell = ws2.cell(row=r, column=c)
        cell.value = v
        cell.fill  = solid(bg)
        cell.font  = normal_font()
        cell.alignment = center()
        cell.border = thin_border()
        if c == 5:
            cell.number_format = '0.0%'

tot_row = data_start + len(all_grades)
for c, v in enumerate(["TOTAL", n_ativo, n_inativo, total, 1.0], 1):
    cell = ws2.cell(row=tot_row, column=c)
    cell.value = v
    cell.fill  = solid(DARK_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()
    if c == 5:
        cell.number_format = '0.0%'

# Bar chart
chart = BarChart()
chart.type   = "col"
chart.title  = "Cargos Ativos por Grade"
chart.y_axis.title = "Quantidade"
chart.x_axis.title = "Grade"
chart.style  = 10
chart.height = 12
chart.width  = 18
data_ref = Reference(ws2, min_col=2, max_col=2, min_row=2, max_row=data_start + len(all_grades) - 1)
cats_ref = Reference(ws2, min_col=1, min_row=data_start, max_row=data_start + len(all_grades) - 1)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)
chart.series[0].graphicalProperties.solidFill = MID_BLUE
ws2.add_chart(chart, "G2")
auto_width(ws2)


# ════════════════════════════════════════════════════════════════════════════
# SHEET 3 – Análise por Família GBS
# ════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Análise por Família GBS")
ws3.sheet_view.showGridLines = False

families    = sorted(df['FAMÍLIA GBS'].unique())
hier_order  = ['GA', 'GS', 'Coordenador', 'Supervisor', 'Especialista',
               'Analista', 'Assistente', 'Auxiliar']
hier_present= [h for h in hier_order if h in df['GRUPO HIERARQUICO'].unique()]

total_cols3 = 3 + len(hier_present)
ws3.merge_cells('A1:' + get_column_letter(total_cols3) + '1')
t = ws3['A1']
t.value = "Análise por Família GBS × Grupo Hierárquico"
t.fill  = solid(DARK_BLUE)
t.font  = Font(bold=True, color=WHITE, size=14)
t.alignment = center()
ws3.row_dimensions[1].height = 30

ws3.cell(row=2, column=1).value = "Família GBS"
ws3.cell(row=2, column=2).value = "Sub Família"
for ci, h in enumerate(hier_present, 3):
    ws3.cell(row=2, column=ci).value = h
ws3.cell(row=2, column=total_cols3).value = "TOTAL"
style_header_row(ws3, 2, total_cols3, bg=MID_BLUE)

r = 3
for fam_i, fam in enumerate(families):
    sub = df[df['FAMÍLIA GBS'] == fam]
    subfams = sorted(sub['SUB FAMÍLIA'].unique())
    for sf in subfams:
        sub2 = sub[sub['SUB FAMÍLIA'] == sf]
        bg = "EBF3FB" if fam_i % 2 == 0 else WHITE
        ws3.cell(row=r, column=1).value = fam if sf == subfams[0] else ""
        ws3.cell(row=r, column=2).value = sf
        row_total = 0
        for ci, h in enumerate(hier_present, 3):
            cnt = len(sub2[sub2['GRUPO HIERARQUICO'] == h])
            ws3.cell(row=r, column=ci).value = cnt if cnt > 0 else ""
            row_total += cnt
        ws3.cell(row=r, column=total_cols3).value = row_total
        for c in range(1, total_cols3 + 1):
            cell = ws3.cell(row=r, column=c)
            cell.fill  = solid(bg)
            cell.font  = normal_font()
            cell.alignment = center()
            cell.border = thin_border()
        ws3.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center')
        ws3.cell(row=r, column=2).alignment = Alignment(horizontal='left', vertical='center')
        r += 1
    # Subtotal
    ws3.cell(row=r, column=1).value = f"Subtotal – {fam}"
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    fam_total = 0
    for ci, h in enumerate(hier_present, 3):
        cnt = len(sub[sub['GRUPO HIERARQUICO'] == h])
        ws3.cell(row=r, column=ci).value = cnt if cnt > 0 else ""
        fam_total += cnt
    ws3.cell(row=r, column=total_cols3).value = fam_total
    for c in range(1, total_cols3 + 1):
        cell = ws3.cell(row=r, column=c)
        cell.fill  = solid(LIGHT_BLUE)
        cell.font  = Font(bold=True, size=10, color=DARK_BLUE)
        cell.alignment = center()
        cell.border = thin_border()
    ws3.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center')
    r += 1

# Grand total
ws3.cell(row=r, column=1).value = "TOTAL GERAL"
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
grand = 0
for ci, h in enumerate(hier_present, 3):
    cnt = len(df[df['GRUPO HIERARQUICO'] == h])
    ws3.cell(row=r, column=ci).value = cnt
    grand += cnt
ws3.cell(row=r, column=total_cols3).value = grand
for c in range(1, total_cols3 + 1):
    cell = ws3.cell(row=r, column=c)
    cell.fill  = solid(DARK_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()
auto_width(ws3, max_w=40)


# ════════════════════════════════════════════════════════════════════════════
# SHEET 4 – Análise por Empresa
# ════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Análise por Empresa")
ws4.sheet_view.showGridLines = False

ws4.merge_cells('A1:G1')
t = ws4['A1']
t.value = "Análise por Empresa"
t.fill  = solid(DARK_BLUE)
t.font  = Font(bold=True, color=WHITE, size=14)
t.alignment = center()
ws4.row_dimensions[1].height = 30

for c, h in enumerate(["Empresa","Total Cargos","Ativos","Inativos",
                        "% Ativo","Grades Distintos","Famílias GBS Distintas"], 1):
    cell = ws4.cell(row=2, column=c)
    cell.value = h
    cell.fill  = solid(MID_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()

for i, emp in enumerate(sorted(df['EMPRESA'].unique())):
    r = 3 + i
    sub = df[df['EMPRESA'] == emp]
    na2 = len(sub[sub['SITUAÇÃO'] == 'ATIVO'])
    ni2 = len(sub[sub['SITUAÇÃO'] == 'INATIVO'])
    tot2 = len(sub)
    bg = ALT_ROW if i % 2 == 0 else WHITE
    for c, v in enumerate([emp, tot2, na2, ni2, na2/tot2 if tot2 > 0 else 0,
                            sub['GRADE'].nunique(), sub['FAMÍLIA GBS'].nunique()], 1):
        cell = ws4.cell(row=r, column=c)
        cell.value = v
        cell.fill  = solid(bg)
        cell.font  = normal_font()
        cell.alignment = center()
        cell.border = thin_border()
        if c == 5:
            cell.number_format = '0.0%'

n_emp_rows = df['EMPRESA'].nunique()
for c, v in enumerate(["TOTAL", total, n_ativo, n_inativo, n_ativo/total, "", ""], 1):
    cell = ws4.cell(row=3 + n_emp_rows, column=c)
    cell.value = v
    cell.fill  = solid(DARK_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()
    if c == 5 and isinstance(v, float):
        cell.number_format = '0.0%'
auto_width(ws4)


# ════════════════════════════════════════════════════════════════════════════
# SHEET 5 – Benchmarking Mercer
# ════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Benchmarking Mercer")
ws5.sheet_view.showGridLines = False

ws5.merge_cells('A1:F1')
t = ws5['A1']
t.value = "Benchmarking Mercer – Cobertura 2025 vs 2026"
t.fill  = solid(DARK_BLUE)
t.font  = Font(bold=True, color=WHITE, size=14)
t.alignment = center()
ws5.row_dimensions[1].height = 30

n_m25  = df['Cod. Mercer 2025'].notna().sum()
n_m26  = df['Cod. Mercer 2026'].notna().sum()
n_s26  = df['Cod. Mercer 2026'].isna().sum()
pct_cob= n_m26 / total

summ_data = [
    ("Indicador",              "Quantidade", "% Total"),
    ("Cargos com Mercer 2025", n_m25,         n_m25/total),
    ("Cargos com Mercer 2026", n_m26,         n_m26/total),
    ("Cargos SEM Mercer 2026", n_s26,         n_s26/total),
    ("% Cobertura Mercer 2026","",            pct_cob),
]
for ri, row_data in enumerate(summ_data):
    r = 3 + ri
    for c, v in enumerate(row_data, 2):
        cell = ws5.cell(row=r, column=c)
        cell.value = v
        if ri == 0:
            cell.fill  = solid(MID_BLUE)
            cell.font  = header_font()
        else:
            bg_map = {1: "C6EFCE", 2: "C6EFCE", 3: "FFEB9C", 4: LIGHT_BLUE}
            cell.fill  = solid(bg_map.get(ri, WHITE))
            cell.font  = Font(bold=(c == 2), size=10)
        cell.alignment = center()
        cell.border = thin_border()
        if c == 4 and ri > 0 and isinstance(v, float):
            cell.number_format = '0.0%'

# List of cargos SEM Mercer 2026
sem_df = df[df['Cod. Mercer 2026'].isna()][['CARGO','EMPRESA','GRADE','FAMÍLIA GBS','SITUAÇÃO']].copy()
sem_df = sem_df.sort_values(['FAMÍLIA GBS', 'GRADE'])

row_offset = 11
ws5.merge_cells(f'A{row_offset}:E{row_offset}')
hdr = ws5.cell(row=row_offset, column=1)
hdr.value = f"Cargos SEM Cod. Mercer 2026 ({len(sem_df)} cargos)"
hdr.fill  = solid(DARK_BLUE)
hdr.font  = Font(bold=True, color=WHITE, size=12)
hdr.alignment = center()
ws5.row_dimensions[row_offset].height = 24

for c, h in enumerate(["Cargo","Empresa","Grade","Família GBS","Situação"], 1):
    cell = ws5.cell(row=row_offset + 1, column=c)
    cell.value = h
    cell.fill  = solid(MID_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()

for i, (_, row_d) in enumerate(sem_df.iterrows()):
    r = row_offset + 2 + i
    bg = ALT_ROW if i % 2 == 0 else WHITE
    for c, v in enumerate([row_d['CARGO'], row_d['EMPRESA'], row_d['GRADE'],
                            row_d['FAMÍLIA GBS'], row_d['SITUAÇÃO']], 1):
        cell = ws5.cell(row=r, column=c)
        cell.value = v
        cell.fill  = solid(bg)
        cell.font  = normal_font()
        cell.alignment = Alignment(horizontal='left' if c in [1, 4] else 'center',
                                    vertical='center')
        cell.border = thin_border()
auto_width(ws5, max_w=55)


# ════════════════════════════════════════════════════════════════════════════
# SHEET 6 – Pirâmide Hierárquica
# ════════════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Pirâmide Hierárquica")
ws6.sheet_view.showGridLines = False

ws6.merge_cells('A1:D1')
t = ws6['A1']
t.value = "Pirâmide Hierárquica – Cargos Ativos"
t.fill  = solid(DARK_BLUE)
t.font  = Font(bold=True, color=WHITE, size=14)
t.alignment = center()
ws6.row_dimensions[1].height = 30

for c, h in enumerate(["Grupo Hierárquico","Cargos Ativos","Cargos Inativos","Total"], 1):
    cell = ws6.cell(row=2, column=c)
    cell.value = h
    cell.fill  = solid(MID_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()

hier_full = ['GA','GS','Coordenador','Supervisor','Especialista','Analista','Assistente','Auxiliar']
present_order = [h for h in hier_full if h in df['GRUPO HIERARQUICO'].unique()]
present_order += [h for h in df['GRUPO HIERARQUICO'].unique() if h not in present_order]

for i, grp in enumerate(present_order):
    r = 3 + i
    na2 = len(ativos[ativos['GRUPO HIERARQUICO'] == grp])
    ni2 = len(inativos[inativos['GRUPO HIERARQUICO'] == grp])
    bg = ALT_ROW if i % 2 == 0 else WHITE
    for c, v in enumerate([grp, na2, ni2, na2 + ni2], 1):
        cell = ws6.cell(row=r, column=c)
        cell.value = v
        cell.fill  = solid(bg)
        cell.font  = normal_font()
        cell.alignment = center()
        cell.border = thin_border()

tot_r6 = 3 + len(present_order)
for c, v in enumerate(["TOTAL", n_ativo, n_inativo, total], 1):
    cell = ws6.cell(row=tot_r6, column=c)
    cell.value = v
    cell.fill  = solid(DARK_BLUE)
    cell.font  = header_font()
    cell.alignment = center()
    cell.border = thin_border()

# Horizontal bar chart
chart6 = BarChart()
chart6.type   = "bar"
chart6.title  = "Cargos Ativos por Grupo Hierárquico"
chart6.x_axis.title = "Quantidade"
chart6.y_axis.title = "Grupo Hierárquico"
chart6.style  = 10
chart6.height = 14
chart6.width  = 20
data6 = Reference(ws6, min_col=2, max_col=2, min_row=2, max_row=2 + len(present_order))
cats6 = Reference(ws6, min_col=1, min_row=3, max_row=2 + len(present_order))
chart6.add_data(data6, titles_from_data=True)
chart6.set_categories(cats6)
chart6.series[0].graphicalProperties.solidFill = MID_BLUE
ws6.add_chart(chart6, "F2")
auto_width(ws6)


# ── Save ─────────────────────────────────────────────────────────────────────
wb.save(OUTPUT)
print(f"Saved: {OUTPUT}")
