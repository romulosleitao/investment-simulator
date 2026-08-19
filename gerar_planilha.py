import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Criando o workbook
wb = openpyxl.Workbook()

# -------------------------------------------------------------
# ABA 1: APP (Dashboard Principal)
# -------------------------------------------------------------
ws_app = wb.active
ws_app.title = "APP"
ws_app.views.sheetView[0].showGridLines = True

# Paleta de Cores e Estilos Corporativos
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

accent_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
accent_font = Font(name="Calibri", size=11, bold=True, color="000000")

bold_font = Font(name="Calibri", size=11, bold=True)
regular_font = Font(name="Calibri", size=11)

# Título do Dashboard
ws_app["B2"] = "SIMULADOR DE INVESTIMENTOS EM FIIs"
ws_app["B2"].font = Font(name="Calibri", size=16, bold=True, color="1F4E78")

# 1. Configurações Globais
ws_app["B9"] = "CONFIGURAÇÕES"
ws_app["B9"].font = header_font
ws_app["B9"].fill = header_fill

configs = [
    ("Salário", 2000),
    ("Rendimento Carteira", 0.006),
    ("Sugestão de Investimento (30%)", "=B10*0.3"),
]

for idx, (label, val) in enumerate(configs, start=10):
    ws_app[f"B{idx}"] = label
    ws_app[f"B{idx}"].font = regular_font
    ws_app[f"D{idx}"] = val
    ws_app[f"D{idx}"].font = regular_font
    if label == "Rendimento Carteira":
        ws_app[f"D{idx}"].number_format = "0.0%"
    elif "Salário" in label or "Sugestão" in label:
        ws_app[f"D{idx}"].number_format = "R$ #,##0.00"

# 2. Simulador de Investimento Mensal
ws_app["B14"] = "INVESTIMENTO MENSAL"
ws_app["B14"].font = header_font
ws_app["B14"].fill = header_fill

sim_inputs = [
    ("Quanto investir por mês ?", 200),
    ("Por Quantos Anos ?", 5),
    ("Taxa de Rendimento mensal ?", "=D11"),
    ("Patrimônio acumulado ?", "=FV(D17; D16*12; -D15)"),
    ("Dividendos Mensais ?", "=D18*D11"),
]

for idx, (label, val) in enumerate(sim_inputs, start=15):
    ws_app[f"B{idx}"] = label
    ws_app[f"B{idx}"].font = regular_font
    ws_app[f"D{idx}"] = val
    ws_app[f"D{idx}"].font = regular_font
    if "por mês" in label or "acumulado" in label or "Dividendos" in label:
        ws_app[f"D{idx}"].number_format = "R$ #,##0.00"
    elif "Taxa" in label:
        ws_app[f"D{idx}"].number_format = "0.00%"

# 3. Tabela de Cenários
ws_app["B21"] = "Cenários (Projeção de Longo Prazo)"
ws_app["B21"].font = header_font
ws_app["B21"].fill = header_fill

ws_app["B22"] = "Anos"
ws_app["C22"] = "Patrimônio Acumulado"
ws_app["D22"] = "Dividendos Mensais"
for col in ["B", "C", "D"]:
    cell = ws_app[f"{col}22"]
    cell.font = accent_font
    cell.fill = accent_fill

anos_lista = [2, 5, 10, 20, 30]
for idx, anos in enumerate(anos_lista, start=23):
    ws_app[f"B{idx}"] = anos
    ws_app[f"B{idx}"].font = regular_font
    ws_app[f"C{idx}"] = f"=FV($D$17; B{idx}*12; -$D$15)"
    ws_app[f"C{idx}"].font = regular_font
    ws_app[f"C{idx}"].number_format = "R$ #,##0.00"

    ws_app[f"D{idx}"] = f"=C{idx}*$D$17"
    ws_app[f"D{idx}"].font = regular_font
    ws_app[f"D{idx}"].number_format = "R$ #,##0.00"

# 4. Alocação de Ativos por Perfil
ws_app["B29"] = "DISTRIBUIÇÃO DE CARTEIRA POR PERFIL"
ws_app["B29"].font = header_font
ws_app["B29"].fill = header_fill

ws_app["B30"] = "PERFIL SELECIONADO"
ws_app["C30"] = "Moderado"
ws_app["C30"].font = bold_font

ws_app["B31"] = "VALOR A SER INVESTIDO POR MÊS"
ws_app["C31"] = "=D15"
ws_app["C31"].font = bold_font
ws_app["C31"].number_format = "R$ #,##0.00"

headers_alloc = ["TIPO DE FII", "Percentual Sugerido", "Valores (R$)"]
for i, h in enumerate(headers_alloc, start=2):
    col_letter = get_column_letter(i)
    cell = ws_app[f"{col_letter}33"]
    cell.value = h
    cell.font = accent_font
    cell.fill = accent_fill

fii_types = ["PAPEL", "TIJOLO", "HÍBRIDOS", "FOFs", "DESENVOLVIMENTO", "HOTELARIAS"]

for idx, fii in enumerate(fii_types, start=34):
    ws_app[f"B{idx}"] = fii
    ws_app[f"B{idx}"].font = regular_font
    ws_app[f"C{idx}"] = (
        f'=XLOOKUP($C$30 & "-" & B{idx}, Database_Profiles!$A$2:$A$19, Database_Profiles!$D$2:$D$19, 0)'
    )
    ws_app[f"C{idx}"].font = regular_font
    ws_app[f"C{idx}"].number_format = "0.0%"

    ws_app[f"D{idx}"] = f"=C{idx}*$C$31"
    ws_app[f"D{idx}"].font = regular_font
    ws_app[f"D{idx}"].number_format = "R$ #,##0.00"

total_row = 34 + len(fii_types)
ws_app[f"B{total_row}"] = "TOTAL"
ws_app[f"B{total_row}"].font = bold_font
ws_app[f"C{total_row}"] = f"=SUM(C34:C{total_row-1})"
ws_app[f"C{total_row}"].font = bold_font
ws_app[f"C{total_row}"].number_format = "0.0%"

ws_app[f"D{total_row}"] = f"=SUM(D34:D{total_row-1})"
ws_app[f"D{total_row}"].font = bold_font
ws_app[f"D{total_row}"].number_format = "R$ #,##0.00"


# -------------------------------------------------------------
# ABA 2: Database_Profiles (Backend de Referência)
# -------------------------------------------------------------
ws_db = wb.create_sheet(title="Database_Profiles")
ws_db.views.sheetView[0].showGridLines = True

headers_db = ["CHAVE", "PERFIL", "TIPO DE FII", "%"]
for col_idx, h in enumerate(headers_db, start=1):
    cell = ws_db.cell(row=1, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill

data_db = [
    ("Conservador-PAPEL", "Conservador", "PAPEL", 0.30),
    ("Conservador-TIJOLO", "Conservador", "TIJOLO", 0.50),
    ("Conservador-HÍBRIDOS", "Conservador", "HÍBRIDOS", 0.10),
    ("Conservador-FOFs", "Conservador", "FOFs", 0.10),
    ("Conservador-DESENVOLVIMENTO", "Conservador", "DESENVOLVIMENTO", 0.00),
    ("Conservador-HOTELARIAS", "Conservador", "HOTELARIAS", 0.00),
    ("Moderado-PAPEL", "Moderado", "PAPEL", 0.32),
    ("Moderado-TIJOLO", "Moderado", "TIJOLO", 0.35),
    ("Moderado-HÍBRIDOS", "Moderado", "HÍBRIDOS", 0.08),
    ("Moderado-FOFs", "Moderado", "FOFs", 0.05),
    ("Moderado-DESENVOLVIMENTO", "Moderado", "DESENVOLVIMENTO", 0.10),
    ("Moderado-HOTELARIAS", "Moderado", "HOTELARIAS", 0.10),
    ("Agressivo-PAPEL", "Agressivo", "PAPEL", 0.50),
    ("Agressivo-TIJOLO", "Agressivo", "TIJOLO", 0.10),
    ("Agressivo-HÍBRIDOS", "Agressivo", "HÍBRIDOS", 0.05),
    ("Agressivo-FOFs", "Agressivo", "FOFs", 0.05),
    ("Agressivo-DESENVOLVIMENTO", "Agressivo", "DESENVOLVIMENTO", 0.20),
    ("Agressivo-HOTELARIAS", "Agressivo", "HOTELARIAS", 0.10),
]

for row_idx, row_data in enumerate(data_db, start=2):
    for col_idx, val in enumerate(row_data, start=1):
        cell = ws_db.cell(row=row_idx, column=col_idx, value=val)
        cell.font = regular_font
        if col_idx == 4:
            cell.number_format = "0.0%"

# Ajustar larguras
for ws in [ws_app, ws_db]:
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

wb.save("fii_investment_simulator.xlsx")
print("Planilha gerada com sucesso!")
