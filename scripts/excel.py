import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

def format_sheet(ws, ods_name):
    """Formata uma aba do Excel conforme o template desejado."""
    # --- CABEÇALHO PRINCIPAL ---
    title = f"SDG{ods_name.split('ODS')[-1]}: {ods_name}"  # Ex: "SDG5: Gender Equality"
    ws.append([title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    
    # Estilo do título
    title_cell = ws.cell(row=1, column=1)
    title_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    title_cell.font = Font(color="FFFFFF", bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # --- CABEÇALHOS DA TABELA ---
    headers = [
        "Type",
        "Metric and indicator reference",
        "Metric / Indicator", 
        "Value\n(for continuous data)",
        "Yes/No",
        "Evidence",
        "Public\n(Yes/No)"
    ]
    ws.append(headers)

    # Estilo dos cabeçalhos
    header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=2, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # --- AJUSTES DE FORMATAÇÃO ---
    # Bordas finas
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    # Largura das colunas
    col_widths = [12, 12, 50, 15, 12, 12, 12]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    # Altura das linhas
    for row in range(1, 50):  # Ajuste conforme necessário
        ws.row_dimensions[row].height = 25

def generate_formatted_excel():
    # Criar pasta se não existir
    excel_dir = Path('scripts/excel')
    excel_dir.mkdir(parents=True, exist_ok=True)
    
    # Listar CSVs e agrupar por ODS (assumindo que o nome do CSV contém "ODSX")
    data_dir = Path('data')
    ods_data = {}  # Dicionário para armazenar dados por ODS

    for csv_file in data_dir.glob('*.csv'):
        try:
            df = pd.read_csv(csv_file)
            ods_name = csv_file.stem.split('_')[0]  # Ex: "ODS5" do arquivo "ODS5_gender.csv"
            
            if ods_name not in ods_data:
                ods_data[ods_name] = []
            ods_data[ods_name].append(df)
            
            print(f"Processado: {csv_file.name} → {ods_name}")
        except Exception as e:
            print(f"Erro ao processar {csv_file.name}: {e}")

    # Criar Excel com abas formatadas
    excel_path = excel_dir / 'sdg-data-formatado.xlsx'
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for ods_name, dfs in ods_data.items():
            combined_df = pd.concat(dfs, ignore_index=True)
            combined_df.to_excel(writer, sheet_name=ods_name, index=False)
            
            # Acessar a aba criada para formatar
            ws = writer.sheets[ods_name]
            format_sheet(ws, ods_name)

    print(f"Excel formatado gerado: {excel_path}")

if __name__ == "__main__":
    generate_formatted_excel()