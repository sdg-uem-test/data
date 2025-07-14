import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import re

def clean_sheet_name(name):
    """Limpa o nome da aba para ser válido no Excel."""
    # Remove caracteres inválidos para nomes de aba
    name = re.sub(r'[\\/*?:"<>|]', '', str(name))
    # Limita a 31 caracteres (limite do Excel)
    if len(name) > 31:
        name = name[:28] + "..."
    return name

def create_indicator_sheet(wb, ods_name, indicator_data, sheet_name):
    """Cria uma aba para um indicador específico."""
    
    # Criar nova aba
    ws = wb.create_sheet(title=sheet_name)
    
    # --- CABEÇALHO PRINCIPAL ---
    title = f"SDG{ods_name.split('ODS')[-1]}: {ods_name}"
    ws.cell(row=1, column=1, value=title)
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
    
    # Adicionar cabeçalhos
    for col, header in enumerate(headers, start=1):
        ws.cell(row=2, column=col, value=header)
    
    # Estilo dos cabeçalhos
    header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=2, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # --- ADICIONAR DADOS DO INDICADOR ---
    # Mapear dados do CSV para as colunas do template
    row_data = []
    
    # Tente mapear as colunas do CSV para o template
    # Ajuste este mapeamento conforme a estrutura real dos seus CSVs
    column_mapping = {
        "Type": ["Type", "type", "Tipo"],
        "Metric and indicator reference": ["Metric and indicator reference", "Reference", "Referencia", "Ref"],
        "Metric / Indicator": ["Metric / Indicator", "Metric", "Indicator", "Indicador", "Metrica"],
        "Value\n(for continuous data)": ["Value", "Valor", "Value (for continuous data)"],
        "Yes/No": ["Yes/No", "YesNo", "Sim/Não", "Status"],
        "Evidence": ["Evidence", "Evidencia", "Proof"],
        "Public\n(Yes/No)": ["Public", "Publico", "Public (Yes/No)"]
    }
    
    for header in headers:
        value = ""
        # Tentar encontrar a coluna correspondente no CSV
        for possible_col in column_mapping.get(header, [header]):
            if possible_col in indicator_data:
                value = indicator_data[possible_col]
                break
        row_data.append(value)
    
    # Adicionar linha de dados
    for col, value in enumerate(row_data, start=1):
        ws.cell(row=3, column=col, value=value)
    
    # --- AJUSTES DE FORMATAÇÃO ---
    # Bordas finas
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    
    # Aplicar bordas a todas as células com dados
    for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
    
    # Largura das colunas
    col_widths = [12, 12, 50, 15, 12, 12, 12]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    # Altura das linhas
    for row in range(1, 4):
        ws.row_dimensions[row].height = 25
    
    # Alinhamento para células de dados
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

def generate_formatted_excel():
    # Criar pasta se não existir
    excel_dir = Path('scripts/excel')
    excel_dir.mkdir(parents=True, exist_ok=True)
    
    # Verificar se a pasta data existe
    data_dir = Path('data')
    if not data_dir.exists():
        print(f"Erro: A pasta {data_dir} não existe!")
        return
    
    # Listar CSVs
    csv_files = list(data_dir.glob('*.csv'))
    
    if not csv_files:
        print(f"Nenhum arquivo CSV encontrado em {data_dir}")
        return
    
    print(f"Encontrados {len(csv_files)} arquivos CSV em {data_dir}")
    
    # Criar Excel
    excel_path = excel_dir / 'sdg-data-formatado.xlsx'
    wb = openpyxl.Workbook()
    
    # Remover a aba padrão
    wb.remove(wb.active)
    
    total_sheets = 0
    
    for csv_file in csv_files:
        try:
            print(f"Processando: {csv_file.name}")
            df = pd.read_csv(csv_file)
            
            # Extrair nome do ODS do nome do arquivo
            ods_name = csv_file.stem.split('_')[0]  # Ex: "ODS5" do arquivo "ODS5_gender.csv"
            
            print(f"  → ODS: {ods_name}")
            print(f"  → Colunas encontradas: {df.columns.tolist()}")
            print(f"  → Número de indicadores: {len(df)}")
            
            # Criar uma aba para cada indicador (cada linha do CSV)
            for idx, row in df.iterrows():
                # Criar nome da aba baseado no indicador
                if 'Metric / Indicator' in df.columns:
                    indicator_name = str(row['Metric / Indicator'])
                elif 'Metric' in df.columns:
                    indicator_name = str(row['Metric'])
                elif 'Indicator' in df.columns:
                    indicator_name = str(row['Indicator'])
                else:
                    indicator_name = f"Indicator_{idx+1}"
                
                # Limpar nome da aba
                sheet_name = clean_sheet_name(f"{ods_name}_{idx+1}_{indicator_name}")
                
                # Criar aba para este indicador
                create_indicator_sheet(wb, ods_name, row, sheet_name)
                total_sheets += 1
                
                print(f"    → Aba criada: {sheet_name}")
            
        except Exception as e:
            print(f"Erro ao processar {csv_file.name}: {e}")
    
    if total_sheets == 0:
        print("Nenhuma aba foi criada!")
        return
    
    # Salvar o arquivo
    wb.save(excel_path)
    print(f"\nExcel formatado gerado com sucesso: {excel_path}")
    print(f"Total de abas criadas: {total_sheets}")

if __name__ == "__main__":
    generate_formatted_excel()
