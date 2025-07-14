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

def extract_main_indicator(indicator_ref):
    """Extrai o indicador principal (ex: 1 de 1.1.1 ou 1.2.3)."""
    if pd.isna(indicator_ref):
        return "Unknown"
    
    # Converter para string
    indicator_str = str(indicator_ref).strip()
    
    # Tentar extrair apenas o primeiro número (antes do primeiro ponto)
    match = re.match(r'^(\d+)', indicator_str)
    if match:
        return match.group(1)
    
    # Se não conseguir extrair, retornar o próprio valor
    return indicator_str

def create_indicator_sheet(wb, ods_name, indicator_group, main_indicator, grouped_data):
    """Cria uma aba para um indicador principal com todos os seus sub-indicadores."""
    
    # Nome da aba
    sheet_name = clean_sheet_name(f"{ods_name}_{main_indicator}")
    
    # Criar nova aba
    ws = wb.create_sheet(title=sheet_name)
    
    # --- CABEÇALHO PRINCIPAL ---
    title = f"SDG{ods_name.split('ODS')[-1]}: {main_indicator}"
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
    
    # --- ADICIONAR DADOS DOS SUB-INDICADORES ---
    current_row = 3
    
    # Mapeamento de colunas (ajuste conforme sua estrutura de CSV)
    column_mapping = {
        "Type": ["Type", "type", "Tipo"],
        "Metric and indicator reference": ["Metric and indicator reference", "Reference", "Referencia", "Ref", "Indicator"],
        "Metric / Indicator": ["Metric / Indicator", "Metric", "Indicator", "Indicador", "Metrica"],
        "Value\n(for continuous data)": ["Value", "Valor", "Value (for continuous data)", "Value\n(for continuous data)"],
        "Yes/No": ["Yes/No", "YesNo", "Sim/Não", "Status"],
        "Evidence": ["Evidence", "Evidencia", "Proof"],
        "Public\n(Yes/No)": ["Public", "Publico", "Public (Yes/No)", "Public\n(Yes/No)"]
    }
    
    # Adicionar cada sub-indicador como uma linha
    for _, row_data in grouped_data.iterrows():
        row_values = []
        
        for header in headers:
            value = ""
            # Tentar encontrar a coluna correspondente no CSV
            for possible_col in column_mapping.get(header, [header]):
                if possible_col in row_data:
                    value = row_data[possible_col]
                    break
            
            # Se não encontrou, tentar sem quebra de linha
            if value == "" and "\n" in header:
                clean_header = header.replace("\n", " ")
                if clean_header in row_data:
                    value = row_data[clean_header]
            
            row_values.append(value)
        
        # Adicionar linha de dados
        for col, value in enumerate(row_values, start=1):
            ws.cell(row=current_row, column=col, value=value)
        
        current_row += 1
    
    # --- AJUSTES DE FORMATAÇÃO ---
    # Bordas finas
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    
    # Aplicar bordas a todas as células com dados
    for row in ws.iter_rows(min_row=1, max_row=current_row-1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
    
    # Largura das colunas
    col_widths = [12, 12, 50, 15, 12, 12, 12]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    # Altura das linhas
    for row in range(1, current_row):
        ws.row_dimensions[row].height = 25
    
    # Alinhamento para células de dados
    for row in range(3, current_row):
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=col)
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
            print(f"  → Número de linhas: {len(df)}")
            
            # Encontrar a coluna de referência do indicador
            ref_column = None
            for col in ["Metric and indicator reference", "Reference", "Referencia", "Ref", "Indicator"]:
                if col in df.columns:
                    ref_column = col
                    break
            
            if ref_column is None:
                print(f"  → AVISO: Coluna de referência não encontrada. Usando índice.")
                df['temp_ref'] = df.index.astype(str)
                ref_column = 'temp_ref'
            
            # Extrair indicador principal para cada linha
            df['main_indicator'] = df[ref_column].apply(extract_main_indicator)
            
            # Agrupar por indicador principal
            grouped = df.groupby('main_indicator')
            
            print(f"  → Indicadores principais encontrados: {list(grouped.groups.keys())}")
            
            # Criar uma aba para cada indicador principal
            for main_indicator, group_data in grouped:
                create_indicator_sheet(wb, ods_name, main_indicator, main_indicator, group_data)
                total_sheets += 1
                print(f"    → Aba criada: {ods_name}_{main_indicator} ({len(group_data)} indicadores)")
            
        except Exception as e:
            print(f"Erro ao processar {csv_file.name}: {e}")
            import traceback
            traceback.print_exc()
    
    if total_sheets == 0:
        print("Nenhuma aba foi criada!")
        return
    
    # Salvar o arquivo
    wb.save(excel_path)
    print(f"\nExcel formatado gerado com sucesso: {excel_path}")
    print(f"Total de abas criadas: {total_sheets}")

if __name__ == "__main__":
    generate_formatted_excel()
