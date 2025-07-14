import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

def format_sheet(ws, ods_name):
    """Formata uma aba do Excel conforme o template desejado."""
    # Encontrar onde os dados reais começam (após os cabeçalhos originais do CSV)
    data_start_row = 2  # Assumindo que os dados começam na linha 2
    
    # Inserir 2 linhas no topo para o título e novos cabeçalhos
    ws.insert_rows(1, 2)
    
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
    
    # Substituir os cabeçalhos originais pelos novos
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
    
    # --- AJUSTES DE FORMATAÇÃO ---
    # Bordas finas para todas as células com dados
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    
    # Aplicar bordas a todas as células com dados
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
    
    # Largura das colunas
    col_widths = [12, 12, 50, 15, 12, 12, 12]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    # Altura das linhas
    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 25

def generate_formatted_excel():
    # Criar pasta se não existir
    excel_dir = Path('scripts/excel')
    excel_dir.mkdir(parents=True, exist_ok=True)
    
    # Verificar se a pasta data existe
    data_dir = Path('data')
    if not data_dir.exists():
        print(f"Erro: A pasta {data_dir} não existe!")
        return
    
    # Listar CSVs e agrupar por ODS
    ods_data = {}
    csv_files = list(data_dir.glob('*.csv'))
    
    if not csv_files:
        print(f"Nenhum arquivo CSV encontrado em {data_dir}")
        return
    
    print(f"Encontrados {len(csv_files)} arquivos CSV em {data_dir}")
    
    for csv_file in csv_files:
        try:
            print(f"Processando: {csv_file.name}")
            df = pd.read_csv(csv_file)
            
            # Extrair nome do ODS do nome do arquivo
            ods_name = csv_file.stem.split('_')[0]  # Ex: "ODS5" do arquivo "ODS5_gender.csv"
            
            if ods_name not in ods_data:
                ods_data[ods_name] = []
            ods_data[ods_name].append(df)
            
            print(f"  → Adicionado ao grupo: {ods_name} ({len(df)} linhas)")
            
        except Exception as e:
            print(f"Erro ao processar {csv_file.name}: {e}")
    
    if not ods_data:
        print("Nenhum dado foi processado com sucesso!")
        return
    
    # Criar Excel com abas formatadas - VOLTA AO MÉTODO ORIGINAL
    excel_path = excel_dir / 'sdg-data-formatado.xlsx'
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for ods_name, dfs in ods_data.items():
            print(f"Criando aba para: {ods_name}")
            
            # Combinar todos os DataFrames do mesmo ODS
            combined_df = pd.concat(dfs, ignore_index=True)
            
            # Escrever dados na aba (isso mantém os dados originais)
            combined_df.to_excel(writer, sheet_name=ods_name, index=False)
            
            # Acessar a aba criada para formatar
            ws = writer.sheets[ods_name]
            format_sheet(ws, ods_name)
            
            print(f"  → Aba {ods_name} criada com {len(combined_df)} linhas")
    
    print(f"\nExcel formatado gerado com sucesso: {excel_path}")
    print(f"Abas criadas: {list(ods_data.keys())}")

if __name__ == "__main__":
    generate_formatted_excel()
