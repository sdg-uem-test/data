import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import re

def format_sheet_hierarchical(ws, sdg_number, metrics_data):
    """Formata uma aba do Excel com estrutura hierárquica SDG > Métrica > Indicador."""
    
    # Limpar a planilha
    ws.delete_rows(1, ws.max_row)
    
    # --- CABEÇALHO PRINCIPAL ---
    title = f"SDG{sdg_number}: Sustainable Development Goal {sdg_number}"
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
    
    # Inserir cabeçalhos
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
    
    # --- INSERIR DADOS ORGANIZADOS ---
    current_row = 3
    
    # Ordenar métricas numericamente
    sorted_metrics = sorted(metrics_data.items(), key=lambda x: float(x[0]))
    
    for metric_num, indicators in sorted_metrics:
        # Inserir linha da métrica
        metric_row = current_row
        ws.cell(row=metric_row, column=1, value="Metric")
        ws.cell(row=metric_row, column=2, value=f"{sdg_number}.{metric_num}")
        ws.cell(row=metric_row, column=3, value=f"Metric {sdg_number}.{metric_num}")
        
        # Estilo da linha da métrica
        metric_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        metric_font = Font(bold=True)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=metric_row, column=col)
            cell.fill = metric_fill
            cell.font = metric_font
        
        current_row += 1
        
        # Ordenar indicadores numericamente
        sorted_indicators = sorted(indicators.items(), key=lambda x: float(x[0]))
        
        # Inserir dados dos indicadores
        for indicator_num, df in sorted_indicators:
            for _, row_data in df.iterrows():
                ws.cell(row=current_row, column=1, value="Indicator")
                ws.cell(row=current_row, column=2, value=f"{sdg_number}.{metric_num}.{indicator_num}")
                
                # Mapear as colunas do CSV para as colunas do template
                # Assumindo que o CSV tem estrutura padrão, adapte conforme necessário
                if len(row_data) >= 1:
                    ws.cell(row=current_row, column=3, value=str(row_data.iloc[0]) if not pd.isna(row_data.iloc[0]) else "")
                if len(row_data) >= 2:
                    ws.cell(row=current_row, column=4, value=str(row_data.iloc[1]) if not pd.isna(row_data.iloc[1]) else "")
                if len(row_data) >= 3:
                    ws.cell(row=current_row, column=5, value=str(row_data.iloc[2]) if not pd.isna(row_data.iloc[2]) else "")
                if len(row_data) >= 4:
                    ws.cell(row=current_row, column=6, value=str(row_data.iloc[3]) if not pd.isna(row_data.iloc[3]) else "")
                if len(row_data) >= 5:
                    ws.cell(row=current_row, column=7, value=str(row_data.iloc[4]) if not pd.isna(row_data.iloc[4]) else "")
                
                current_row += 1
    
    # --- AJUSTES DE FORMATAÇÃO ---
    # Bordas finas para todas as células com dados
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    
    # Aplicar bordas a todas as células com dados
    for row in ws.iter_rows(min_row=1, max_row=current_row-1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
    
    # Largura das colunas
    col_widths = [12, 20, 50, 15, 12, 12, 12]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    # Altura das linhas
    for row in range(1, current_row):
        ws.row_dimensions[row].height = 25

def parse_filename(filename):
    """Extrai SDG, métrica e indicador do nome do arquivo."""
    # Padrão: indicator_x-y-z.csv
    match = re.match(r'indicator_(\d+)-(\d+)-(\d+)\.csv', filename)
    if match:
        sdg = match.group(1)
        metric = match.group(2)
        indicator = match.group(3)
        return sdg, metric, indicator
    return None, None, None

def generate_formatted_excel():
    # Criar pasta se não existir
    excel_dir = Path('scripts/excel')
    excel_dir.mkdir(parents=True, exist_ok=True)
    
    # Verificar se a pasta data existe
    data_dir = Path('data')
    if not data_dir.exists():
        print(f"Erro: A pasta {data_dir} não existe!")
        return
    
    # Listar CSVs e organizar por SDG > Métrica > Indicador
    sdg_data = {}
    csv_files = list(data_dir.glob('indicator_*.csv'))
    
    if not csv_files:
        print(f"Nenhum arquivo CSV com padrão 'indicator_x-y-z.csv' encontrado em {data_dir}")
        return
    
    print(f"Encontrados {len(csv_files)} arquivos CSV em {data_dir}")
    
    for csv_file in csv_files:
        try:
            print(f"Processando: {csv_file.name}")
            
            # Extrair SDG, métrica e indicador do nome do arquivo
            sdg, metric, indicator = parse_filename(csv_file.name)
            
            if not all([sdg, metric, indicator]):
                print(f"  → Arquivo ignorado: nome não segue o padrão indicator_x-y-z.csv")
                continue
            
            # Ler dados do CSV
            df = pd.read_csv(csv_file)
            
            # Organizar hierarquicamente
            if sdg not in sdg_data:
                sdg_data[sdg] = {}
            
            if metric not in sdg_data[sdg]:
                sdg_data[sdg][metric] = {}
            
            sdg_data[sdg][metric][indicator] = df
            
            print(f"  → Adicionado: SDG{sdg} > Métrica {metric} > Indicador {indicator} ({len(df)} linhas)")
            
        except Exception as e:
            print(f"Erro ao processar {csv_file.name}: {e}")
    
    if not sdg_data:
        print("Nenhum dado foi processado com sucesso!")
        return
    
    # Criar Excel com abas organizadas hierarquicamente
    excel_path = excel_dir / 'sdg-data-formatado.xlsx'
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Ordenar SDGs numericamente
        sorted_sdgs = sorted(sdg_data.items(), key=lambda x: int(x[0]))
        
        for sdg_num, metrics_data in sorted_sdgs:
            sheet_name = f"SDG{sdg_num}"
            print(f"Criando aba para: {sheet_name}")
            
            # Criar aba vazia
            df_empty = pd.DataFrame()
            df_empty.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Acessar a aba criada para formatar
            ws = writer.sheets[sheet_name]
            format_sheet_hierarchical(ws, sdg_num, metrics_data)
            
            total_indicators = sum(len(indicators) for indicators in metrics_data.values())
            print(f"  → Aba {sheet_name} criada com {len(metrics_data)} métricas e {total_indicators} indicadores")
    
    print(f"\nExcel formatado gerado com sucesso: {excel_path}")
    print(f"Abas criadas: {[f'SDG{sdg}' for sdg in sorted(sdg_data.keys(), key=int)]}")
    
    # Resumo da estrutura
    print("\n=== ESTRUTURA CRIADA ===")
    for sdg_num in sorted(sdg_data.keys(), key=int):
        print(f"SDG{sdg_num}:")
        for metric_num in sorted(sdg_data[sdg_num].keys(), key=float):
            indicators = sorted(sdg_data[sdg_num][metric_num].keys(), key=float)
            print(f"  └─ Métrica {sdg_num}.{metric_num}: Indicadores {[f'{sdg_num}.{metric_num}.{ind}' for ind in indicators]}")

if __name__ == "__main__":
    generate_formatted_excel()
