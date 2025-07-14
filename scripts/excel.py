import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import re

def clean_sheet_name(name):
    """Clean sheet name to be valid for Excel."""
    name = re.sub(r'[\\/*?:"<>|]', '', str(name))
    if len(name) > 31:
        name = name[:28] + "..."
    return name

def extract_sdg_metric_indicator(indicator_ref):
    """Extract SDG, metric, and indicator from reference (e.g., indicator_1-1-1)."""
    if pd.isna(indicator_ref):
        return "Unknown", "Unknown", indicator_ref
    
    indicator_str = str(indicator_ref).strip()
    match = re.match(r'indicator_(\d+)-(\d+)-(\d+)?', indicator_str)
    if match:
        sdg = f"SDG{match.group(1)}"
        metric = f"{match.group(1)}.{match.group(2)}"
        indicator = f"{match.group(1)}.{match.group(2)}.{match.group(3)}" if match.group(3) else metric
        return sdg, metric, indicator
    return "Unknown", "Unknown", indicator_str

def create_sdg_sheet(wb, sdg_name, sdg_data, sdg_title):
    """Create a single sheet for an SDG with all its metrics and indicators."""
    sheet_name = clean_sheet_name(sdg_name)
    ws = wb.create_sheet(title=sheet_name)
    
    # --- Main Header ---
    title = f"{sdg_name}: {sdg_title}"
    ws.cell(row=1, column=1, value=title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    
    title_cell = ws.cell(row=1, column=1)
    title_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    title_cell.font = Font(color="FFFFFF", bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # --- Table Headers ---
    headers = [
        "Type",
        "Metric and indicator reference",
        "Metric / Indicator",
        "Value\n(for continuous data)",
        "Yes/No",
        "Evidence",
        "Public\n(Yes/No)"
    ]
    
    for col, header in enumerate(headers, start=1):
        ws.cell(row=2, column=col, value=header)
    
    header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=2, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # --- Add Data ---
    current_row = 3
    column_mapping = {
        "Type": ["Type", "type", "Tipo"],
        "Metric and indicator reference": ["Metric and indicator reference", "Reference", "Referencia", "Ref", "Indicator"],
        "Metric / Indicator": ["Metric / Indicator", "Metric", "Indicator", "Indicador", "Metrica"],
        "Value\n(for continuous data)": ["Value", "Valor", "Value (for continuous data)", "Value\n(for continuous data)"],
        "Yes/No": ["Yes/No", "YesNo", "Sim/Não", "Status"],
        "Evidence": ["Evidence", "Evidencia", "Proof"],
        "Public\n(Yes/No)": ["Public", "Publico", "Public (Yes/No)", "Public\n(Yes/No)"]
    }
    
    # Group by metric to ensure metrics are listed before their indicators
    sdg_data['metric'] = sdg_data['Metric and indicator reference'].apply(lambda x: extract_sdg_metric_indicator(x)[1])
    grouped = sdg_data.groupby('metric', sort=False)
    
    for metric, group_data in grouped:
        # Add metric row
        metric_row = group_data.iloc[0]
        row_values = ["Metric", metric, metric_row.get("Metric / Indicator", metric)]
        for header in headers[3:]:
            value = ""
            for possible_col in column_mapping.get(header, [header]):
                if possible_col in metric_row:
                    value = metric_row[possible_col]
                    break
            if value == "" and "\n" in header:
                clean_header = header.replace("\n", " ")
                if clean_header in metric_row:
                    value = metric_row[clean_header]
            row_values.append(value)
        
        for col, value in enumerate(row_values, start=1):
            ws.cell(row=current_row, column=col, value=value)
        
        current_row += 1
        
        # Add indicator rows
        for _, row_data in group_data.iterrows():
            sdg, metric, indicator = extract_sdg_metric_indicator(row_data["Metric and indicator reference"])
            if indicator != metric:  # Only add as indicator if it's not the metric itself
                row_values = ["Indicator", indicator, row_data.get("Metric / Indicator", indicator)]
                for header in headers[3:]:
                    value = ""
                    for possible_col in column_mapping.get(header, [header]):
                        if possible_col in row_data:
                            value = row_data[possible_col]
                            break
                    if value == "" and "\n" in header:
                        clean_header = header.replace("\n", " ")
                        if clean_header in row_data:
                            value = row_data[clean_header]
                    row_values.append(value)
                
                for col, value in enumerate(row_values, start=1):
                    ws.cell(row=current_row, column=col, value=value)
                
                current_row += 1
    
    # --- Formatting ---
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    
    for row in ws.iter_rows(min_row=1, max_row=current_row-1, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
    
    col_widths = [12, 12, 50, 15, 12, 12, 12]
    for idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    for row in range(1, current_row):
        ws.row_dimensions[row].height = 25
    
    for row in range(3, current_row):
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

def generate_formatted_excel():
    excel_dir = Path('scripts/excel')
    excel_dir.mkdir(parents=True, exist_ok=True)
    
    data_dir = Path('data')
    if not data_dir.exists():
        print(f"Error: The folder {data_dir} does not exist!")
        return
    
    csv_files = list(data_dir.glob('*.csv'))
    if not csv_files:
        print(f"No CSV files found in {data_dir}")
        return
    
    print(f"Found {len(csv_files)} CSV files in {data_dir}")
    
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    total_sheets = 0
    sdg_dataframes = {}
    
    # SDG titles (you can expand this dictionary as needed)
    sdg_titles = {
        "SDG1": "End poverty in all its forms everywhere",
        "SDG2": "Zero hunger",
        # Add more SDG titles as needed
    }
    
    for csv_file in csv_files:
        try:
            print(f"Processing: {csv_file.name}")
            df = pd.read_csv(csv_file)
            
            ref_column = None
            for col in ["Metric and indicator reference", "Reference", "Referencia", "Ref", "Indicator"]:
                if col in df.columns:
                    ref_column = col
                    break
            
            if ref_column is None:
                print(f"  → WARNING: Reference column not found. Using index.")
                df['temp_ref'] = df.index.astype(str)
                ref_column = 'temp_ref'
            
            df[['sdg', 'metric', 'indicator']] = df[ref_column].apply(extract_sdg_metric_indicator).apply(pd.Series)
            
            for sdg, group_data in df.groupby('sdg'):
                if sdg != "Unknown":
                    if sdg in sdg_dataframes:
                        sdg_dataframes[sdg] = pd.concat([sdg_dataframes[sdg], group_data], ignore_index=True)
                    else:
                        sdg_dataframes[sdg] = group_data
                    
                    print(f"  → SDG: {sdg}, Rows: {len(group_data)}")
            
        except Exception as e:
            print(f"Error processing {csv_file.name}: {e}")
            import traceback
            traceback.print_exc()
    
    for sdg, data in sdg_dataframes.items():
        sdg_title = sdg_titles.get(sdg, "Unknown SDG")
        create_sdg_sheet(wb, sdg, data, sdg_title)
        total_sheets += 1
        print(f"  → Sheet created: {sdg} ({len(data)} rows)")
    
    if total_sheets == 0:
        print("No sheets were created!")
        return
    
    excel_path = excel_dir / 'sdg-data-formatado.xlsx'
    wb.save(excel_path)
    print(f"\nFormatted Excel generated successfully: {excel_path}")
    print(f"Total sheets created: {total_sheets}")

if __name__ == "__main__":
    generate_formatted_excel()
