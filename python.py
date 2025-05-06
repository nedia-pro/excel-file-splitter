import pandas as pd
import math
import os

header_row_count = 6
real_header_row = 4  # الصف 5 في الإكسل = index 4 في بايثون
max_data_rows = max_total_rows - header_row_count


filtered_data = data_rows[data_rows[upc_column_index].astype(str).str.len() >= 12]
num_parts = math.ceil(len(filtered_data) / max_data_rows)
os.makedirs(output_folder, exist_ok=True)


thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)


    start = i * max_data_rows
    end = start + max_data_rows
    chunk = pd.concat([header_rows, filtered_data.iloc[start:end]], ignore_index=True)

    filename = f"products_part_{i+1} - الجزء_{i+1}.xlsx"
    filepath = os.path.join(output_folder, filename)
    chunk.to_excel(filepath, index=False, header=False)

    wb = load_workbook(filepath)
    ws = wb.active

            cell.font = bold_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = center_align

   
