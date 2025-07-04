


filtered_data = data_rows[data_rows[upc_column_index].astype(str).str.len() >= 12]
num_parts = math.ceil(len(filtered_data) / max_data_rows)
os.makedirs(output_folder, exist_ok=True)


thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

    filename = f"products_part_{i+1} - _{i+1}.xlsx"
    filepath = os.path.join(output_folder, filename)
    chunk.to_excel(filepath, index=False, header=False)

    wb = load_workbook(filepath)
    ws = wb.active

            cell.font = bold_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = center_align

   
