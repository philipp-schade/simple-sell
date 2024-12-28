import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def format_excel_file(input_file, output_file):
    # Load the unformatted Excel file
    #input_file = "02_unformatted.xlsx"
    #output_file = "NEW_FORMATTED.xlsx"

    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active


    # Keep columns A, E, G and remove others
    columns_to_keep = [1, 5, 7]  # Corresponding to A, E, G
    all_columns = list(range(1, sheet.max_column + 1))
    columns_to_delete = [col for col in all_columns if col not in columns_to_keep]

    for col in sorted(columns_to_delete, reverse=True):
        for row in sheet.iter_rows():
            row[col - 1].value = None

    # Center column "order_item_quantity" (Assumed as Column B after deletions)
    for row in range(1, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=2)  # Access the second column (B)
        cell.alignment = Alignment(horizontal="center")

    # Filter "order_item_sku" entries with "4S" and format as bold
    for row in range(1, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=7)  # Access Column G
        if cell.value and "4S" in str(cell.value):
            cell.font = Font(bold=True)

    # Filter "order_item_sku" entries with "6S" and format as italics and underlined
    for row in range(1, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=7)  # Access Column G
        if cell.value and "6S" in str(cell.value):
            cell.font = Font(italic=True, underline="single")

    # Filter "order_item_sku" entries with "-2-" and format as bold
    for row in range(1, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=7)  # Access Column G
        if cell.value and "-2-" in str(cell.value):
            cell.font = Font(bold=True)

    # Highlight rows with "order_item_quantity" > 1 in light blue
    for row in range(2, sheet.max_row + 1):  # Skip header row
        quantity_cell = sheet.cell(row=row, column=5)  # Column E
        if quantity_cell.value and quantity_cell.value > 1:
            fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row=row, column=col).fill = fill

    # Draw a bold box around rows with the same reference code (Column C after deletions)
    reference_groups = {}
    for row in range(2, sheet.max_row + 1):
        reference_code = sheet.cell(row=row, column=1).value  # Column C for reference code
        if reference_code not in reference_groups:
            reference_groups[reference_code] = []
        reference_groups[reference_code].append(row)

    for group in reference_groups.values():
        if len(group) > 1:
            for col in range(1, sheet.max_column + 1):
                # Apply top border to the first row in the group
                top_cell = sheet.cell(row=group[0], column=col)
                top_cell.border = Border(
                    top=Side(border_style="thick"),
                    left=top_cell.border.left,
                    right=top_cell.border.right,
                    bottom=top_cell.border.bottom
                )

                # Apply bottom border to the last row in the group
                bottom_cell = sheet.cell(row=group[-1], column=col)
                bottom_cell.border = Border(
                    bottom=Side(border_style="thick"),
                    left=bottom_cell.border.left,
                    right=bottom_cell.border.right,
                    top=bottom_cell.border.top
                )

            for row in group:
                for col in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=row, column=col)
                    if col == 1:  # Apply left border
                        cell.border = Border(
                            left=Side(border_style="thick"),
                            top=cell.border.top,
                            right=cell.border.right,
                            bottom=cell.border.bottom
                        )
                    if col == sheet.max_column:  # Apply right border
                        cell.border = Border(
                            right=Side(border_style="thick"),
                            top=cell.border.top,
                            left=cell.border.left,
                            bottom=cell.border.bottom
                        )

    # Insert bold line after every new reference code
    for row in range(2, sheet.max_row):
        current_code = sheet.cell(row=row, column=1).value
        next_code = sheet.cell(row=row + 1, column=1).value
        if current_code != next_code:
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row + 1, column=col)
                cell.border = Border(
                    top=Side(border_style="thick")
                )

    # Delete columns
    sheet.delete_cols(2, 3)
    sheet.delete_cols(3, 1)
    sheet.delete_cols(4, 1)
    
    # Add a filter to the sheet
    sheet.auto_filter.ref = sheet.dimensions

    # Save the formatted file
    wb.save(output_file)

