import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
import streamlit as st


def max_row_for_col(ws, col_num):
    max_row = 1
    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=col_num).value is not None:
            max_row = row
    return max_row


def investors_conversion_macro_ben(input_csv):
    # Read the CSV file into a DataFrame
    df = pd.read_csv(input_csv)

    # Create a new Excel workbook and select the active worksheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write the DataFrame to the worksheet including headers
    for col_num, header in enumerate(df.columns, start=1):
        ws.cell(row=1, column=col_num).value = header

    for row_num, row in enumerate(df.values, start=2):
        for col_num, value in enumerate(row, start=1):
            ws.cell(row=row_num, column=col_num).value = value

    # Delete column J (10th column) and column A (1st column)
    if ws.max_column >= 10:
        ws.delete_cols(10)
    if ws.max_column >= 1:
        ws.delete_cols(1)

    # Insert 3 rows at the top of the sheet
    ws.insert_rows(1, amount=3)

    # Add the text "Client: XXX Investments" in cell A2
    ws["A2"] = "Client: XXX Investments"

    # Align the text in cell A2 to the left
    ws["A2"].alignment = Alignment(horizontal="left")

    # Auto-fit column width for column A
    max_length = 0
    column = 'A'
    for i in range(1, ws.max_row + 1):
        cell_value = ws[f"{column}{i}"].value
        if cell_value:
            max_length = max(max_length, len(str(cell_value)))
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

    # Find the last filled row in columns D, F, G, I, and K
    lastRowD = max_row_for_col(ws, 4)
    lastRowF = max_row_for_col(ws, 6)
    lastRowG = max_row_for_col(ws, 7)
    lastRowI = max_row_for_col(ws, 9)
    lastRowK = max_row_for_col(ws, 11)

    # Insert the text "Total" four cells after the last filled cell in column D
    ws.cell(row=lastRowD + 4, column=4).value = "Total"

    # Insert Average Formula in column F and keep the percentage symbol without converting values
    formulaF = f"=AVERAGE(F5:F{lastRowF})"
    formulaCellF = ws.cell(row=lastRowF + 4, column=6)
    formulaCellF.value = formulaF
    formulaCellF.number_format = '0.00"%"'

    # Insert Sum Formula in column G
    formulaG = f"=SUM(G5:G{lastRowG})"
    formulaCellG = ws.cell(row=lastRowG + 4, column=7)
    formulaCellG.value = formulaG
    formulaCellG.number_format = "0"

    # Insert Sum Formula in column I
    formulaI = f"=SUM(I5:I{lastRowI})"
    formulaCellI = ws.cell(row=lastRowI + 4, column=9)
    formulaCellI.value = formulaI
    formulaCellI.number_format = "0"

    # Insert SUMPRODUCT Formula in column K and keep the percentage symbol without converting values
    sumGCellAddress = ws.cell(row=lastRowG + 4, column=7).coordinate
    formulaK = f"=SUMPRODUCT(G5:G{lastRowG}, K5:K{lastRowK}) / {sumGCellAddress}"
    formulaCellK = ws.cell(row=lastRowK + 4, column=11)
    formulaCellK.value = formulaK
    formulaCellK.number_format = '0.00"%"'

    # Set column widths from B to L to 15 and center-align all cells from A to L
    for col in range(2, 13):
        ws.column_dimensions[get_column_letter(col)].width = 15
    for row in ws.iter_rows(min_col=1, max_col=12, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


st.title("Investor Breakdown Conversion Final Output_BM")

uploaded_file = st.file_uploader("Upload a CSV file", type=["csv"])

if uploaded_file:
    output_file = investors_conversion_macro_ben(uploaded_file)

    st.download_button(
        label="Download Modified Excel file",
        data=output_file.getvalue(),
        file_name="modified_excel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
