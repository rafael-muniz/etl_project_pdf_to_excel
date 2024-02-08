import os
import tabula
import pandas as pd
import xlwings as xw
import platform


# Get the current directory
current_directory = os.getcwd()

# List all files in the current directory
all_files = os.listdir(current_directory)

# Filter only PDF files
pdf_files = [file for file in all_files if file.lower().endswith('.pdf')]


data = []

# Process each PDF file
for pdf_file in pdf_files:

    ## 1) If the file is file1.pdf:
    if pdf_file == "file1.pdf":
        # Construct the full path to the PDF file
        pdf_path = os.path.join(current_directory, pdf_file)

        # Read tables from the PDF
        tables = tabula.read_pdf(pdf_path, pages='1', multiple_tables=True)

        # 1) Extracting the reference month information:
        info_column = tables[0].iloc[:, 0]
        information = info_column[9]
        month_reference = information.split()[0]
        bill_due_date = information.split()[2]
        # 2) Getting the bill's total cost:
        real_cost = information.split()[1]
        # 3) Looking for the TAX information - PIS, COFINS and ICMS:
        info_column2 = tables[1].iloc[:, 1]
        information2 = info_column2[10]
        # 3.1) ICMS tax percentage:
        icms = information2.split()[16]
        # 3.2) PIS tax percentage:
        pis = information2.split()[17]
        # 3.3) COFINS tax percentage:
        cofins = information2.split()[18]

        # 4) Extracting the values for 'Consumo Ativo' - 'TUSD' and 'TE'
        # 4.1) Consumo Ativo - TUSD:
        ca_tusd = information2.split()[9]
        # 4.2) Consumo Ativo - TE:
        ca_te = information2.split()[11]

        # 5) Extracting the client code:
        information3 = info_column[6]
        numbers = information3.split()[0]
        client_code = numbers[:-9]

        # 6) Extracting the compensated energy:
        info_column3 = tables[1].iloc[:, 0]
        information4 = info_column3[34]
        comp_energy = information4.split()[14]
        
        # Adding to dataset
        data.append({
            "Cliente": client_code,
            "Mes ref": month_reference,
            "Vencimento": bill_due_date,
            "Consumo ativo-TUSD": ca_tusd,
            "Consumo ativo-TE": ca_te,
            "ICMS": icms,
            "PIS": pis,
            "COFINS": cofins,
            "Energia compensada": comp_energy,
            "Custo real": real_cost
        })


    ## 2) If the file is file2.pdf:
    elif pdf_file == "file2.pdf":
        # Construct the full path to the PDF file
        pdf_path = os.path.join(current_directory, pdf_file)

        # Read tables from the PDF
        tables = tabula.read_pdf(pdf_path, pages='1', multiple_tables=True)
        
        # 1) Extracting the reference month information:
        info_column = tables[0].iloc[:, 0]
        information = info_column[4]
        month_reference = information.split()[0]
        bill_due_date = information.split()[2]
        # 2) Getting the bill's total cost:
        real_cost = information.split()[1]
        # 3) Looking for the TAX information - PIS, COFINS and ICMS:
        info_column2 = tables[0].iloc[:, 1]
        information2 = info_column2[10]
        # 3.1) ICMS tax percentage:
        icms = information2.split()[16]
        # 3.2) PIS tax percentage:
        pis = information2.split()[17]
        # 3.3) COFINS tax percentage:
        cofins = information2.split()[18]

        # 4) Extracting the values for 'Consumo Ativo' - 'TUSD' and 'TE'
        # 4.1) Consumo Ativo - TUSD:
        ca_tusd = information2.split()[9]
        # 4.2) Consumo Ativo - TE:
        ca_te = information2.split()[11]

        # 5) Extracting the client code:
        information3 = info_column[1]
        numbers = information3.split()[0]
        client_code = numbers[:-9]

        # 6) Extracting the compensated energy:
        information4 = info_column[35]
        comp_energy = information4.split()[14]

        # Adding to dataset
        data.append({
            "Cliente": client_code,
            "Mes ref": month_reference,
            "Vencimento": bill_due_date,
            "Consumo ativo-TUSD": ca_tusd,
            "Consumo ativo-TE": ca_te,
            "ICMS": icms,
            "PIS": pis,
            "COFINS": cofins,
            "Energia compensada": comp_energy,
            "Custo real": real_cost
        })


    ## 3) If the file is file3.pdf:
    elif pdf_file == "file3.pdf":
        # Construct the full path to the PDF file
        pdf_path = os.path.join(current_directory, pdf_file)

        # Read tables from the PDF
        tables = tabula.read_pdf(pdf_path, pages='1', multiple_tables=True)

        # 1) Extracting the reference month information:
        info_column = tables[0].iloc[:, 0]
        information = info_column[4]
        month_reference = information.split()[0]
        bill_due_date = information.split()[2]
        # 2) Getting the bill's total cost:
        real_cost = information.split()[1]
        # 3) Looking for the TAX information - PIS, COFINS and ICMS:
        info_column2 = tables[0].iloc[:, 1]
        information2 = info_column2[10]
        # 3.1) ICMS tax percentage:
        icms = information2.split()[16]
        # 3.2) PIS tax percentage:
        pis = information2.split()[17]
        # 3.3)COFINS tax percentage:
        cofins = information2.split()[18]

        # 4) Extracting the values for 'Consumo Ativo' - 'TUSD' and 'TE'
        # 4.1) Consumo Ativo - TUSD:
        ca_tusd = information2.split()[9]
        # 4.2) Consumo Ativo - TE:
        ca_te = information2.split()[11]

        # 5) Extracting the client code:
        information3 = info_column[1]
        client_code = information3.split()[3]

        # 6) Extracting the compensated energy:
        information4 = info_column[34]
        comp_energy = information4.split()[14]

        # Adding to dataset
        data.append({
            "Cliente": client_code,
            "Mes ref": month_reference,
            "Vencimento": bill_due_date,
            "Consumo ativo-TUSD": ca_tusd,
            "Consumo ativo-TE": ca_te,
            "ICMS": icms,
            "PIS": pis,
            "COFINS": cofins,
            "Energia compensada": comp_energy,
            "Custo real": real_cost
        })

    ## 4) If the file is file4.pdf:
    elif pdf_file == "file4.pdf":
        # Construct the full path to the PDF file
        pdf_path = os.path.join(current_directory, pdf_file)

        # Read tables from the PDF
        tables = tabula.read_pdf(pdf_path, pages='1', multiple_tables=True)

        # 1) Extracting the reference month information:
        info_column = tables[0].iloc[:, 0]
        information = info_column[8]
        month_reference = information.split()[0]
        bill_due_date = information.split()[2]
        # 2) Getting the bill's total cost:
        real_cost = information.split()[1]
        # 3) Looking for the TAX information - PIS, COFINS and ICMS:
        info_column2 = tables[0].iloc[:, 1]
        information2 = info_column2[14]
        # 3.1) ICMS tax percentage:
        icms = information2.split()[16]
        # 3.2) PIS tax percentage:
        pis = information2.split()[17]
        # 3.3) COFINS tax percentage:
        cofins = information2.split()[18]

        # 4) Extracting the values for 'Consumo Ativo' - 'TUSD' and 'TE'
        # 4.1) Consumo Ativo - TUSD:
        ca_tusd = information2.split()[9]
        # 4.2) Consumo Ativo - TE:
        ca_te = information2.split()[11]

        # 5) Extracting the client code:
        information3 = info_column[5]
        numbers = information3.split()[0]
        client_code = numbers[:-10]

        # 6) Extracting the compensated energy:
        information4 = info_column[38]
        comp_energy = information4.split()[14]

        # Adding to dataset
        data.append({
            "Cliente": client_code,
            "Mes ref": month_reference,
            "Vencimento": bill_due_date,
            "Consumo ativo-TUSD": ca_tusd,
            "Consumo ativo-TE": ca_te,
            "ICMS": icms,
            "PIS": pis,
            "COFINS": cofins,
            "Energia compensada": comp_energy,
            "Custo real": real_cost
        })

    else:
        print("error")


# Creation of the Data Frame
df = pd.DataFrame(data)


## Data Transformation
# 1) Converting commas to point as a decimal separator and making those variables numeric ones (float)
df["Consumo ativo-TUSD"] = df["Consumo ativo-TUSD"].str.replace(',', '.').astype(float)
df["Consumo ativo-TE"] = df["Consumo ativo-TE"].str.replace(',', '.').astype(float)
df["ICMS"] = df["ICMS"].str.replace(',', '.').astype(float)
df["PIS"] = df["PIS"].str.replace(',', '.').astype(float)
df["COFINS"] = df["COFINS"].str.replace(',', '.').astype(float)
df["Energia compensada"] = df["Energia compensada"].astype(float)
df["Custo real"] = df["Custo real"].str.replace(',', '.').astype(float)

# 2) Converting numbers into percentages
df["ICMS"] = df["ICMS"]/100
df["PIS"] = df["PIS"]/100
df["COFINS"] = df["COFINS"]/100


## Creating the Excel file and placing info in it:
# Create a new Excel workbook
wb = xw.Book()

# Access the default sheet in the workbook
sheet = wb.sheets[0]

# Placing values in specific cells
sheet.range('A1').value = df
if (df['Consumo ativo-TUSD'].nunique() == 1) and (df['Consumo ativo-TE'].nunique() == 1):
    sheet.range("A8").value = "Cons ativo-TUSD"
    sheet.range("B8").value = df["Consumo ativo-TUSD"].iloc[0]
    sheet.range("A9").value = "Cons ativo-TE"
    sheet.range("B9").value = df["Consumo ativo-TE"].iloc[0]
else:
    sheet.range("A8").value = "Tarifas diferentes, por favor checar."

if (df['ICMS'].nunique() == 1) and (df['PIS'].nunique() == 1) and (df['COFINS'].nunique() == 1):
    sheet.range("A10").value = "ICMS"
    sheet.range("B10").value = df["ICMS"].iloc[0]
    sheet.range("A11").value = "PIS"
    sheet.range("B11").value = df["PIS"].iloc[0]
    sheet.range("A12").value = "COFINS"
    sheet.range("B12").value = df["COFINS"].iloc[0]
else:
    sheet.range("A10").value = "Tributos diferentes, por favor checar."

# Construct the full path to save the workbook in the current folder
excel_file_path = os.path.join(current_directory, 'Data.xlsx')

# Save the workbook to a file
wb.save(excel_file_path)

# Determine the operating system
os_type = platform.system()

# Open the saved Excel file - using the correct command depending on which operating system is running the script
if os_type == 'Windows':
    os.system(f'start excel "{excel_file_path}"')
elif os_type == 'Darwin':  # 'Darwin' is the platform identifier for macOS
    os.system(f'open "{excel_file_path}"')
else:
    print(f"Sistema operacional sem suporte: {os_type}")