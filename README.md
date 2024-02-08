# ETL project - PDF to Excel
### ETL project - Data from PDF file to Excel

The code is essentially a data extraction and processing script for PDF files related to energy bill information, with the final result being an Excel file containing the extracted data.
It is an ETL project for Cosern's enegy bill (a Brazilian energy company situated in the Northeast of Brazil - in the State of Rio Grande do Norte), for customers served in low voltage.

**<ins>This code performs the following tasks:</ins>**
1) It gets the current directory and lists all files in it, filtering only PDF files.
2) It defines an empty list data to store extracted information.
3) It iterates over each PDF file in the directory and performs different operations based on the file name.
4) For each PDF file, it extracts specific information such as reference month, bill due date, total cost, tax percentages (ICMS, PIS, COFINS), active consumption (TUSD and TE), client code, and compensated energy.
5) It appends this information as dictionaries to the data list.
6) After processing all PDF files, it creates a pandas DataFrame from the collected data.
7) It transforms certain columns in the DataFrame, converting comma-separated numbers to float and converting percentages to actual values.
8) It creates an Excel file using xlwings, places the DataFrame into the file, and adds some additional information in specific cells.
9) It determines the operating system and opens the saved Excel file accordingly.

------------------------------------------------------------------------------
(In Portuguese)
