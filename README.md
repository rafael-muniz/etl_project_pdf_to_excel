# ETL project - PDF to Excel

**<ins>ETL project in Python - Data from PDF file to Excel</ins>**

The code is essentially a data extraction and processing script for PDF files related to energy bill information, with the final result being an Excel file containing the extracted data.<br>
It is an ETL project for Cosern's enegy bill (a Brazilian energy company situated in the Northeast of Brazil - in the State of Rio Grande do Norte), for customers served in low voltage. <br><br>
This Python project aimed to streamline a monthly task, where an individual had to manually open each PDF document, extract specific values, and perform calculations. With this script, the individual can now automate this process, saving significant time and effort. By compiling all desired information into a single Excel file, the individual can efficiently review and copy the required data, eliminating the need for repetitive manual tasks. This approach enhances efficiency, allowing for a more streamlined and effective workflow, ultimately saving valuable time and resources. <br><br>

**<ins>This code performs the following tasks:</ins>**
1) It gets the current directory and lists all files in it, filtering only PDF files.
2) It defines an empty list data to store extracted information.
3) It iterates over each PDF file in the directory and performs different operations based on the file name.
4) For each PDF file, it extracts specific information such as reference month, bill due date, total cost, tax percentages (ICMS, PIS, COFINS), active consumption (TUSD and TE), client code, and compensated energy.
5) It appends this information as dictionaries to the data list.
6) After processing all PDF files, it creates a pandas DataFrame from the collected data.
7) It transforms certain columns in the DataFrame, converting comma-separated numbers to float and converting values to percentages.
8) It creates an Excel file using xlwings library, places the DataFrame into the file, and adds some additional information in specific cells.
9) It determines the operating system and opens the saved Excel file accordingly.
<br><br>
Version 1 - 07/feb/2024
------------------------------------------------------------------------------
<br>(In Portuguese)<br><br>
**<ins>Projeto ETL em Python - Dado de um arquivo PDF para o Excel</ins>**

O código é essencialmente um script de extração e processamento de dados de arquivos PDF relacionados às informações da conta de energia, tendo como resultado final um arquivo Excel contendo os dados extraídos.<br>
É um projeto de ETL da conta de energia da Cosern (distribuidora de energia do Rio Grande do Norte), para clientes atendidos em baixa tensão.<br><br>
Esse projeto em Python teve como objetivo agilizar uma tarefa mensal, onde uma pessoa tinha que abrir manualmente cada documento PDF, extrair valores específicos e realizar cálculos. Com esse script, a pessoa agora pode automatizar esse processo, economizando tempo e esforço significativos. Ao compilar todas as informações desejadas em um único arquivo Excel, a pessoa pode revisar e copiar com eficiência os dados necessários, eliminando a necessidade de tarefas manuais repetitivas. Essa abordagem aumenta a eficiência, permitindo um fluxo de trabalho mais simplificado e eficaz, economizando tempo e recursos valiosos. <br><br>

**<ins>Esse código executa as seguintes tarefas:</ins>**
1) Obtém o diretório atual e lista todos os arquivos nele contidos, filtrando apenas os arquivos PDF.
2) Define uma lista vazia de dados para armazenar as informações extraídas.
3) Ele interage com cada arquivo PDF no referido diretório e executa diferentes operações com base no nome do arquivo.
4) Para cada arquivo PDF extrai informações específicas como mês de referência, data de vencimento da fatura, custo total, percentuais de impostos (ICMS, PIS, COFINS), consumo ativo (TUSD e TE), código do cliente, etc.
5) Anexa essas informações como dicionários à lista de dados.
6) Após processar todos os arquivos PDF, ele cria um DataFrame do pandas a partir dos dados coletados.
7) Efetua transformações certas colunas do DataFrame, corrigindo o separador de decimais, convertendo "string" em "float" e convertendo outros valores em porcentagens.
8) Ele cria um arquivo Excel usando a biblioteca xlwings, coloca o DataFrame no arquivo e adiciona algumas informações adicionais em células específicas, para facilitar a vida da pessoa que irá usar os dados.
9) Determina o sistema operacional e abre o arquivo Excel de acordo com o sistema.
<br><br>

Versão 1 - 07/fev/2024
