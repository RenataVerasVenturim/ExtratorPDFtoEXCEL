#extrair dados dos empenhos pdf de uma pasta e enviar para uma planilha excel (planilhadas e organizadas)

import os
import pdfquery
import openpyxl

directory_path = r'C:\Users\Renata\Desktop\teste'
extension = '.pdf'
excel_path = r'C:\Users\Renata\Desktop\teste\Planilhar.xlsx'

# Definir as coordenadas dos elementos desejados
coordinates = [
    {'left': 200.0, 'top': 549.52, 'width': 16.68, 'height': 10.0},  # Coordenadas do final do número de empenho
    {'left': 41.0, 'top': 418.52, 'width': 374.62, 'height': 10.0},  # Coordenadas do fornecedor (nome e CNPJ) da nota de empenho
    {'left': 421.0, 'top': 642.52, 'width': 50.02, 'height': 10.0},   # Coordenadas do valor da nota de empenho 
    {'left': 200.0, 'top': 464.52, 'width': 139.57, 'height': 10.0},   # Coordenadas do número do processo
    {'left': 200.0, 'top': 503.52, 'width': 55.6, 'height': 10.0},     # Coordenadas da fonte de despesa
    {'left': 43.0, 'top': 627.52, 'width': 387.29, 'height': 10.0} ,   # Coordenadas da natureza da despesa
    {'left': 125.0, 'top': 306.52, 'width': 122.66, 'height': 10.0} ,   # Modalidade da licitação
    {'left': 407.0, 'top': 503.52,'width': 33.36, 'height': 10.0} ,   # UGR
]

 
# Obter lista de arquivos PDF no diretório
pdf_files = [file for file in os.listdir(directory_path) if file.endswith(extension)]

# Abrir o arquivo Excel
workbook = openpyxl.load_workbook(excel_path)

# Selecionar a planilha desejada
sheet = workbook['EMPENHOS']

# Definir a coluna inicial para a inserção dos resultados
start_column = 1

# Iterar sobre os arquivos PDF
for i, pdf_file in enumerate(pdf_files):
    # Caminho completo para o arquivo PDF
    pdf_path = os.path.join(directory_path, pdf_file)

    # Inicializar o objeto PDFQuery
    pdf = pdfquery.PDFQuery(pdf_path)
    pdf.load()

    # Iterar sobre as coordenadas dos elementos
    for j, coord in enumerate(coordinates):
        target_left = coord['left']
        target_top = coord['top']
        target_width = coord['width']
        target_height = coord['height']

        # Procurar o elemento com base nas coordenadas
        element = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (target_left, target_top, target_left + target_width, target_top + target_height))

        # Extrair o texto do elemento encontrado
        text = element.text().strip()

        # Inserir o texto na célula correspondente
        sheet.cell(row=i+1, column=start_column+j).value = text

# Salvar as alterações no arquivo Excel
workbook.save(excel_path)

print("Dados inseridos na planilha.")
