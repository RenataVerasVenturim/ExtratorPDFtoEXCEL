#seletor para indicar localização dos elementos a serem extraidos do PDF

import pdfquery

pdf = pdfquery.PDFQuery(r'C:\Users\Renata\Desktop\teste\NE_153248_2022NE000015_v002_PF1522756_20220823181624.pdf')
pdf.load()

# Exemplo de seletor para localizar um elemento com base em seu texto
elementos = pdf.pq('LTTextLineHorizontal:contains("288.000,00")')

# Loop para iterar sobre os elementos encontrados
for elemento in elementos:
    # Acessar propriedades do elemento
    left = float(elemento.attrib['x0'])
    top = float(elemento.attrib['y0'])
    width = float(elemento.attrib['width'])
    height = float(elemento.attrib['height'])
    texto = elemento.text

    # Imprimir informações do elemento
    print(f'Posição: left={left}, top={top}')
    print(f'Tamanho: width={width}, height={height}')
    print(f'Texto: {texto}')
