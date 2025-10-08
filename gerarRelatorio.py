# Célula de Configuração

import pandas as pd

from datetime import date
import os

from PyPDF2 import PdfReader, PdfWriter

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# Puxa pedidos, cria etiquetas e insere na pasta "hoje"

def gerar_etiqueta(output_path, signer, pedido):
    '''Cria PDF de etiqueta com o <pedido> informado e caso informado G ou M no campo <signer>
       Adiciona a etiqueta de assinatura do Gustavo e da Michelle respectivamente.
       O <output_path> é o nome e caminho do pdf final
    '''

    c = canvas.Canvas(output_path, pagesize=A4)

    def cria_linha(lateral, altura, txt, linha=1):
        if linha > 0 : linha-= 1 # linha--
        cord_altura = altura - 12 * linha

        c.drawString(lateral, cord_altura, txt)

    x = 100
    y = 290

    cria_linha(x, y, f'Pedido: {pedido}')
    cria_linha(x, y, 'Nat: 311018',2)
    cria_linha(x, y, 'IC: 11801', 3)
    cria_linha(x, y, 'CC: 220109', 4)

    if signer.upper() == 'G':
        cria_linha(x+250, y, 'Gustavo Carvalho Daquano', 3)
        cria_linha(x+250, y, '       Export Supervisor', 4)

    elif signer.upper() == 'M':
        cria_linha(x+250, y, 'Michelle Corrêa Lisboa', 3)
        cria_linha(x+250, y, '      Export Manager', 4)


    c.save()

# Gera dataframe com pedidos

def gerar_base(sheet_dir, dateEmission = str(date.today())):
    '''A partir da planilha <sheet_dir>, gera o dataframe com os pedidos do dia em <dateEmission> (se omitido, usa data de hoje)'''

    base = pd.read_excel(sheet_dir)

    base = base.get(['DATA', 'CT-E','Pedido'])

    base = base.dropna()

    base = base.drop(base[base.DATA != dateEmission].index)

    return base



# Mescla de PDF

def merge_pdfs_in_directory(directory_path, output_filename, pdf_files_order):
    """
    Merges PDF files in a specified directory into a single PDF file, preserving the order
    specified in pdf_files_order.

    Args:
        directory_path (str): The path to the directory containing the PDF files.
        output_filename (str): The name of the output merged PDF file.
        pdf_files_order (list): A list of PDF filenames in the desired merging order.
    """
    writer = PdfWriter()

    if not pdf_files_order:
        print(f"No PDF files specified for merging in directory: {directory_path}")
        return

    for pdf_file in pdf_files_order:
        pdf_path = os.path.join(directory_path, pdf_file)
        if not os.path.exists(pdf_path):
            print(f"Warning: File not found and skipped: {pdf_file}")
            continue
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"Could not read or process file {pdf_file}: {e}")

    # Write the merged PDF to the output file
    with open(output_filename, "wb") as output_file:
        writer.write(output_file)

    print(f"\nTodos os {len(pdf_files_order)} arquivos foram mesclados e estão no arquivo {output_filename}")

import os
import re
import pandas as pd
import pdfplumber

# Caminho dos arquivos
diretorio_pdfs = "./Leitor de CTE"

# Dicionário com os padrões
padroes = {
    "FORNECEDOR": r"\b(CURITIBA|SAO PAULO|BRASILIA)\b",  # Extrai uma das cidades específicas
    "Nº CTE": r"N[ºº]?[^\d]*([\d\.]+)",  # Extrai número do CTE após "Nº"
    "VALOR": r"RECEBER[^\d]*([\d\.,]+)",  # Extrai valor após a palavra "RECEBER"
    "EMISSÃO": r"\b(\d{2}/\d{2}/\d{4})\b",  # Extrai datas no formato dd/mm/aaaa
    "AWB": r"\bW[B8][^\d]*(\d{6,})\b"  # Extrai número após "WB" ou "W8"
}

print('Aguarde...')

# Lista para armazenar os dados extraídos de cada PDF
dados_extraidos = []

arquivos = os.listdir(diretorio_pdfs)
arquivos.sort()

if arquivos == []:
    print('Nenhum arquivo na pasta "./Leitor de CTE"')
else:

    # Percorre todos os arquivos no diretório
    for arquivo in arquivos:
        if arquivo.lower().endswith(".pdf"):
            caminho_pdf = os.path.join(diretorio_pdfs, arquivo)

            # Abre o PDF
            with pdfplumber.open(caminho_pdf) as pdf:
                texto_total = ""

                # Percorre todas as páginas do PDF
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if texto:
                        texto_total += texto + "\n"

                # Dicionário para armazenar os dados de um PDF específico
                dados_pdf = {"ARQUIVO": arquivo}

                # Para cada campo definido, aplica a regex e salva o resultado
                for campo, padrao in padroes.items():
                    match = re.search(padrao, texto_total, re.IGNORECASE)
                    dados_pdf[campo] = match.group(1).strip() if match else ""

                dados_extraidos.append(dados_pdf)

    # Cria um DataFrame com os dados extraídos
    df_resultado = pd.DataFrame(dados_extraidos)

    # Salva o DataFrame em um arquivo Excel
    df_resultado.to_excel("relatório_cte.xlsx", index=False)

    merge_pdfs_in_directory('./Leitor de CTE', f'./CTES.pdf', arquivos)

    print('Gerado o arquivo unificado "CTES.pdf" e o "relatório_cte.xlsx" com os pdfs na pasta Leitor de CTE')