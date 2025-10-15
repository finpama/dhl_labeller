import pandas as pd

from datetime import date
import os

from reportlab.lib.pagesizes import A4

import re
import pdfplumber
from util import unir_pdfs

# Caminho dos arquivos
diretorio_pdfs = "./Leitor de CTE"

# Dicionário com os padrões
padroes = {
    "TIPO": r"\b(Normal|Complemento)\b", # Encontra o tipo da CT-e
    "FORNECEDOR": r"\b(CURITIBA|SAO PAULO|BRASILIA)\b",  # Extrai uma das cidades específicas
    "Nº CTE": r"N[ºº]?[^\d]*([\d\.]+)",  # Extrai número do CTE após "Nº"
    "VALOR": r"RECEBER[^\d]*([\d\.,]+)",  # Extrai valor após a palavra "RECEBER"
    "EMISSÃO": r"\b(\d{2}/\d{2}/\d{4})\b",  # Extrai datas no formato dd/mm/aaaa
    "AWB": r"\bW[B8][^\d]*(\d{6,})\b"  # Extrai número após "WB" ou "W8"
}

print('\n\nAguarde...\n')

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

    unir_pdfs('./Leitor de CTE', f'./CTES.pdf', arquivos)

    print('Gerado o arquivo unificado "CTES.pdf" e o "relatório_cte.xlsx" com os pdfs na pasta Leitor de CTE')