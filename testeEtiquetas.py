import re
from datetime import datetime

import pdfplumber
import pandas as pd

from gerarEtiquetas import gerar_base

caminho_pdf = './DHL 2-10-2025.pdf'
dataDesejada = '02/10/2025'

# Dicionário com os padrões
padroes = {
    "CT-E": r"N[ºº]?[^\d]*([\d\.]+)",  # Extrai número do CTE após "Nº"
    "Pedido": r"Pedido:\s*(\d+)",  # Extrai o número após "Pedido:"
}

# Lista para armazenar os dados extraídos de cada PDF
dados_extraidos = []

data = datetime.strptime(dataDesejada, '%d/%m/%y').date()

# Abre o PDF
with pdfplumber.open(caminho_pdf) as pdf:

    # Percorre todas as páginas do PDF
    for i in range(len(pdf.pages)):

        pagina = pdf.pages[i]
        texto = pagina.extract_text()


        if texto:
            # Dicionário para armazenar os dados de uma página
            dados_pdf = {"Página": i + 1}

            # Para cada campo definido, aplica a regex e salva o resultado
            for campo, padrao in padroes.items():
                match = re.search(padrao, texto, re.IGNORECASE)

                valorEncontrado = match.group(1).strip() if match else ""

                if valorEncontrado != "":
                    dados_pdf[campo] = valorEncontrado
                else:
                    dados_pdf = {"Página": i + 1, "Pedido":"n/a"}

        else:
            dados_pdf = {"Página": i + 1, "Pedido":-1}


        dados_extraidos.append(dados_pdf)


df_resultado = pd.DataFrame(dados_extraidos) #.drop('Página', axis=1)

df_planilha = gerar_base('./Medições DHL.xlsm', str(data))
df_planilha = df_planilha.drop('DATA', axis=1).reset_index(drop=True)


# Converter as colunas 'CT-E' e 'Pedido' em df_resultado para numérico (int), lidando com erros
df_resultado['CT-E'] = pd.to_numeric(df_resultado['CT-E'], errors='coerce').astype('Int64')
df_resultado['Pedido'] = pd.to_numeric(df_resultado['Pedido'], errors='coerce').astype('Int64')


# Converter as colunas 'CT-E' e 'Pedido' em df_planilha para int
df_planilha[['CT-E', 'Pedido']] = df_planilha[['CT-E', 'Pedido']].astype(int)

if len(df_planilha) == len(df_resultado):
    for i in range(len(df_planilha)):
        pag = df_resultado['Página'].iloc[i]

        if df_resultado['Pedido'].iloc[i] != -1:
            if df_planilha['Pedido'].iloc[i] == df_resultado['Pedido'].iloc[i]:
                print(f'Página {pag}: Pedido Correto')
            else:
                print(f'Página {pag}: Pedido NÃO está de acordo com planilha')
        else:
            print(f'Página {pag}: Pedido NÃO ENCONTRADO')
else:
    print("Erro!")
    print(f"A planilha tem {len(df_planilha)} linhas mas o arquivo DHL tem {len(df_resultado)} CT-Es")
