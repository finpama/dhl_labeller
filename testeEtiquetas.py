import re
from datetime import datetime
import os
import sys

import pdfplumber
import pandas as pd

from gerarEtiquetas import gerar_base

def main(caminho_saidaRelatorio, caminho_pdf, caminho_medicoes, dataDesejada):
    
    buscaPor = "CT-E"

    # Dicionário com os padrões
    padroes = {
        "CT-E": r"N[ºº]?[^\d]*([\d\.]+)",  # Extrai número do CTE após "Nº"
        "Pedido": r"Pedido:\s*(\d+)",  # Extrai o número após "Pedido:"
    }

    # Lista para armazenar os dados extraídos de cada PDF
    dados_extraidos = []

    data = datetime.strptime(dataDesejada, '%d/%m/%Y').date()

    # Abre o PDF
    with pdfplumber.open(caminho_pdf) as pdf:

        # Percorre todas as páginas do PDF
        for i in range(len(pdf.pages)):

            pagina = pdf.pages[i]
            texto = pagina.extract_text()

            dados_pdf:dict = {"Página": i + 1}

            if texto:
                # Dicionário para armazenar os dados de uma página

                # Para cada campo definido, aplica a regex e salva o resultado
                for campo, padrao in padroes.items():
                    match = re.search(padrao, texto, re.IGNORECASE)

                    valorEncontrado = match.group(1).strip() if match else ""

                    if valorEncontrado:
                        dados_pdf[campo] = valorEncontrado
                    else:
                        dados_pdf[campo] = "<não-encontrado>"

            else:
                dados_pdf = {"Página": i + 1, "CT-E":"<página-ilegível>", "Pedido":"<página-ilegível>"}


            dados_extraidos.append(dados_pdf)


    df_pdf = pd.DataFrame(dados_extraidos) #.drop('Página', axis=1)

    df_planilha = gerar_base(caminho_medicoes, str(data))
    df_planilha = df_planilha.drop('DATA', axis=1).reset_index(drop=True)


    # Converter as colunas 'CT-E' e 'Pedido' em df_pdf para numérico (int), lidando com erros
    df_pdf['CT-E'] = pd.to_numeric(df_pdf['CT-E'], errors='coerce').astype('Int64')
    df_pdf['Pedido'] = pd.to_numeric(df_pdf['Pedido'], errors='coerce').astype('Int64')

    df_pdf.dropna(inplace=True)
    df_pdf.reset_index(inplace=True)

    # Converter as colunas 'CT-E' e 'Pedido' em df_planilha para int
    df_planilha[['CT-E', 'Pedido']] = df_planilha[['CT-E', 'Pedido']].astype(int)

    df_resultado = df_planilha.copy()

    if buscaPor == "Pedido":

        df_resultado["Página"] = "Pedido não encontrado"
        df_resultado["Núm. Pedidos Encontrados"] = 0

        # df.loc[row_indexer, "col"] 

        for i in range(len(df_resultado)):
            Pedido = df_resultado.loc[i, "Pedido"]
            nPedido = 0

            for j in range(len(df_pdf)):
                if Pedido == df_pdf.loc[j, "Pedido"]:
                    nPedido += 1
                    df_resultado.loc[i, "Página"] = df_pdf.loc[j, "Página"]
            
            df_resultado.loc[i, "Núm. Pedidos Encontrados"] = nPedido
    else: # buscaPor == "CT-E"

        df_resultado["Página"] = "CT-E não encontrada"
        df_resultado["Núm. CT-Es Encontradas"] = 0

        for i in range(len(df_resultado)):
            cte = df_resultado.loc[i, "CT-E"]
            nCte = 0

            for j in range(len(df_pdf)):
                if cte == df_pdf.loc[j, "CT-E"]:
                    nCte += 1
                    df_resultado.loc[i, "Página"] = df_pdf.loc[j, "Página"]
            
            df_resultado.loc[i, "Núm. CT-Es Encontradas"] = nCte

    print('\n')

    try:
        df_resultado.to_excel(f'{caminho_saidaRelatorio}/relatorioErros_{buscaPor}.xlsx')
    except PermissionError:
        input(f'O arquivo "relatorioErros_{buscaPor}.xlsx" já existe na pasta, após remover, pressione qualquer tecla para continuar.')
        df_resultado.to_excel(f'{caminho_saidaRelatorio}/relatorioErros_{buscaPor}.xlsx')
    
    print(f'Teste de etiquetas concluído e disponível no arquivo: "relatorioErros_{buscaPor}.xlsx"')


if __name__ == '__main__':
    try:
        main(sys.argv[1], sys.argv[2])
    except IndexError:
        raise TypeError('É necessário dois parâmetros para executar o testeEtiquetas.py: "caminho_pdf" e "dataDesejada"')
    
    