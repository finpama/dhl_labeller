# Célula de Configuração

from datetime import datetime
import os
from _util import rmPasta
import re

import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pdfplumber

from PyQt6.QtWidgets import QMessageBox

from _util import gerar_base, gerar_etiqueta, unir_pdfs
import testeEtiquetas


def main(input_signer:str, input_data:str, caminhoArquivoUnico:str, caminhoMedicoes:str, caminhoSaida:str):
    
    if input_data != '':
        data = datetime.strptime(input_data, '%d/%m/%Y').date()
    else:
        data = datetime.today().date()

    data_formatada = f'{data.day}-{data.month}-{data.year}'

    caminho_saida = f'{caminhoSaida}/DHL {data_formatada}.pdf'
    
    # Para cada pedido, gera a etiqueta se tiver algo nos pedidos da planilha "Medições DHL.xlsx" no dia inserido no input. Depois sobrepor as CTEs com as etiquetas no arquivo "DHL <data do input>.pdf"'''

    df_planilha = gerar_base(caminhoMedicoes, str(data))
    df_planilha[['CT-E', 'Pedido']] = df_planilha[['CT-E', 'Pedido']].astype(int)

    pdf_order = []
    ctes_semPedido = 0


    padrao = r"N[ºº]?[^\d]*([\d\.]+)"  # Regex que extrai número do CTE após "Nº"


    # se a df_planilha tiver pedidos, gera uma etiqueta pra cada pedido em uma pasta temporária
    if df_planilha.Pedido.empty:
        QMessageBox.critical(None, 'Erro...', f'Nenhum pedido de {data_formatada} na planilha')
        return

    print('\nAguarde...')

    if not os.path.exists(f'./{data_formatada}'):
        os.mkdir(f'./{data_formatada}')
    else:
        rmPasta(f'./{data_formatada}')
        os.mkdir(f'./{data_formatada}')

    df_planilha.reset_index(inplace=True)

    # Abre o PDF

    with pdfplumber.open(caminhoArquivoUnico) as pdf:

        # Percorre todas as páginas do PDF
        for i in range(len(pdf.pages)):

            pagina = pdf.pages[i]
            texto = pagina.extract_text()

            if not texto:
                raise Exception(f"ERRO: Nenhum texto encontrado no arquivo {caminhoArquivoUnico}")

            # Aplica a regex
            match = re.search(padrao, texto, re.IGNORECASE)

            cte = int(match.group(1).strip()) if match else ""

            i_cte = df_planilha.index[df_planilha['CT-E'] == cte]

            try:
                pedido = int(df_planilha['Pedido'].iloc[i_cte].iloc[0])
            except:
                QMessageBox.critical(None, 'Erro...', f'Pág. {i+1}: Ocorreu um erro ao procurar a CT-e, favor preencher manualmente...')
                
                pedido = ""
                ctes_semPedido += 1

            n_pagina = f'Pág.{i+1}'

            nome_etiqueta = f'{n_pagina} Etiqueta.pdf'
            nome_cte_mesclada = f'{n_pagina} Mesclada.pdf'

            path_etiqueta = f'./{data_formatada}/{nome_etiqueta}'
            path_cte_mesclada = f'./{data_formatada}/{nome_cte_mesclada}'

            pdf_order.append(nome_cte_mesclada)

            gerar_etiqueta(path_etiqueta, input_signer, pedido)

            # Sobreposição (overlay) das cte com a etiqueta
            output = PdfWriter()

            with open(caminhoArquivoUnico, 'rb') as arquivoCTEs:
                pdfCTE = PdfReader(arquivoCTEs)

                with open(path_etiqueta, 'rb') as arquivoEtiqueta:
                    text_pdf = PdfReader(arquivoEtiqueta)

                    pag_cte = pdfCTE.pages[i]
                    pag_etiqueta = text_pdf.pages[0]

                    pag_cte.merge_page(pag_etiqueta)
                    output.add_page(pag_cte)

                    with open(path_cte_mesclada, "wb") as out_pdf:
                        output.write(out_pdf)
                        
    
    # merge das ctes etiquetadas
    unir_pdfs(f'./{data_formatada}', caminho_saida, pdf_order)

    print(f'\n\nEtiquetas e CTEs sobrepostas no arquivo "{caminho_saida}"')

    print(f'{ctes_semPedido} CT-e(s) ficaram sem pedido, favor preencher manualmente')

    testeEtiquetas.main(caminhoSaida, caminho_saida, caminhoMedicoes, input_data)
    
    return 0
    
    
