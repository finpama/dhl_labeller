# Célula de Configuração

from datetime import date
from datetime import datetime
import os
import shutil
import re

import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pdfplumber

from util import gerar_base, gerar_etiqueta, merge_pdfs_in_directory


def main():
    # Para cada pedido, gera a etiqueta se tiver algo nos pedidos da planilha "Medições DHL.xlsx" no dia inserido no input. Depois sobrepor as CTEs com as etiquetas no arquivo "DHL <data do input>.pdf"'''

    df_planilha = gerar_base('./Medições DHL.xlsm', str(data))
    df_planilha[['CT-E', 'Pedido']] = df_planilha[['CT-E', 'Pedido']].astype(int)

    pdf_order = []
    ctes_semPedido = 0


    padrao = r"N[ºº]?[^\d]*([\d\.]+)"  # Regex que extrai número do CTE após "Nº"


    # se a df_planilha tiver pedidos, gera uma etiqueta pra cada pedido em uma pasta temporária
    if df_planilha.Pedido.empty:
        print(f'Nenhum pedido de {data_formatada} na planilha')
        return

    print('\nAguarde...')

    if not os.path.exists(f'./{data_formatada}'):
        os.mkdir(f'./{data_formatada}')
    else:
        shutil.rmtree(f'./{data_formatada}')
        os.mkdir(f'./{data_formatada}')

    df_planilha.reset_index(inplace=True)

    # Abre o PDF

    with pdfplumber.open(path_ctes) as pdf:

        if len(pdf.pages) != len(df_planilha.Pedido):
            print(f'O arquivo CTES.pdf tem {len(pdf.pages)} CT-Es, já a planilha no dia {data_formatada} tem {len(df_planilha.Pedido)} pedidos, favor verificar')
            return

        # Percorre todas as páginas do PDF
        for i in range(len(pdf.pages)):

            pagina = pdf.pages[i]
            texto = pagina.extract_text()

            if not texto:
                raise Exception(f"ERRO: Nenhum texto encontrado no arquivo {path_ctes}")

            # Aplica a regex
            match = re.search(padrao, texto, re.IGNORECASE)

            cte = int(match.group(1).strip()) if match else ""

            i_cte = df_planilha.index[df_planilha['CT-E'] == cte]

            try:
                pedido = int(df_planilha['Pedido'].iloc[i_cte].iloc[0])
            except:
                print(f'Pág. {i+1}: Ocorreu um erro ao procurar a CT-e, favor preencher manualmente...')
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

            with open(path_ctes, 'rb') as arquivoCTEs:
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
    merge_pdfs_in_directory(f'./{data_formatada}', path_saida, pdf_order)

    # exclusão da pasta temporária
    try:
        shutil.rmtree(f'./{data_formatada}')
    except:
        print(f'Pasta ./{data_formatada} já foi excluída')
        pass




    print(f'\n\nEtiquetas e CTEs sobrepostas no arquivo "{path_saida}"')

    print(f'{ctes_semPedido} CT-e(s) ficaram sem pedido, favor preencher manualmente')


# Geração
if __name__ == '__main__':

    # coleta de dados
    input_signer = input('Quem vai assinar? \nM: Michelle \nG: Gustavo \nQualquer outro valor gera sem assinante. \n\n')
    input_data = input('\n\n\nGerar etiquetas de qual data? \nUtilizar formato dd/mm/aa \n\n')

    if input_data != '':
        data = datetime.strptime(input_data, '%d/%m/%y').date()
    else:
        data = datetime.today().date()

    data_formatada = f'{data.day}-{data.month}-{data.year}'

    path_ctes = './CTES.pdf'
    path_saida = f'./DHL {data_formatada}.pdf'

    main()
