# Célula de Configuração

from datetime import date
from datetime import datetime
import os
from util import rmPasta

from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pdfplumber

from util import gerar_etiqueta, unir_pdfs


def main():
    pdf_order = []

    print('\nAguarde...')

    if not os.path.exists(f'./{data_formatada}'):
        os.mkdir(f'./{data_formatada}')
    else:
        rmPasta(f'./{data_formatada}')
        os.mkdir(f'./{data_formatada}')

    # Abre o PDF

    with pdfplumber.open(path_ctes) as pdf:

        # Percorre todas as páginas do PDF
        for i in range(len(pdf.pages)):

            pagina = pdf.pages[i]
            texto = pagina.extract_text()

            if not texto:
                raise Exception(f"ERRO: Nenhum texto encontrado no arquivo {path_ctes}")

            n_pagina = f'Pág.{i+1}'

            nome_etiqueta = f'{n_pagina} Etiqueta.pdf'
            nome_cte_mesclada = f'{n_pagina} Mesclada.pdf'

            path_etiqueta = f'./{data_formatada}/{nome_etiqueta}'
            path_cte_mesclada = f'./{data_formatada}/{nome_cte_mesclada}'

            pdf_order.append(nome_cte_mesclada)

            gerar_etiqueta(path_etiqueta, input_signer)

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
    unir_pdfs(f'./{data_formatada}', path_saida, pdf_order)

    # exclusão da pasta temporária
    try:
        rmPasta(f'./{data_formatada}')
    except:
        print(f'Pasta ./{data_formatada} já foi excluída')
        pass




    print(f'\n\nEtiquetas e CTEs sobrepostas no arquivo "{path_saida}"')


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
