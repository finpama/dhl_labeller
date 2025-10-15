import pandas as pd

from datetime import date
import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pymupdf

def rmPasta(dir):
    'Remove uma pasta por completo (incluindo os arquivos)'
    
    arquivos = os.listdir(dir)
    
    for arquivo in arquivos:
        path = os.path.join(dir, arquivo)
        os.remove(path)
        
    os.rmdir(dir)

def findKey(dict, searchValue):
    'Encontra a primeira key do *value* em *dict*'
    
    for key, value in dict.items():
        if value == searchValue:
            return key
    
    raise ValueError(f'[findKey] A key do valor não foi encontrada')

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
def unir_pdfs(directory_path, output_filename, pdf_files_order):
    'Une os pdfs na pasta *directory_path* para o arquivo *output_filename*, utilizando a ordem em *pdf_files_order*'

    assert pdf_files_order, '[unir_pdfs] Sem arquivos em *pdf_files_order*'
    
    filenamesDict = {}

    for filename in pdf_files_order:
        if filename in filenamesDict:
            filenamesDict[filename] += 1
        else:
            filenamesDict[filename] = 1
            
    for filename, ocurrence in filenamesDict.items():
        assert ocurrence < 2, f'[unir_pdfs] Erro: o arquivo {filename}, aparece {ocurrence} vezes em *pdf_files_order*'
            
            
    doc = pymupdf.open()

    for pdf_filename in pdf_files_order:

        pdf_path = os.path.join(directory_path, pdf_filename)

        if not os.path.exists(pdf_path):
            print(f"Aviso: arquivo não encontrado: {pdf_filename}")

            continue
        
        try:
            doc.insert_file(pdf_path)  
            
        except Exception as e:
            print(f"Não foi possível ler o arquivo {pdf_filename}: {e}")
            

    doc.save(output_filename)
    doc.close()
    
    print(f"\nTodos os {len(pdf_files_order)} arquivos foram mesclados e estão no arquivo {output_filename}")
