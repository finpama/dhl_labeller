import os
from util import rmPasta

input('Antes de limpar, certifique-se de fechar as planilhas abertas e salvar arquivos que não devem ser excluídos... \n(Pressione qualquer tecla para continuar)')

try:
    os.mkdir("./Leitor de CTE")
except:
    print('\nPasta "/Leitor de CTE" já existe...')

    inp = input('Deseja excluí-la? (s/n) > ').upper()
    print()

    if inp == "S":
        rmPasta("./Leitor de CTE")
        os.mkdir("./Leitor de CTE")

try:
    os.remove('./CTES.pdf')
except:
    pass

try:
    os.remove('./Medições DHL.xlsm')
except:
    pass

try:
    os.remove('./relatório_cte.xlsx')
except:
    pass

try:
    os.remove('./relatorioErros_CT-E.xlsx')
except:
    pass
