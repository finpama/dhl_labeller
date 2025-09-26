import os
import shutil

try:
    os.mkdir("./Leitor de CTE")
except:
    print('\nPasta "/Leitor de CTE" já existe...')

    inp = input('Deseja excluí-la? (s/n) > ').upper()
    print()

    if inp == "S":
        shutil.rmtree("./Leitor de CTE")
        os.mkdir("./Leitor de CTE")

try:
    os.remove('./CTES.pdf')
    os.remove('./Medições DHL.xlsm')
    os.remove('./relatório_cte.xlsx')
except:
    pass