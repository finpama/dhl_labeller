import os
import shutil

try:
    os.mkdir("./Leitor de CTE")
except:
    print('Pasta ./Leitor de CTE já existe')

try:
    shutil.rmtree( './sample_data')
except:
    pass

!pip install PyPDF2
!pip install reportlab
!pip install pdfplumber