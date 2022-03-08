import eel
from PyPDF2 import PdfFileMerger
import pathlib
import os
import pandas as pd

eel.init('web')

@eel.expose
def juntar():
    merger = PdfFileMerger()
    pathparent = str(pathlib.Path(str(pathlib.Path().resolve())).parent)
    control = pd.read_excel(pathparent+'/Controle.xls')
    if control.count == 0:
        for file in os.listdir(pathparent+"/merge"):
            path = os.path.join(pathparent+"/merge",file)
            f = open(path,"rb")
            merger.append(f)
        path1 = os.path.join(pathparent+"/merged","resultado.pdf")
        resultado = open(path1,"wb")
        merger.write(resultado)
    else:
        for i in range(1,control.count()['Arquivo']+1):
            row = control.loc[control['Ordem']==i]
            path = os.path.join(pathparent+"/merge",row['Arquivo'].item()+'.pdf')
            f = open(path,"rb")
            pg_i = int(row['Página inicial'].item())
            pg_f = int(row['Página final'].item())
            merger.append(f,pages=(pg_i-1,pg_f))
        path1 = os.path.join(pathparent+"/merged","resultado.pdf")
        resultado = open(path1,"wb")
        merger.write(resultado)
eel.start('index.html')