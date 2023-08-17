###Programa Funcional OK Mayo 23/2023

import tabula
import pandas as pd
import xlwt
import openpyxl
import time
import xlsxwriter
import io
import os
import datetime
from os.path import splitext
import json
from gooey import Gooey, GooeyParser



path = r'RouteDelivery.pdf'
nombre, extension = splitext(path)

#timestamp = os.path.getmtime(path)
# convert timestamp into DateTime object
#datestamp = datetime.datetime.fromtimestamp(timestamp)
#print('Modified Date/Time:', datestamp.date())

# file creation
c_timestamp = os.path.getctime(path)
print(c_timestamp)
# convert creation timestamp into DateTime object
c_datestamp = datetime.datetime.fromtimestamp(c_timestamp)
 
#print('Created Date/Time on:', c_datestamp)
nomArchivo = nombre+c_datestamp.strftime("_%Y%m%d_%H%M%S")#+'.xlsx'


#print('nombreArchivo', nomArchivo)
#dl = pds.DataFrame() ##### Debug key
tabula.convert_into(path, "temporal235.csv", output_format="csv", pages='all')
datos = pd.read_csv("temporal235.csv")
os.remove(r'temporal235.csv')
print(datos.head(5))
os.rename(path, nomArchivo+'.pdf')
indexDeleted = datos[datos['Job #'] == 'Job #'].index  
datos.drop(indexDeleted,inplace=True)
datos.to_excel(nomArchivo+'.xlsx', index=False)

