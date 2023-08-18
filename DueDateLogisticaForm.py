###Programa Funcional OK Mayo 15/2023

import pandas as pd
import numpy as np
import math
import tabula
import xlwt
import openpyxl
import time
import xlsxwriter
import gspread
import pygsheets
from oauth2client.service_account import ServiceAccountCredentials
import json

#from datetime import  timedelta
#from Fun_DueDateLogistica import *
from Fun_EficienciaRutas import *
from datetime import datetime,timedelta
#import datetime
#from datetime import 
import io
import os
from os.path import splitext
from urllib.parse import urlparse
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from gooey import Gooey, GooeyParser

import gspread
import sys
import warnings
import gspread_dataframe as gd
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
#from pandas.core.common import SettingWithCopyWarning

#warnings.simplefilter(action="ignore", category=SettingWithCopyWarning)
print("Verificando contenido del archivo .....")


file_id0 = "1F0L_aHVNNhGuV-KNnuT6nCr_X1Af3l3E" #cortes2023.xlxs
file_id1 = "15vHlzGFgi9MjxyclqmNArvheijJhLSK5" #tiempos.xls
file_id2 = "17-gj5CshmFmb-hrFm3Dh4a12P8DcR4bB" #ReporteProduccionDBresultado.xlsx
scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials)# authenticate the JSON key with gspread
#service = build("drive", "v3", credentials=credentials)
#nombre = 'Duedate_Route_Deliveries.xls'
#service = build("drive", "v3", credentials=credentials)
service = build("drive", "v3", credentials=credentials,static_discovery=False)
  # Define the URL to download the file from
file_url = service.files().get(fileId=file_id2, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)

  # Define the filename to save the downloaded file as
filename = f"ReporteProduccionDB.xlsx"

  # Download the file
try:
    request = service.files().get_media(fileId=file_id2)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")

# Define the URL to download the file from
file_url = service.files().get(fileId=file_id0, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)

# Define the filename to save the downloaded file as
filename = f"cortes2023.xlsx"

# Download the file
try:
    request = service.files().get_media(fileId=file_id0)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")

# Define the URL to download the file from
file_url = service.files().get(fileId=file_id1, fields="webContentLink").execute()["webContentLink"]
parsed_url = urlparse(file_url)

# Define the filename to save the downloaded file as
filename = f"tiempos.xlsx"

# Download the file
try:
    request = service.files().get_media(fileId=file_id1)
    file = io.BytesIO()
    downloader = io.BytesIO()
    downloader.write(request.execute())
    downloader.seek(0)
    with open(filename, "wb") as f:
        f.write(downloader.getbuffer())
    print(f"File downloaded as {filename}")
except HttpError as error:
    print(f"An error occurred: {error}")    



@Gooey(program_name="Calculo de Due_Date Logistica")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.

    """
    #script_dir = os.path.dirname(os.path.abspath(__file__))

    # Cambia el directorio de trabajo al directorio del script
    #os.chdir(script_dir)
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Calculo de Due_Date de Trabajos pendientes Ver.3.0 Ago 17/2023')
    #parser.add_argument('Archivo_Produccion',
    #                    action='store',
    #                    default=stored_args.get('cust_file'),
    #                    widget='FileChooser',
    #                    help='Ej. ReporteProduccionDBresultado.xlsx')
    parser.add_argument('Archivo_Rutas',
                        action='store',
                        default=stored_args.get('cust_file'),
                        widget='FileChooser',
                        help='Ej. RouteDelivery.xlsx')
    parser.add_argument('Directorio_de_trabajo',
                        action='store',
                        default=stored_args.get('data_directory'),
                        widget='DirChooser',
                        help="Directorio con los archivos .XLSX/.CSV ")
    
    #parser.add_argument('Fecha', help='Seleccione Fecha del Reporte',
    #                    default=stored_args.get('Fecha'),
    #                    widget='DateChooser')
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args


def Principal(Directorio_de_trabajo,Rutas_pendientes):
  #print(Directorio_de_trabajo)
  # Obtiene la ruta del directorio donde se encuentra el script
  #credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json", scopes) #access the json key you downloaded earlier 
  #file = gspread.authorize(credentials)
  directory = os.getcwd()
  print("Directorio original: ",directory)
  print("Directorio de trabajo: ",Directorio_de_trabajo)
  """# Define the Drive API client
  service = build("drive", "v3", credentials=credentials)
  # Define the URL to download the file from
  file_url = service.files().get(fileId=file_id2, fields="webContentLink").execute()["webContentLink"]
  parsed_url = urlparse(file_url)

  # Define the filename to save the downloaded file as
  filename = f"ReporteProduccionDB.xlsx"

  # Download the file
  try:
      request = service.files().get_media(fileId=file_id2)
      file = io.BytesIO()
      downloader = io.BytesIO()
      downloader.write(request.execute())
      downloader.seek(0)
      with open(filename, "wb") as f:
          f.write(downloader.getbuffer())
      print(f"File downloaded as {filename}")
  except HttpError as error:
      print(f"An error occurred: {error}")

 """

  path = Rutas_pendientes
  nombre = os.path.basename(path)
  #nombre, extension = splitext(path)
  #print("nombre :",nombre)
 
  datos = pd.read_excel(path)
  ####os.remove(r'temporal235.csv')
  ####print(datos.head(5))
  ####os.rename(path, nomArchivo+'.pdf')
  indexDeleted = datos[datos['Job #'] == 'Job #'].index  
  datos.drop(indexDeleted,inplace=True)
  rutas = datos.copy()
  ####datos.to_excel(nomArchivo+'.xlsx', index=False)

  filename = f"ReporteProduccionDB.xlsx"
  #rutas= pd.read_excel(path)
  ####rutas= pd.read_excel(nomArchivo+'.xlsx')
  #concatenated_data=pd.read_excel(r'ReporteProduccionDB.xlsx')
  concatenated_data=pd.read_excel(filename)
  merged_data = pd.merge(rutas,concatenated_data[['Job #','Customer','Drop Location','R #','Stock #','Interchange','Part Description Summary',
      'Part Price','Created','Ship Via','Order Store #','Part Store #','Due']],on=['Job #'],how="left")
  # Write the concatenated data to a new .xlsx file
  ########merged_data.to_excel(nomArchivo+'temporal.xlsx', index=False)
  #merged_data.to_excel('temporalmerged.xlsx', index=False)
  ##merged_data["Delivery time"]= datetime.now() 
  #indexDeleted = ds2[ds2['Job Status'] ==  'Pickup'].index
  #ds2.drop(indexDeleted,inplace=True)
  #indexDeleted = merged_data[merged_data['Drop Location'] == ' '].index  # dejando solo las 253 en copia de audit trial
  #merged_data.drop(indexDeleted,inplace=True)
  merged_data_drop=merged_data.dropna(subset=['Drop Location']) # Nuevo Dataframe con solo registros con Drop Location
  dt = merged_data_drop.copy() #crea Copia de DataFrame para trabajar en el.
  merged_data = merged_data[merged_data['Drop Location'].isnull()]
  ##merged_data = merged_data.drop('Delivery time', axis=1)
  #merged_data.to_excel(r'Job no encontrados en Produccion.xlsx', index=False)
  #dl = pds.DataFrame() ##### Debug key
  #ds.to_excel(r'temporal.xlsx', sheet_name='BD-2022',header=True, index = False) 
  dt.to_excel(r'temporal.xlsx', index=False)
  ds= pd.read_excel(r'temporal.xlsx')
  os.remove(r'temporal.xlsx')
  #ds=merged_data.dropna(subset=['Drop Location'])



  hora = 0 
  #ds = pd.read_excel(nombre)
  ####columnas = sys.argv[2]

  ####columnas = columnas.split(',')
  #columnas=[12,16]
  columnas=[11,15]
  nombre1 = nombre.split('.')
  nombre2 = nombre.split('_')
  #nombre3 = nombre.split('\\')
  print(nombre2)
  #print(nombre3)
  

  

  nombre1 = nombre1[0] + "_Reporte.xlsx"
  #print("Ruta de salida  " , nombre1)
  filename2= Directorio_de_trabajo+'\\' +nombre1
  #print("asi se llama: ",filename2)
  # Define the Drive API client
  #service = build("drive", "v3", credentials=credentials)

  """# Define the URL to download the file from
  file_url = service.files().get(fileId=file_id0, fields="webContentLink").execute()["webContentLink"]
  parsed_url = urlparse(file_url)

  # Define the filename to save the downloaded file as
  filename = f"cortes2023.xlsx"

  # Download the file
  try:
      request = service.files().get_media(fileId=file_id0)
      file = io.BytesIO()
      downloader = io.BytesIO()
      downloader.write(request.execute())
      downloader.seek(0)
      with open(filename, "wb") as f:
          f.write(downloader.getbuffer())
      print(f"File downloaded as {filename}")
  except HttpError as error:
      print(f"An error occurred: {error}")

  # Define the URL to download the file from
  file_url = service.files().get(fileId=file_id1, fields="webContentLink").execute()["webContentLink"]
  parsed_url = urlparse(file_url)

  # Define the filename to save the downloaded file as
  filename = f"tiempos.xlsx"

  # Download the file
  try:
      request = service.files().get_media(fileId=file_id1)
      file = io.BytesIO()
      downloader = io.BytesIO()
      downloader.write(request.execute())
      downloader.seek(0)
      with open(filename, "wb") as f:
          f.write(downloader.getbuffer())
      print(f"File downloaded as {filename}")
  except HttpError as error:
      print(f"An error occurred: {error}")    
"""

  df = pd.read_excel(r'Tiempos.xlsx')
  dc = pd.read_excel(r'Cortes2023.xlsx')
  #dp = pd.read_excel(r'periodos.xlsx')



  df1 = pd.DataFrame({
      #"Unnamed: 0": "FALSE",
      "Job #": '',
      #"Order #": '',
      #"Type": '',
      #"Customer":'',
      "Interchange":'',
      #"Store #":'',
      "Stock #":'',
      #"Year":'',
      #"Model":'',
      #"Price":'',
      #"Created":'',
      "Due":'',
      #"Route":'',
      #"Salesperson":'',
      #"Driver":'',
      #"Event":'',
      #"Reason":'',
      #"Date":'',
      "Delivery Time":'',
      #Pickup Time":'',
      "Due Date Calculado":'',
      "Dias de atraso":'',
      "Conciliacion":'',
      "Diferencia DueDates":''
  }, index=["Dummy"])
  																						

  #date = pd.to_datetime(sys.argv[4])
  # Funcion Main para buscar todos los trabajos

  #dc= dc.set_index('DIA')
  df= df.set_index('Store')
  #dc = dc.to_dict()

  dc = dict(dc.set_index('DIA').groupby(level = 0).\
      apply(lambda x : x.to_dict(orient= 'list')))
  #print(dc)


  ds2 = timeFix(columnas,hora,ds)
  #dia = date.weekday()
  #fechaa =date + timedelta(hours = 16)
  #print(fechaa)


  ds2['Fecha Compromiso']=" "
  ds2['Due Date']=" "
  ds2['Dias de atraso']=" "
  ds2['Dia']= ds2['Created_y'].dt.dayofweek
  ds2['Dia'].mask(ds2['Dia'] == 6, 0, inplace=True)

  ds2['tiempo'] = pd.to_datetime(ds2['Created_y']).dt.time
  ds2['Fecha'] = pd.to_datetime(ds2['Created_y']).dt.date
  ds2['Conciliacion']=" "
  ds2['Delivery time']= datetime.now().date()###############
  #print(ds2['Delivery time'])####################
  #Asigno el valor de ruta
  dscompleto= ds2.copy()

  #indexDeleted = ds2[ds2['Drop Location'] == "nan"].index  # dejando solo las 253 en copia de audit trial
  #ds2.drop(indexDeleted,inplace=True)
  #ds2= ds2.replace('', np.nan).dropna()
  #ds2=ds2.dropna(subset=['Drop Location'])
  ####print("Creando archivo !!!!!!", nombre1)
  ####ds2.to_excel(r'prueba1.xlsx', sheet_name='BD-2022',header=True, index = False) 
  ####print(len(ds2))
  #ds2=ds3.copy()
  for i in range(len(ds2)) :

      Rt =ds2['Drop Location'][i]
      St =ds2['Part Store #'][i]
     
      print(Rt)
      print(St)
      if ( pd.isna(Rt)== False):
       ds2['Fecha Compromiso'][i]=df.at[St,Rt]

      
  ds2['Conciliacion'].mask(ds2['Fecha Compromiso'] == 99, 'Revisar', inplace=True)    

  def tabla(i,tiempo,c,b):
   pd.options.mode.chained_assignment = None 
   if ds2['tiempo'][i] < tiempo.time() :
          a=dc.get(ds2['Dia'][i])
          delt = a.get(c)        
          ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = delt[0])
          
   else :
          a=dc.get(ds2['Dia'][i])
          delt = a.get(b)      

          ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = delt[0])
          
  def tabla1(i,tiempo,c,b):
   if ds2['tiempo'][i] < tiempo.time() :
          ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = 0)     
   else :      
          ds2['Due Date'][i]= pd.to_datetime(ds2['Fecha'][i]) + timedelta(hours = 0)
          
  tiempo1 = datetime(2022,1,1,12,31,00) # asigno tiempos iniciales para comparar 12:31
  tiempo2 = datetime(2022,1,1,13,1,00) # asigno tiempos iniciales para comparar 13:01
  tiempo3 = datetime(2022,1,1,14,1,00) # asigno tiempos iniciales para comparar 14:01
  tiempo4 = datetime(2022,1,1,16,1,00) # asigno tiempos iniciales para comparar 16:01
  tiempo5 = datetime(2022,1,1,17,1,00) # asigno tiempos iniciales para comparar 17:01  Todas las tiendas cierre 
  tiempo6 = datetime(2022,1,1,15,1,00) # asigno tiempos iniciales para comparar 15:01 Economy Sabado

  for i in range(len(ds2)):
   if ds2['Fecha Compromiso'][i] != 99:
       if ds2['Dia'][i] in range(0,6) :   
        if ds2['Fecha Compromiso'][i] == 1 and ds2['Dia'][i]==5: 
          tabla(i,tiempo2,'1.2','1.3')
        elif ds2['Fecha Compromiso'][i] == 1:
          tabla(i,tiempo4,1,'1.1')     
        if ds2['Fecha Compromiso'][i] == 2 and ds2['Dia'][i]==5: 
          tabla(i,tiempo2,'2.2','2.3')
        elif ds2['Fecha Compromiso'][i] == 2:  
          tabla(i,tiempo4,2,'2.1')
        if ds2['Fecha Compromiso'][i] == 3 and ds2['Dia'][i]==5:
          tabla(i,tiempo2,'3.2','3.3')
        elif ds2['Fecha Compromiso'][i] == 3:
          tabla(i,tiempo4,3,'3.1')      
        if ds2['Fecha Compromiso'][i] == 4: 
          tabla(i,tiempo3,4,'4.1')     
        if ds2['Fecha Compromiso'][i] == 5:    
          tabla(i,tiempo1,5,'5.1')  
        if ds2['Fecha Compromiso'][i] == 6:    
          tabla(i,tiempo4,6,'6.1')         
        if ds2['Fecha Compromiso'][i] == 7:    
          tabla(i,tiempo3,7,'7.1')  
        if ds2['Fecha Compromiso'][i] == 8:    
          tabla(i,tiempo1,8,'8.1')
        if ds2['Fecha Compromiso'][i] == 9:    
          tabla(i,tiempo4,9,'9.1')
        if ds2['Fecha Compromiso'][i] == 10:  
          tabla(i,tiempo4,10,'10.1')
        if ds2['Fecha Compromiso'][i] == 11: 
          tabla(i,tiempo4,11,'11.1') 
        if ds2['Fecha Compromiso'][i] == 12:    
          tabla(i,tiempo1,12,'12.1')
        if ds2['Fecha Compromiso'][i] == 13:
          tabla(i,tiempo4,13,'13.1')
        if ds2['Fecha Compromiso'][i] == 14:    
          tabla(i,tiempo3,14,'14.1')
        if ds2['Fecha Compromiso'][i] == 15:    
          tabla(i,tiempo1,15,'15.1')
        if ds2['Fecha Compromiso'][i] == 16:    
          tabla(i,tiempo4,16,'16.1')  

        if ds2['Fecha Compromiso'][i] == 17 and ds2['Dia'][i]==5: 
          tabla(i,tiempo3,'17.2','17.3')     
        elif ds2['Fecha Compromiso'][i] == 17: 
          tabla(i,tiempo5,17,'17.1') 
        if ds2['Fecha Compromiso'][i] == 18 and ds2['Dia'][i]==5:
          tabla(i,tiempo6,'18.2','18.3')           
        elif ds2['Fecha Compromiso'][i] == 18:     
          tabla(i,tiempo5,18,'18.1')
        if ds2['Fecha Compromiso'][i] == 19 and ds2['Dia'][i]==5: 
          tabla(i,tiempo2,'19.2','19.3')   
        elif ds2['Fecha Compromiso'][i] == 19:   
          tabla(i,tiempo5,19,'19.1') 
        if ds2['Fecha Compromiso'][i] == 20 and ds2['Dia'][i]==5: 
          tabla(i,tiempo3,'20.2','20.3')    
        elif ds2['Fecha Compromiso'][i] == 20:   
          tabla(i,tiempo5,20,'20.1')        
              
   elif ds2['Fecha Compromiso'][i] == 99: 
       tabla1(i,tiempo4,1,'1.1')

  #for i in range(len(ds2)) :

  #    Due =ds2['Due Date'][i]
  #    print(Due)
  #    if ( pd.isna(Due)== False):
  #      ds2['Due Date'][i] = pd.to_datetime(ds2['Due Date'][i]).dt.date
  #writer = pd.ExcelWriter('prueba1.xlsx', engine='xlsxwriter')
  # Convert the dataframe to an XlsxWriter Excel object.

  #ds2[ds2['Drop Location']==''] = 'T'
  #print(ds2['Drop Location']=='T')   

  #bool_series = pd.notnull(ds2['Due Date'])  
  #print(bool_series)
  #def set_perf(row):
  #    if (row['Drop Location']==""):
  #        return row['Drop Location']    
  #    else:        
  #        return 1#row[pd.to_datetime(ds2['Due Date']).dt.date]

  #ds2['Due Date'] = ds2.apply(set_perf, axis=1)  
  ds2['Due Date'] = pd.to_datetime(ds2['Due Date']).dt.date
  ds2['temp1'] = pd.to_datetime(ds2['Delivery time']).dt.date
  ds2['Due_y'] = pd.to_datetime(ds2['Due_y']).dt.date

  #for j in range(len(ds2)):
  # if pd.isna(ds2.loc[j,'Pickup Time']) is False:
  #  ds2['temp1'][j] = ds2['temp'][j]

  #print(ds2['Due Date'])
  #print("ddddd")
  #print(ds2['Delivery Time'])
  ###print("Creando archivo  ?????????", nombre1)
  ####ds2.to_excel(r'prueba1.xlsx', sheet_name='BD-2022',header=True, index = False)  
  ds2['Diferencia DueDates']= ds2['Due_y'] - ds2['Due Date']
  ds2['Dias de atraso']= ds2['temp1'] - ds2['Due Date']
  del ds2["Dia"]
  del ds2['tiempo']
  del ds2['Fecha']
  #del ds2['Due_x']
  del ds2['temp1']
  del ds2['Fecha Compromiso']
  #del ds2['Delivery time']
  merged_data['Due_y']= merged_data['Due_x']
  #print(len(ds2['Unnamed: 0'])+1)
  #ds2['Unnamed: 0'][len(ds2['Unnamed: 0'])+1] = "Fin"
  ds2 = ds2.append(df1)
  #ds2.reindex(ds2.columns[ds2.columns != 'Conciliacion'].union(['Conciliacion']), axis=1)
  del ds2['Due Date Calculado']
  del ds2['Due']
  ds2 = ds2.append(merged_data)
  del ds2['Due_x']
  del ds2['Delivery Time']
  del ds2['Delivery time']
  ds2 = ds2.rename(columns={'Due_y': 'Due_Date_Vendedor', 'Due Date': 'Due_Date_Calculado'})
  writer = pd.ExcelWriter(filename2, engine='xlsxwriter')
  # Convert the dataframe to an XlsxWriter Excel object.
  print("Creando archivo", nombre1)
  ds2.to_excel(writer, sheet_name='Rutas Pendientes',header=True, index = False)

  while True:
      try:
          writer.close()
      except xlsxwriter.exceptions.FileCreateError as e:
          decision = input("Exception caught in workbook.close(): %s\n"
                           "Please close the file if it is open in Excel.\n"
                           "Try to write file again? [Y/n]: " % e)
          if decision != 'n':
              continue

      break

  insertRow = ["","","","","","","","","","","","","","","","","","","","","","","","","","","",]

if __name__ == '__main__':

  conf = parse_args()
  Principal(conf.Directorio_de_trabajo,conf.Archivo_Rutas)