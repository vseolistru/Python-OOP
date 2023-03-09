import datetime
import time
import xlrd
import xlwt 
from xlwt import Workbook
import csv
import itertools
from datetime import timedelta
import httplib2 
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import sqlite3
"""version 1.41"""

CREDENTIALS_FILE = 'creds.json'  
#______________________________________________________________
def sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1):
  name = ''.join(name)
  a= ','.join(data_time)
  b=','.join(goods)
  c=','.join(quantity)
  d=','.join(pay_cost)
  e=','.join(sum_1)
  conn = sqlite3.connect("scrap_base.db") 
  cursor = conn.cursor()
  cursor.execute("INSERT INTO "+str(name)+" ('date_time','goods','quantity','pay_cost','sum_1') VALUES ('"+str(a)+"','"+str(b)+"','"+str(c)+"','"+str(d)+"','"+str(e)+"')") 
  conn.commit()  
  conn.close()
#_______________________________________________________________
def download_sheet_to_csv(spreadsheet_id, sheet_name):
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range='A:N').execute()
    output_file = f'scrapping_values.csv'

    with open(output_file, 'w') as f:
        writer = csv.writer(f)
        writer.writerows(result.get('values'))
    f.close()
    print(f'загрузил scrapping_values.csv')
#________________________________________________________________________
def download_sheet_to_cs(spreadsheet_id, sheet_name):
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range='A:AD').execute()
    output_file = f'order_list.csv'

    with open(output_file, 'w') as f:
        writer = csv.writer(f)
        writer.writerows(result.get('values'))
    f.close()
    print(f'Загрузил order_list.csv')
#________________________________________________________________________
def google_table(batch_data,batch_id, batch_mass):
  print('Записываю накладные')
  lis=batch_data
  lis2=batch_mass
  lis3=batch_id
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "A3:A",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(1)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "C3:C",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis2], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(1)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "B3:B",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis3], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
#___________________scapping_post_to_ google_______________________

def google_table_post_scrapp_4_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2, quantity_3, pay_cost_3, sum_3, quantity_4, pay_cost_4, sum_4):
  print('Записываю списание')
  lis_1 = goods
  lis_2 = quantity_1
  lis_3 = pay_cost_1
  lis_4 = sum_1
  lis_5 = quantity_2
  lis_6 = pay_cost_2
  lis_7 = sum_2
  lis_8 = quantity_3
  lis_9 = pay_cost_3
  lis_10 = sum_3
  lis_11 = quantity_4
  lis_12 = pay_cost_4
  lis_13 = sum_4
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "A3:A",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_1], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "C3:C",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_2], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "D3:D",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_3], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "E3:E",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_4], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(4)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "F3:F",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_5], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "G3:G",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_6], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "H3:H",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_7], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(4)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "I3:I",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_8], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "J3:J",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_9], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "K3:K",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_10], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(4)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "L3:L",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_11], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "M3:M",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_12], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "N3:N",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_13], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
#_____________________________________________________________________________________
def google_table_post_scrapp_3_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2, quantity_3, pay_cost_3, sum_3):
  print('Записываю списание')
  lis_1 = goods
  lis_2 = quantity_1
  lis_3 = pay_cost_1
  lis_4 = sum_1
  lis_5 = quantity_2
  lis_6 = pay_cost_2
  lis_7 = sum_2
  lis_8 = quantity_3
  lis_9 = pay_cost_3
  lis_10 = sum_3
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "A3:A",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_1], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "C3:C",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_2], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "D3:D",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_3], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "E3:E",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_4], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(4)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "F3:F",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_5], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "G3:G",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_6], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "H3:H",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_7], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(4)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "I3:I",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_8], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "J3:J",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_9], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "K3:K",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_10], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
#_________________________________________________________________

def google_table_post_scrapp_2_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2):
  print('Записываю списание')
  lis_1 = goods
  lis_2 = quantity_1
  lis_3 = pay_cost_1
  lis_4 = sum_1
  lis_5 = quantity_2
  lis_6 = pay_cost_2
  lis_7 = sum_2
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "A3:A",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_1], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "C3:C",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_2], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "D3:D",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_3], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "E3:E",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_4], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  time.sleep(4)
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "F3:F",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_5], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "G3:G",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_6], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
  results = service.spreadsheets().values().batchUpdate(
    spreadsheetId = spreadsheet_id,
     body = {
        "valueInputOption": "USER_ENTERED", 
        "data": [
            {"range": "H3:H",
            "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
            "values": [
                    [i_value for i_value in lis_7], # Заполняем первую строку                   
                   ]}
          ]
      }
  ).execute()
#______________________________________________________________________________________
date = datetime.datetime.today().strftime("%d.%m.%y")
now = datetime.datetime.today()
one_day= timedelta(1)
next_date = now+one_day
new_date=next_date.strftime("%d.%m.%y")
#______________________________________________
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1lH6zhhEpEoUoz3sqcUjZAkt8P_RQOgk_2oHi5Ny_Sy0'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_cs(spreadsheet_id, 'Sheet1')
time.sleep(2)
#__________________________________________________________________________________
# - Кор Горького 6 
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data = []
batch_id =[]
batch_mass_t1 = []

for i in reader: 
  x = i['name']
  z = i['id']
  y = i['korolev-gor-6']
  if y !='':
    batch_data.append(str(x))
    batch_mass_t1.append(y) 
    batch_id.append(str(z)) 
f_1.close()  
#________________________________________________________________________________________________
#Пушкино Моск Проспект
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t2 = []
batch_id_t2 =[]
batch_mass_t2 = []
for i in reader: 
         x = i['name']
         z = i['id'] 
         y = i['pushkino-mosk-prosp']#pushkino-cheh-40-7.csv
         if y !='': 
            batch_data_t2.append(str(x))           
            batch_mass_t2.append(y)
            batch_id_t2.append(z)  
f_1.close()  
#_________________________________________________________________________________________________
#Пушкино Чехова
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t3 = []
batch_id_t3 =[]
batch_mass_t3 = []
for i in reader: 
         x = i['name']
         z = i['id'] 
         y = i['pushkino-cheh-40-7']#pushkino-oblaka
         if y !='': 
            batch_data_t3.append(str(x))           
            batch_mass_t3.append(y) 
            batch_id_t3.append(z) 
f_1.close() 
#________________________________________________________________________________________________
#Пушкино Облака
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t4 = []
batch_id_t4 =[]
batch_mass_t4 = []
for i in reader: 
         x = i['name']
         z = i['id'] 
         y = i['pushkino-oblaka']#sposad-skobyanka
         if y !='': 
            batch_data_t4.append(str(x))           
            batch_mass_t4.append((y))  
            batch_id_t4.append(z)
f_1.close()
#_______________________________________________________________________________________________
# Посад Скобянка
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t5 = []
batch_id_t5 =[]
batch_mass_t5 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['sposad-skobyanka']#sposad-armii-2
         if y !='': 
            batch_data_t5.append(str(x))           
            batch_mass_t5.append((y))  
            batch_id_t5.append(z)
f_1.close()
#_______________________________________________________________________________________________
#Красной армии - 203
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t6 = []
batch_id_t6 =[]
batch_mass_t6 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['sposad-armii-203']#
         if y !='': 
            batch_data_t6.append(str(x))           
            batch_mass_t6.append((y)) 
            batch_id_t6.append(z) 
f_1.close()
#_____________________________________________________________________________________________________
# Красной армии - 182
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t7 = []
batch_id_t7 =[]
batch_mass_t7 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['s-posad-armii-182']#
         if y !='': 
            batch_data_t7.append(str(x))           
            batch_mass_t7.append(y) 
            batch_id_t7.append(z) 
f_1.close()
#________________________________________________________________________________________________
# Красной армии - 185
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t8 = []
batch_id_t8 =[]
batch_mass_t8 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['s-posad-armii-185']#ivanteevka-sloboda
         if y !='': 
            batch_data_t8.append(str(x))           
            batch_mass_t8.append(y)
            batch_id_t8.append(z)   
f_1.close()
#________________________________________________________________________________________________
# Ивантеевка слобода
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t9 = []
batch_id_t9 =[]
batch_mass_t9 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['ivanteevka-sloboda']#chernogolovka
         if y !='': 
            batch_data_t9.append(str(x))           
            batch_mass_t9.append((y))  
            batch_id_t9.append(z)  
f_1.close()
#_______________________________________________________________________________________________
#Черноголовка
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t10 = []
batch_id_t10 =[]
batch_mass_t10 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['chernogolovka']#chernogolovka-2
         if y !='': 
            batch_data_t10.append(str(x))           
            batch_mass_t10.append(y)
            batch_id_t10.append(z) 
f_1.close()
#_______________________________________________________________________________________________
#Черноголовка-2
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t11 = []
batch_id_t11 =[]
batch_mass_t11 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['chernogolovka-2']#
         if y !='': 
            batch_data_t11.append(str(x))           
            batch_mass_t11.append((y)) 
            batch_id_t11.append(z)  
f_1.close()
#_______________________________________________________________________________________________
# Ивантеевка бережок
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t12 = []
batch_id_t12 =[]
batch_mass_t12 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['ivanteevka-beregok']#sofrino
         if y !='': 
            batch_data_t12.append(str(x))           
            batch_mass_t12.append((y))
            batch_id_t12.append(z)
f_1.close()
#_______________________________________________________________________________________________
#Софрино
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t13 = []
batch_id_t13 =[]
batch_mass_t13 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['sofrino']#
         if y !='': 
            batch_data_t13.append(str(x))           
            batch_mass_t13.append(y)
            batch_id_t13.append(z)
f_1.close()
#_______________________________________________________________________________________________
# Посад кр армии 3
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t14 = []
batch_id_t14 =[]
batch_mass_t14 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['sposad-armii-3']#
         if y !='': 
            batch_data_t14.append(str(x))           
            batch_mass_t14.append((y))
            batch_id_t14.append(z)
f_1.close()
#_______________________________________________________________________________________________
#Заветы ильича
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t15 = []
batch_id_t15 =[]
batch_mass_t15 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['zavety-ilicha']#
         if y !='': 
            batch_data_t15.append(str(x))           
            batch_mass_t15.append((y))
            batch_id_t15.append(z)
f_1.close()
#______________________________________________________________________________________________
# Арманд
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t16 = []
batch_id_t16 =[]
batch_mass_t16 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['pushkino-armand']#
         if y !='': 
            batch_data_t16.append(str(x))           
            batch_mass_t16.append((y))
            batch_id_t16.append(z)
f_1.close()
#______________________________________________________________________________________________
# Тургенева
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t17 = []
batch_id_t17 =[]
batch_mass_t17 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['pushkino-turgeneva']#
         if y !='': 
            batch_data_t17.append(str(x))           
            batch_mass_t17.append(y)
            batch_id_t17.append(z)
f_1.close()
#______________________________________________________________________________________________
# Рыбная
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t18 = []
batch_id_t18 =[]
batch_mass_t18 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['sposad-ribnay']#
         if y !='': 
            batch_data_t18.append(str(x))           
            batch_mass_t18.append(y)
            batch_id_t18.append(z)
f_1.close()
#____________________________________________________________________________________________________
# Интернет магазин
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t19 = []
batch_id_t19 =[]
batch_mass_t19 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['internet-magazin']#
         if y !='': 
            batch_data_t19.append(str(x))           
            batch_mass_t19.append(y)
            batch_id_t19.append(z)
f_1.close()
#__________________________________________________________
# Щелково
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t20 = []
batch_id_t20 =[]
batch_mass_t20 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['schelkovo']#
         if y !='': 
            batch_data_t20.append(str(x))           
            batch_mass_t20.append((y))
            batch_id_t20.append(z)
f_1.close()

#__________________________________________________________
# Спосад Кр-арм-5
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_t21 = []
batch_id_t21 =[]
batch_mass_t21 = []
for i in reader: 
         x = i['name']
         z = i['id']
         y = i['sposad-armii-5']#
         if y !='': 
            batch_data_t21.append(str(x))           
            batch_mass_t21.append((y))
            batch_id_t21.append(z)
f_1.close()
#_________маршруты_____________________________________________________________
#___________________________________________________
#vsego-1
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_name = []
batch_mass_item_b = []
batch_mass_item_c = []
batch_mass_item_d = []
batch_mass_item_e = []
batch_mass_item_f = []
batch_mass_item_i =[]
for i in reader: 
         a = i['name']
         b = i['korolev-gor-6']
         c = i['pushkino-mosk-prosp']
         d = i['pushkino-cheh-40-7']
         e = i['pushkino-oblaka']
         f = i['vsego-1']
         i = i['id']
         if f !='0': 
            batch_data_name.append(str(a))           
            batch_mass_item_b.append((b))
            batch_mass_item_c.append((c))
            batch_mass_item_d.append((d))
            batch_mass_item_e.append((e))
            batch_mass_item_f.append((f))
            batch_mass_item_i.append(str(i))
f_1.close()
#_____________________________________________
#vsego-2
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_name_1 = []
batch_mass_item_b_1 = []
batch_mass_item_c_1 = []
batch_mass_item_d_1 = []
batch_mass_item_e_1 = []
batch_mass_item_f_1 = []
batch_mass_item_i_1 = []
for i in reader: 
         a = i['name']
         b = i['sposad-skobyanka']
         c = i['sposad-armii-203']#
         d = i['s-posad-armii-182']
         e = i['s-posad-armii-185']
         f = i['vsego-2']
         i = i['id']
         if f !='0': 
            batch_data_name_1.append(str(a))           
            batch_mass_item_b_1.append(str(b))
            batch_mass_item_c_1.append(str(c))
            batch_mass_item_d_1.append(str(d))
            batch_mass_item_e_1.append(str(e))
            batch_mass_item_f_1.append(str(f))
            batch_mass_item_i_1.append(str(i))
f_1.close()
#__________________________________________________________
#vsego-3
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_name_2 = []
batch_mass_item_b_2 = []
batch_mass_item_c_2 = []
batch_mass_item_d_2 = []
batch_mass_item_f_2 = []
batch_mass_item_i_2 = []
for i in reader: 
         a = i['name']
         b = i['ivanteevka-sloboda']
         c = i['chernogolovka']#
         d = i['chernogolovka-2']
         f = i['vsego-3']
         i = i['id']
         if f !='0': 
            batch_data_name_2.append(str(a))           
            batch_mass_item_b_2.append((b))
            batch_mass_item_c_2.append((c))
            batch_mass_item_d_2.append((d))
            batch_mass_item_f_2.append((f))
            batch_mass_item_i_2.append(str(i))
f_1.close()
#________________________________________________
#vsego-4
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_name_3 = []
batch_mass_item_i_3 = []
batch_mass_item_b_3 = []
batch_mass_item_c_3 = []
batch_mass_item_d_3 = []
batch_mass_item_f_3 = []
batch_mass_item_z_3 = []
for i in reader: 
         a = i['name']
         m = i['id']
         b = i['sofrino']
         c = i['sposad-armii-3']
         d = i['sposad-armii-5']
         f = i['sposad-ribnay']
         z = i['vsego-4']

         if z !='0': 
            batch_data_name_3.append(str(a))  
            batch_mass_item_i_3.append(str(m))         
            batch_mass_item_b_3.append((b))
            batch_mass_item_c_3.append((c))
            batch_mass_item_d_3.append((d))
            batch_mass_item_f_3.append((f))
            batch_mass_item_z_3.append((z))

f_1.close()
#__________________________________________________________
#vsego-5
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_name_4 = []
batch_mass_item_b_4 = []
batch_mass_item_c_4 = []
batch_mass_item_d_4 = []
batch_mass_item_e_4 = []
batch_mass_item_f_4 = []
batch_mass_item_i_4 = []
for i in reader: 
         a = i['name']
         e = i['internet-magazin']
         b = i['zavety-ilicha']
         c = i['pushkino-armand']#
         d = i['pushkino-turgeneva']
         f = i['vsego-5']
         m = i['id']
         if f !='0': 
            batch_data_name_4.append(str(a)) 
            batch_mass_item_e_4.append((e))          
            batch_mass_item_b_4.append((b))
            batch_mass_item_c_4.append((c))
            batch_mass_item_d_4.append((d))
            batch_mass_item_f_4.append((f))
            batch_mass_item_i_4.append(str(m))
f_1.close()
#_________________________________________________________
#vsego-6
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_name_5 = []
batch_mass_item_b_5 = []
batch_mass_item_c_5 = []
batch_mass_item_f_5 = []
batch_mass_item_i_5 =[]
for i in reader: 
         a = i['name']
         b = i['schelkovo']
         c = i['ivanteevka-beregok']#
         f = i['vsego-6']
         m = i['id']
         if f !='0': 
            batch_data_name_5.append(str(a))           
            batch_mass_item_b_5.append((b))
            batch_mass_item_c_5.append((c))
            batch_mass_item_f_5.append((f))
            batch_mass_item_i_5.append(str(m))
f_1.close()
#___________________________________________________________
#sklad
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
batch_data_name_6 = []
batch_mass_item_f_6 = []
batch_mass_item_i_6 = []
for i in reader: 
         a = i['name']
         m = i['id']
         f = i['sklad']
         if f !='0': 
            batch_data_name_6.append(str(a))           
            batch_mass_item_f_6.append((f))
            batch_mass_item_i_6.append(str(m))
f_1.close()
#___________________________________________________________
#Main_result
f_1 = open ('order_list.csv', 'r', newline='')
reader = csv.DictReader(f_1, delimiter=',')
take_a = []
take_b =[]
take_c = []
take_d =[]
take_e = []
take_f =[]
take_g = []
take_h =[]
take_k = []
take_l = []
take_m =[]
take_n = []
take_o =[]
take_p = []
take_r =[]
take_j = []
take_s =[]
take_t = []
take_q =[]
take_w = []
take_y =[]
take_ab = []
take_ac =[]
take_ad = []
take_ae =[]
take_af =[]
take_ag = []
take_ah = []
take_all = []
for i in reader: 
         a = i['name']
         b = i['korolev-gor-6']#,,,,
         c = i['pushkino-mosk-prosp']
         d = i['pushkino-cheh-40-7']
         e = i['pushkino-oblaka']
         f = i['vsego-1']
         g = i['sposad-skobyanka']#,,,,
         h = i['sposad-armii-203']
         k = i['s-posad-armii-182']
         l = i['s-posad-armii-185']
         m = i['vsego-2']
         n = i['ivanteevka-sloboda']#,,,
         o = i['chernogolovka']
         p = i['chernogolovka-2']
         r = i['vsego-3']
         s = i['sofrino']
         t = i['sposad-armii-3']
         ah= i['sposad-armii-5']
         q = i['sposad-ribnay']
         w = i['vsego-4']
         j = i['internet-magazin']#,,
         y = i['zavety-ilicha']#,,,
         ab= i['pushkino-armand']
         ac= i['pushkino-turgeneva']
         ad= i['vsego-5']
         ae= i['schelkovo']#
         af= i['ivanteevka-beregok']
         ag= i['vsego-6']
         z = i['sklad']
         if z !='0': 
            take_a.append(str(a)) 
            take_b.append(str(b))
            take_c.append(str(c))
            take_d.append(str(d))
            take_e.append(str(e)) 
            take_f.append(str(f)) 
            take_g.append(str(g))
            take_h.append(str(h))
            take_k.append(str(k))
            take_l.append(str(l))              
            take_m.append(str(m)) 
            take_n.append(str(n))
            take_o.append(str(o))
            take_p.append(str(p))
            take_r.append(str(r))              
            take_j.append(str(j)) 
            take_s.append(str(s))
            take_t.append(str(t))
            take_q.append(str(q))
            take_w.append(str(w))              
            take_y.append(str(y)) 
            take_ab.append(str(ab))
            take_ac.append(str(ac))
            take_ad.append(str(ad))
            take_ae.append(str(ae))              
            take_af.append(str(af))
            take_ag.append(str(ag))
            take_ah.append(str(ah))
            take_all.append((z))
f_1.close()
#___________________________________________________________
wb=Workbook()
style_string='font:bold on; borders: bottom dashed; borders: top dotted; borders: right dashed;'
style_string_1='font:bold on; borders: bottom dashed; borders: top dotted;'
style1=xlwt.easyxf(style_string_1)
style=xlwt.easyxf(style_string)
#____________________________________________________________
sheet1 = wb.add_sheet('Основной заказ')
sheet1.write(0, 1, 'Основной заказ')
sheet1.write(0, 3, 'Дата')
sheet1.write(0, 4, date, style=style)
sheet1.write(4, 0, 'Наименование товара') #date
sheet1.write(4, 1, '', style=style1)
sheet1.write(4, 2, '', style=style)

sheet1.write(4, 3, 'Королев Горького 6', style=style) 
sheet1.write(4, 4, 'Пушкино Московский проспект', style=style)
sheet1.write(4, 5, 'Пушкино Чехова 40/7', style=style)
sheet1.write(4, 6, 'Пушкино Облака', style=style)
sheet1.write(4, 7, 'Всего в 1-ом', style=style)#
sheet1.write(4, 8, 'С-Посад Скобянка', style=style)
sheet1.write(4, 9, 'С-Посад Кр-армии-2', style=style)
sheet1.write(4, 10, 'С-Посад Кр-армии-182', style=style)
sheet1.write(4, 11, 'С-Посад Кр-армии-185', style=style)
sheet1.write(4, 12, 'Всего во 2-ом', style=style)#
sheet1.write(4, 13, 'Ивантеевка Н-Слобода', style=style)
sheet1.write(4, 14, 'Черноголовка', style=style)
sheet1.write(4, 15, 'Черноголовка-2', style=style)
sheet1.write(4, 16, 'Всего в 3-ем', style=style)#
sheet1.write(4, 17, 'Софрино', style=style)
sheet1.write(4, 18, 'С-Посад Кр-армии-3', style=style)
sheet1.write(4, 19, 'С-Посад Кр-армии-5', style=style)
sheet1.write(4, 20, 'С-Посад Рыбная', style=style)
sheet1.write(4, 21, 'Всего в 4-ом', style=style)#
sheet1.write(4, 22, 'Интернет магазин', style=style)
sheet1.write(4, 23, 'Заветы Ильича', style=style)
sheet1.write(4, 24, 'Пушкино Арманд', style=style)
sheet1.write(4, 25, 'Пушкино Тургенева 18', style=style)
sheet1.write(4, 26, 'Всего в 5-ом', style=style)#
sheet1.write(4, 27, 'Щелково', style=style)
sheet1.write(4, 28, 'Ивантеевка бережок', style=style)
sheet1.write(4, 29, 'Всего в 6-ом', style=style)# 
sheet1.write(4, 30, 'Общий заказ', style=style) 
row=5
col=1
for x in take_a:   
  sheet1.write(row, 0, x, style=style1)# name
  sheet1.write(row, 1, '', style=style1)
  sheet1.write(row, 2, '', style=style)
  row+=1
row=5
col=1
for x in take_b:   
  sheet1.write(row, 3, x, style=style)
  row+=1
row=5
col=1
for x in take_c:   
  sheet1.write(row, 4, x, style=style)
  row+=1
row=5
col=1
for x in take_d:   
  sheet1.write(row, 5, x, style=style)
  row+=1
row=5
col=1
for x in take_e:   
  sheet1.write(row, 6, x, style=style)
  row+=1
row=5
col=1
for x in take_f:   
  sheet1.write(row, 7, x, style=style)# vsego
  row+=1
row=5
col=1
for x in take_g:   
  sheet1.write(row, 8, x, style=style)
  row+=1
row=5
col=1
for x in take_h:   
  sheet1.write(row, 9, x, style=style)
  row+=1
row=5
col=1
for x in take_k:   
  sheet1.write(row, 10, x, style=style)
  row+=1
row=5
col=1
for x in take_l:   
  sheet1.write(row, 11, x, style=style)
  row+=1
row=5
col=1
for x in take_m:   
  sheet1.write(row, 12, x, style=style)# vsego-2
  row+=1
row=5
col=1
for x in take_n:   
  sheet1.write(row, 13, x, style=style)
  row+=1
row=5
col=1
for x in take_o:   
  sheet1.write(row, 14, x, style=style)
  row+=1
row=5
col=1
for x in take_p:   
  sheet1.write(row, 15, x, style=style)
  row+=1
row=5
col=1
for x in take_r:   
  sheet1.write(row, 16, x, style=style)#vsego -3
  row+=1
row=5
col=1
for x in take_s:   
  sheet1.write(row, 17, x, style=style)
  row+=1
row=5
col=1
for x in take_t:   
  sheet1.write(row, 18, x, style=style)
  row+=1
row=5
col=1
for x in take_ah:   
  sheet1.write(row, 19, x, style=style)
  row+=1
row=5
col=1
for x in take_q:   
  sheet1.write(row, 20, x, style=style)
  row+=1
row=5
col=1
for x in take_w:   
  sheet1.write(row, 21, x, style=style)#vsego -4
  row+=1
row=5
col=1
for x in take_j:   
  sheet1.write(row, 22, x, style=style)
  row+=1
row=5
col=1
for x in take_y:   
  sheet1.write(row, 23, x, style=style)
  row+=1
row=5
col=1
for x in take_ab:   
  sheet1.write(row, 24, x, style=style)
  row+=1
row=5
col=1
for x in take_ac:   
  sheet1.write(row, 25, x, style=style)
  row+=1
row=5
col=1
for x in take_ad:   
  sheet1.write(row, 26, x, style=style)#vsego -5
  row+=1
row=5
col=1
for x in take_ae:   
  sheet1.write(row, 27, x, style=style)
  row+=1
row=5
col=1
for x in take_af:   
  sheet1.write(row, 28, x, style=style)
  row+=1
row=5
col=1
for x in take_ag:   
  sheet1.write(row, 29, x, style=style)#vsego -6
  row+=1
row=5
col=1
for x in take_all:   
  sheet1.write(row, 30, x, style=style)#vsego купить
  row+=1
#____________________________________________________________
sheet2 = wb.add_sheet('Маршрут -1' , cell_overwrite_ok=True)
sheet2.write(0, 2, 'Маршрут - 1')
sheet2.write(0, 4, 'Дата:', style=style1)
sheet2.write(0, 5, new_date, style=style)
sheet2.write(2, 0, 'Название точек:')
sheet2.write(2, 1, 'Королев Горького 6, Пушкино Московский проспект, Пушкино Чехова 40/7, Пушкино Облака', style=style1)
sheet2.write(4, 0, 'позиции для отгрузки', style=style1)
sheet2.write(4, 1, '', style=style1)
sheet2.write(4, 2, 'ед.изм', style=style)
sheet2.write(4, 3, 'Кор-Горк 6', style=style)
sheet2.write(4, 4, 'Пуш-Моск пр', style=style)
sheet2.write(4, 5, 'Пуш-Чех 40/7', style=style)
sheet2.write(4, 6, 'Пуш-Облак', style=style)
sheet2.write(4, 7, 'Всего:', style=style)
if len(batch_data_name)>5:
   a=batch_data_name[0:55]
   b=batch_data_name[55:114]
   c=batch_data_name[114:173]
   d=batch_data_name[173:232]
   e=batch_data_name[232:291]
   row=5
   col=0
   for x in a:   
     sheet2.write(row, 0, x, style=style1)
     sheet2.write(row, 1, '', style=style1)     
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet2.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(60, 1, '', style=style1)    
     for x in b:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)      
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet2.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(120, 1, '', style=style1)  
     for x in c:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)      
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet2.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(180, 1, '', style=style1)
     sheet2.write(180, 2, '', style=style)
     for x in d:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)      
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet2.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(240, 1, '', style=style1)    
     for x in e:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)      
       row+=1
#_____________________________
if len(batch_mass_item_i)>5:
   a=batch_mass_item_i[0:55]
   b=batch_mass_item_i[55:114]
   c=batch_mass_item_i[114:173]
   d=batch_mass_item_i[173:232]
   e=batch_mass_item_i[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 2, 'ед.изм', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 2, 'ед.изм', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 2, 'ед.изм', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 2, 'ед.изм', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
#____________________________________
if len(batch_mass_item_b)>5:
   a=batch_mass_item_b[0:55]
   b=batch_mass_item_b[55:114]
   c=batch_mass_item_b[114:173]
   d=batch_mass_item_b[173:232]
   e=batch_mass_item_b[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 3, 'Кор-Горк 6', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 3, 'Кор-Горк 6', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 3, 'Кор-Горк 6', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 3, 'Кор-Горк 6', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 3, y, style=style)
       row+=1 


if len(batch_mass_item_c)>5:
   a=batch_mass_item_c[0:55]
   b=batch_mass_item_c[55:114]
   c=batch_mass_item_c[114:173]
   d=batch_mass_item_c[173:232]
   e=batch_mass_item_c[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 4, 'Пуш-Моск пр', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 4, 'Пуш-Моск пр', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 4, 'Пуш-Моск пр', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 4, 'Пуш-Моск пр', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 4, y, style=style)
       row+=1 


if len(batch_mass_item_d)>5:
   a=batch_mass_item_d[0:55]
   b=batch_mass_item_d[55:114]
   c=batch_mass_item_d[114:173]
   d=batch_mass_item_d[173:232]
   e=batch_mass_item_d[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 5, 'Пуш-Чех 40/7', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 5, 'Пуш-Чех 40/7', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 5, 'Пуш-Чех 40/7', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 5, 'Пуш-Чех 40/7', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
if len(batch_mass_item_e)>5:
   a=batch_mass_item_e[0:55]
   b=batch_mass_item_e[55:114]
   c=batch_mass_item_e[114:173]
   d=batch_mass_item_e[173:232]
   e=batch_mass_item_e[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 6, 'Пуш-Облак', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 6, 'Пуш-Облак', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 6, 'Пуш-Облак', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 6, 'Пуш-Облак', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
if len(batch_mass_item_f)>5:
   a=batch_mass_item_f[0:55]
   b=batch_mass_item_f[55:114]
   c=batch_mass_item_f[114:173]
   d=batch_mass_item_f[173:232]
   e=batch_mass_item_f[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 7, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 7, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 7, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 7, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   row+=2
   sheet2.write(row, 0, 'Получил:')
   sheet2.write(row, 2, '________________/_______________/')
   row+=1
#__________________________________________
sheet3 = wb.add_sheet('Маршрут -2') 
sheet3.write(0, 2, 'Маршрут - 2')
sheet3.write(0, 4, 'Дата:', style=style1)
sheet3.write(0, 5, new_date, style=style)
sheet3.write(2, 0, 'Название точек:')
sheet3.write(2, 1, 'С-Посад Скобянка, С-Посад Кр-армии-182, С-Посад Кр-армии-203, С-Посад Кр-армии-185', style=style1)
sheet3.write(4, 0, 'позиции для отгрузки', style=style1)
sheet3.write(4, 1, '', style=style1)
sheet3.write(4, 2, 'ед.изм', style=style)
sheet3.write(4, 3, 'С-Пос Скоб', style=style)
sheet3.write(4, 5, 'С-Пос-203', style=style)
sheet3.write(4, 4, 'С-Пос-182', style=style)
sheet3.write(4, 6, 'С-Пос-185', style=style)
sheet3.write(4, 7, 'Всего:', style=style)
if len(batch_data_name_1)>5:
   a=batch_data_name_1[0:55]
   b=batch_data_name_1[55:114]
   c=batch_data_name_1[114:173]
   d=batch_data_name_1[173:232]
   e=batch_data_name_1[232:291]
   row=5
   col=0
   for x in a:   
     sheet3.write(row, 0, x, style=style1)
     sheet3.write(row, 1, '', style=style1)     
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet3.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(60, 1, '', style=style1)   
     for x in b:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style1)      
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet3.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(120, 1, '', style=style1)     
     for x in c:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style1)      
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet3.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(180, 1, '', style=style1)     
     for x in d:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style1)       
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet3.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(240, 1, '', style=style1)     
     for x in e:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style1)    
       row+=1

#_____________________________
if len(batch_mass_item_i_1)>5:
   a=batch_mass_item_i_1[0:55]
   b=batch_mass_item_i_1[55:114]
   c=batch_mass_item_i_1[114:173]
   d=batch_mass_item_i_1[173:232]
   e=batch_mass_item_i_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 2, 'ед.изм', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 2, 'ед.изм', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 2, 'ед.изм', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 2, 'ед.изм', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
#____________________________________

if len(batch_mass_item_b_1)>5:
   a=batch_mass_item_b_1[0:55]
   b=batch_mass_item_b_1[55:114]
   c=batch_mass_item_b_1[114:173]
   d=batch_mass_item_b_1[173:232]
   e=batch_mass_item_b_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 3, 'С-Пос Скоб', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 3, 'С-Пос Скоб', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 3, 'С-Пос Скоб', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 3, 'С-Пос Скоб', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 3, y, style=style)
       row+=1 
if len(batch_mass_item_c_1)>5:
   a=batch_mass_item_c_1[0:55]
   b=batch_mass_item_c_1[55:114]
   c=batch_mass_item_c_1[114:173]
   d=batch_mass_item_c_1[173:232]
   e=batch_mass_item_c_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 5, 'С-Пос-203', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 5, 'С-Пос-203', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 5, 'С-Пос-203', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 5, 'С-Пос-203', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 5, y, style=style)
       row+=1 
if len(batch_mass_item_d_1)>5:
   a=batch_mass_item_d_1[0:55]
   b=batch_mass_item_d_1[55:114]
   c=batch_mass_item_d_1[114:173]
   d=batch_mass_item_d_1[173:232]
   e=batch_mass_item_d_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 4, 'С-Пос-182', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 4, 'С-Пос-182', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 4, 'С-Пос-182', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 4, 'С-Пос-182', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 4, y, style=style)
       row+=1
if len(batch_mass_item_e_1)>5:
   a=batch_mass_item_e_1[0:55]
   b=batch_mass_item_e_1[55:114]
   c=batch_mass_item_e_1[114:173]
   d=batch_mass_item_e_1[173:232]
   e=batch_mass_item_e_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 6, 'С-Пос-185', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 6, 'С-Пос-185', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 6, 'С-Пос-185', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 6, 'С-Пос-185', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
if len(batch_mass_item_f_1)>5:
   a=batch_mass_item_f_1[0:55]
   b=batch_mass_item_f_1[55:114]
   c=batch_mass_item_f_1[114:173]
   d=batch_mass_item_f_1[173:232]
   e=batch_mass_item_f_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 7, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 7, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 7, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 7, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 7, y, style=style)
       row+=1 
   row+=2
   sheet3.write(row, 0, 'Получил:')
   sheet3.write(row, 2, '________________/_______________/')
   row+=1
#____________________________________
sheet4 = wb.add_sheet('Маршрут -3') 
sheet4.write(0, 2, 'Маршрут - 3')
sheet4.write(0, 4, 'Дата:', style=style1)
sheet4.write(0, 5, new_date, style=style)
sheet4.write(2, 0, 'Название точек:')
sheet4.write(2, 1, 'Ивантеевка Н-Слобода, Черноголовка, Черноголовка-2', style=style1)
sheet4.write(4, 1, '', style=style1)
sheet4.write(4, 2, 'ед.изм', style=style)
sheet4.write(4, 0, 'позиции для отгрузки', style=style1)
sheet4.write(4, 3, 'Ив Н-Слобода', style=style)
sheet4.write(4, 4, 'Черноголов', style=style)
sheet4.write(4, 5, 'Черногол-2', style=style)
sheet4.write(4, 6, 'Всего:', style=style)
if len(batch_data_name_2)>5:
   a=batch_data_name_2[0:55]
   b=batch_data_name_2[55:114]
   c=batch_data_name_2[114:173]
   d=batch_data_name_2[173:232]
   e=batch_data_name_2[232:291]
   row=5
   col=0
   for x in a:   
     sheet4.write(row, 0, x, style=style1)
     sheet4.write(row, 1, '', style=style1)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet4.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(60, 1, '', style=style1)
     for x in b:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)      
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet4.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(120, 1, '', style=style1)    
     for x in c:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)       
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet4.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(180, 1, '', style=style1)    
     for x in d:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)     
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet4.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(240, 1, '', style=style1)  
     for x in e:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)
       row+=1
#_____________________________
if len(batch_mass_item_i_2)>5:
   a=batch_mass_item_i_2[0:55]
   b=batch_mass_item_i_2[55:114]
   c=batch_mass_item_i_2[114:173]
   d=batch_mass_item_i_2[173:232]
   e=batch_mass_item_i_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 2, 'ед.изм', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 2, 'ед.изм', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 2, 'ед.изм', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 2, 'ед.изм', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 2, y, style=style)
       row+=1 
#____________________________________

if len(batch_mass_item_b_2)>5:
   a=batch_mass_item_b_2[0:55]
   b=batch_mass_item_b_2[55:114]
   c=batch_mass_item_b_2[114:173]
   d=batch_mass_item_b_2[173:232]
   e=batch_mass_item_b_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 3, 'Ив Н-Слобода', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 3, 'Ив Н-Слобода', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 3, 'Ив Н-Слобода', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 3, 'Ив Н-Слобода', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 3, y, style=style)
       row+=1 
if len(batch_mass_item_c_2)>5:
   a=batch_mass_item_c_2[0:55]
   b=batch_mass_item_c_2[55:114]
   c=batch_mass_item_c_2[114:173]
   d=batch_mass_item_c_2[173:232]
   e=batch_mass_item_c_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 4, 'Черноголов', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 4, 'Черноголов', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 4, 'Черноголов', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 4, 'Черноголов', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 4, y, style=style)
       row+=1
if len(batch_mass_item_d_2)>5:
   a=batch_mass_item_d_2[0:55]
   b=batch_mass_item_d_2[55:114]
   c=batch_mass_item_d_2[114:173]
   d=batch_mass_item_d_2[173:232]
   e=batch_mass_item_d_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 5, 'Черногол-2', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 5, 'Черногол-2', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 5, 'Черногол-2', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 5, 'Черногол-2', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 5, y, style=style)
       row+=1
if len(batch_mass_item_f_2)>5:
   a=batch_mass_item_f_2[0:55]
   b=batch_mass_item_f_2[55:114]
   c=batch_mass_item_f_2[114:173]
   d=batch_mass_item_f_2[173:232]
   e=batch_mass_item_f_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 6, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 6, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 6, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 6, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 6, y, style=style)
       row+=1
   row+=2
   sheet4.write(row, 0, 'Получил:')
   sheet4.write(row, 2, '________________/_______________/')
   row+=1
#____________________________________
sheet5 = wb.add_sheet('Маршрут -4') 
sheet5.write(0, 2, 'Маршрут - 4')
sheet5.write(0, 4, 'Дата:', style=style1)
sheet5.write(0, 5, new_date, style=style)
sheet5.write(2, 0, 'Название точек:')
sheet5.write(2, 1, 'Софрино, С-Посад Рыбная, С-Посад Кр-армии-5,С-Посад Кр-армии-3,', style=style1)
sheet5.write(4, 0, 'позиции для отгрузки', style=style1)
sheet5.write(4, 1, '', style=style1)
sheet5.write(4, 2, 'ед.изм', style=style)
sheet5.write(4, 3, 'Софрино', style=style)
sheet5.write(4, 4, 'СП-Рыбная', style=style)
sheet5.write(4, 5, 'СП Кр.ар-5', style=style)
sheet5.write(4, 6, 'СП Кр.ар-3', style=style)
sheet5.write(4, 7, 'Всего:', style=style)
if len(batch_data_name_3)>5:
   a=batch_data_name_3[0:55]
   b=batch_data_name_3[55:114]
   c=batch_data_name_3[114:173]
   d=batch_data_name_3[173:232]
   e=batch_data_name_3[232:291]
   row=5
   col=0
   for x in a:   
     sheet5.write(row, 0, x, style=style1)
     sheet5.write(row, 1, '', style=style1)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet5.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(60, 1, '', style=style1)
     for x in b:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet5.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(120, 1, '', style=style1)
     for x in c:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet5.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(180, 1, '', style=style1)
     for x in d:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet5.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(240, 1, '', style=style1)
     for x in e:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       row+=1
if len(batch_mass_item_b_3)>5:
   a=batch_mass_item_b_3[0:55]
   b=batch_mass_item_b_3[55:114]
   c=batch_mass_item_b_3[114:173]
   d=batch_mass_item_b_3[173:232]
   e=batch_mass_item_b_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 3, 'Софрино', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 3, 'Софрино', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 3, 'Софрино', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 3, 'Софрино', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 3, y, style=style)
       row+=1 

#_____________________________
if len(batch_mass_item_i_3)>5:
   a=batch_mass_item_i_3[0:55]
   b=batch_mass_item_i_3[55:114]
   c=batch_mass_item_i_3[114:173]
   d=batch_mass_item_i_3[173:232]
   e=batch_mass_item_i_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 2, 'ед.изм', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 2, 'ед.изм', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 2, 'ед.изм', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 2, 'ед.изм', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 2, y, style=style)
       row+=1 
#____________________________________

if len(batch_mass_item_f_3)>5:
   a=batch_mass_item_f_3[0:55]
   b=batch_mass_item_f_3[55:114]
   c=batch_mass_item_f_3[114:173]
   d=batch_mass_item_f_3[173:232]
   e=batch_mass_item_f_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 4, 'СП-Рыбная', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 4, 'СП-Рыбная', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 4, 'СП-Рыбная', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 4, 'СП-Рыбная', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
if len(batch_mass_item_d_3)>5:
   a=batch_mass_item_d_3[0:55]
   b=batch_mass_item_d_3[55:114]
   c=batch_mass_item_d_3[114:173]
   d=batch_mass_item_d_3[173:232]
   e=batch_mass_item_d_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 5, 'СП Кр.ар-5', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 5, 'СП Кр.ар-5', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 5, 'СП Кр.ар-5', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 5, 'СП Кр.ар-5', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 5, y, style=style)
       row+=1 

if len(batch_mass_item_c_3)>5:
   a=batch_mass_item_c_3[0:55]
   b=batch_mass_item_c_3[55:114]
   c=batch_mass_item_c_3[114:173]
   d=batch_mass_item_c_3[173:232]
   e=batch_mass_item_c_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 6, 'СП Кр.ар-3', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 6, 'СП Кр.ар-3', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 6, 'СП Кр.ар-3', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 6, 'СП Кр.ар-3', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
if len(batch_mass_item_z_3)>5:
   a=batch_mass_item_z_3[0:55]
   b=batch_mass_item_z_3[55:114]
   c=batch_mass_item_z_3[114:173]
   d=batch_mass_item_z_3[173:232]
   e=batch_mass_item_z_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 7, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 7, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 7, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 7, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 7, y, style=style)
       row+=1
   row+=2
   sheet5.write(row, 0, 'Получил:')
   sheet5.write(row, 2, '________________/_______________/')
   row+=1
#__________________________________________
sheet6 = wb.add_sheet('Маршрут -5', cell_overwrite_ok=True) 
sheet6.write(0, 2, 'Маршрут - 5')
sheet6.write(0, 4, 'Дата:', style=style1)
sheet6.write(0, 5, new_date, style=style)
sheet6.write(2, 0, 'Название точек:')
sheet6.write(2, 1, 'Интернет-магазин, Заветы Ильича, Пушкино Арманд, Пушкино Тургенева -18', style=style1)
sheet6.write(4, 0, 'позиции для отгрузки', style=style1)
sheet6.write(4, 1, '', style=style1)
sheet6.write(4, 2, 'ед.изм', style=style)
sheet6.write(4, 3, 'Инт-магазин', style=style)
sheet6.write(4, 4, 'Заветы Ил', style=style)
sheet6.write(4, 5, 'Пуш Арманд', style=style)
sheet6.write(4, 6, 'Пуш Тур-18', style=style)
sheet6.write(4, 7, 'Всего:', style=style)
if len(batch_data_name_4)>5:
   a=batch_data_name_4[0:55]
   b=batch_data_name_4[55:114]
   c=batch_data_name_4[114:173]
   d=batch_data_name_4[173:232]
   e=batch_data_name_4[232:291]
   row=5
   col=0
   for x in a:   
     sheet6.write(row, 0, x, style=style1)
     sheet6.write(row, 1, '', style=style1)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet6.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(60, 1, '', style=style1)
     for x in b:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet6.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(120, 1, '', style=style1)
     for x in c:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet6.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(180, 1, '', style=style1)
     for x in d:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet6.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(240, 1, '', style=style1)
     for x in e:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       row+=1
if len(batch_mass_item_e_4)>5:
   a=batch_mass_item_e_4[0:55]
   b=batch_mass_item_e_4[55:114]
   c=batch_mass_item_e_4[114:173]
   d=batch_mass_item_e_4[173:232]
   e=batch_mass_item_e_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 3, 'Инт-магазин', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 3, 'Инт-магазин', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 3, 'Инт-магазин', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 3, 'Инт-магазин', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 3, y, style=style)
       row+=1 

#_____________________________
if len(batch_mass_item_i_4)>5:
   a=batch_mass_item_i_4[0:55]
   b=batch_mass_item_i_4[55:114]
   c=batch_mass_item_i_4[114:173]
   d=batch_mass_item_i_4[173:232]
   e=batch_mass_item_i_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 2, 'ед.изм', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 2, 'ед.изм', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 2, 'ед.изм', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 2, 'ед.изм', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 2, y, style=style)
       row+=1 
#____________________________________

if len(batch_mass_item_b_4)>5:
   a=batch_mass_item_b_4[0:55]
   b=batch_mass_item_b_4[55:114]
   c=batch_mass_item_b_4[114:173]
   d=batch_mass_item_b_4[173:232]
   e=batch_mass_item_b_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 4, 'Заветы Ил', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 4, 'Заветы Ил', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 4, 'Заветы Ил', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 4, 'Заветы Ил', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 4, y, style=style)
       row+=1 
if len(batch_mass_item_c_4)>5:
   a=batch_mass_item_c_4[0:55]
   b=batch_mass_item_c_4[55:114]
   c=batch_mass_item_c_4[114:173]
   d=batch_mass_item_c_4[173:232]
   e=batch_mass_item_c_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 5, 'Пуш Арманд', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 5, 'Пуш Арманд', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 5, 'Пуш Арманд', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 5, 'Пуш Арманд', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 5, y, style=style)
       row+=1 
if len(batch_mass_item_d_4)>5:
   a=batch_mass_item_d_4[0:55]
   b=batch_mass_item_d_4[55:114]
   c=batch_mass_item_d_4[114:173]
   d=batch_mass_item_d_4[173:232]
   e=batch_mass_item_d_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 6, 'Пуш Тур-18', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 6, 'Пуш Тур-18', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 6, 'Пуш Тур-18', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 6, 'Пуш Тур-18', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 6, y, style=style)
       row+=1 
if len(batch_mass_item_f_4)>5:
   a=batch_mass_item_f_4[0:55]
   b=batch_mass_item_f_4[55:114]
   c=batch_mass_item_f_4[114:173]
   d=batch_mass_item_f_4[173:232]
   e=batch_mass_item_f_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 7, 'Всего', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 7, 'Всего', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 7, 'Всего', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 7, 'Всего', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 7, y, style=style)
       row+=1 
   row+=2
   sheet6.write(row, 0, 'Получил:')
   sheet6.write(row, 2, '________________/_______________/')
   row+=1
#__________________________________________
sheet7 = wb.add_sheet('Маршрут -6') 
sheet7.write(0, 2, 'Маршрут - 6')
sheet7.write(0, 4, 'Дата:', style=style1)
sheet7.write(0, 5, new_date, style=style)
sheet7.write(2, 0, 'Название точек:')
sheet7.write(2, 1, 'Щелково, Ивантеевка-бережок', style=style1)
sheet7.write(4, 0, 'позиции для отгрузки', style=style1)
sheet7.write(4, 1, '', style=style1)
sheet7.write(4, 2, 'ед.изм', style=style)
sheet7.write(4, 3, 'Щелково', style=style)
sheet7.write(4, 4, 'Ив бережок', style=style)
sheet7.write(4, 5, 'Всего:', style=style)
if len(batch_data_name_5)>5:
   a=batch_data_name_5[0:55]
   b=batch_data_name_5[55:114]
   c=batch_data_name_5[114:173]
   d=batch_data_name_5[173:232]
   e=batch_data_name_5[232:291]
   row=5
   col=0
   for x in a:   
     sheet7.write(row, 0, x, style=style1)
     sheet7.write(row, 1, '', style=style1)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet7.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet7.write(60, 1, '', style=style1)
     for x in b:
       sheet7.write(row, 0, x, style=style1)
       sheet7.write(row, 1, '', style=style1)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet7.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet7.write(120, 1, '', style=style1)
     for x in c:
       sheet7.write(row, 0, x, style=style1)
       sheet7.write(row, 1, '', style=style1)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet7.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet7.write(180, 1, '', style=style1)
     for x in b:
       sheet7.write(row, 0, x, style=style1)
       sheet7.write(row, 1, '', style=style1)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet7.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet7.write(240, 1, '', style=style1)
     for x in b:
       sheet7.write(row, 0, x, style=style1)
       sheet7.write(row, 1, '', style=style1)
       row+=1
if len(batch_mass_item_b_5)>5:
   a=batch_mass_item_b_5[0:55]
   b=batch_mass_item_b_5[55:114]
   c=batch_mass_item_b_5[114:173]
   d=batch_mass_item_b_5[173:232]
   e=batch_mass_item_b_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet7.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet7.write(60, 3, 'Щелково', style=style)
     row=61
     col=0 
     for y in b:  
       sheet7.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet7.write(120, 3, 'Щелково', style=style)
     row=121
     col=0 
     for y in c:  
       sheet7.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet7.write(181, 3, 'Щелково', style=style)
     row=182
     col=0 
     for y in d:  
       sheet7.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet7.write(242, 3, 'Щелково', style=style)
     row=243
     col=0 
     for y in e:  
       sheet7.write(row, 3, y, style=style)
       row+=1 
#_____________________________
if len(batch_mass_item_i_5)>5:
   a=batch_mass_item_i_5[0:55]
   b=batch_mass_item_i_5[55:114]
   c=batch_mass_item_i_5[114:173]
   d=batch_mass_item_i_5[173:232]
   e=batch_mass_item_i_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet7.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet7.write(60, 2, 'ед.изм', style=style)
     row=61
     col=0 
     for y in b:  
       sheet7.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet7.write(120, 2, 'ед.изм', style=style)
     row=121
     col=0 
     for y in c:  
       sheet7.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet7.write(180, 2, 'ед.изм', style=style)
     row=181
     col=0 
     for y in d:  
       sheet7.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet7.write(240, 2, 'ед.изм', style=style)
     row=241
     col=0 
     for y in e:  
       sheet7.write(row, 2, y, style=style)
       row+=1 
#____________________________________


if len(batch_mass_item_c_5)>5:
   a=batch_mass_item_c_5[0:55]
   b=batch_mass_item_c_5[55:114]
   c=batch_mass_item_c_5[114:173]
   d=batch_mass_item_c_5[173:232]
   e=batch_mass_item_c_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet7.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet7.write(60, 4, 'Ив бережок', style=style)
     row=61
     col=0 
     for y in b:  
       sheet7.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet7.write(120, 4, 'Ив бережок', style=style)
     row=121
     col=0 
     for y in c:  
       sheet7.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet7.write(180, 4, 'Ив бережок', style=style)
     row=181
     col=0 
     for y in d:  
       sheet7.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet7.write(240, 4, 'Ив бережок', style=style)
     row=241
     col=0 
     for y in e:  
       sheet7.write(row, 4, y, style=style)
       row+=1 
if len(batch_mass_item_f_5)>5:
   a=batch_mass_item_f_5[0:55]
   b=batch_mass_item_f_5[55:114]
   c=batch_mass_item_f_5[114:173]
   d=batch_mass_item_f_5[173:232]
   e=batch_mass_item_f_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet7.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet7.write(60, 5, 'Всего', style=style)
     row=61
     col=0 
     for y in b:  
       sheet7.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet7.write(120, 5, 'Всего', style=style)
     row=121
     col=0 
     for y in c:  
       sheet7.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet7.write(180, 5, 'Всего', style=style)
     row=181
     col=0 
     for y in d:  
       sheet7.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet7.write(240, 5, 'Всего', style=style)
     row=241
     col=0 
     for y in e:  
       sheet7.write(row, 5, y, style=style)
       row+=1 
   row+=2
   sheet7.write(row, 0, 'Получил:')
   sheet7.write(row, 2, '________________/_______________/')
   row+=1
#_________________________________________
sheet8 = wb.add_sheet('Общий заказ') 
sheet8.write(0, 1, 'Общий заказ')
sheet8.write(0, 3, 'Дата:', style=style1)
sheet8.write(0, 4, new_date, style=style)
sheet8.write(2, 0, 'Всего во всех заказах:', style=style1)
sheet8.write(4, 0, 'Наименование товара', style=style1)
sheet8.write(4, 1, '', style=style1)
sheet8.write(4, 2, 'ед.изм', style=style)
sheet8.write(4, 3, 'Всего:', style=style)

if len(batch_data_name_6)>5:
   a=batch_data_name_6[0:55]
   b=batch_data_name_6[55:114]
   c=batch_data_name_6[114:173]
   d=batch_data_name_6[173:232]
   e=batch_data_name_6[232:291]
   row=5
   col=1
   for x in a:  
     sheet8.write(row, 0, x, style=style1)
     row+=1  
   row=61
   col=1
   if len(b)!=0:
       sheet8.write(60, 0, 'Наименование товара', style=style1)
       for x in b:  
         sheet8.write(row, 0, x, style=style1)
         row+=1    
   row=121
   col=1
   if len(c)!=0:
      sheet8.write(120, 0, 'Наименование товара', style=style1)
      for x in c:  
        sheet8.write(row, 0, x, style=style1)
        row+=1 
   
   row=181
   col=1   
   if len(d)!=0:
      sheet8.write(180, 0, 'Наименование товара', style=style1)
      for x in d:  
        sheet8.write(row, 0, x, style=style1)
        row+=1 
   row=241
   col=1
   if len(e)!=0:
      sheet8.write(240, 0, 'Наименование товара', style=style1)
      for x in e:  
        sheet8.write(row, 0, x, style=style1)
        row+=1 
#_____________________________
if len(batch_mass_item_i_6)>5:
   a=batch_mass_item_i_6[0:55]
   b=batch_mass_item_i_6[55:114]
   c=batch_mass_item_i_6[114:173]
   d=batch_mass_item_i_6[173:232]
   e=batch_mass_item_i_6[232:291]  
   row=5
   col=0
   for y in a:
    sheet8.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet8.write(60, 2, 'ед.изм', style=style)
     row=61
     col=0 
     for y in b:  
       sheet8.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet8.write(120, 2, 'ед.изм', style=style)
     row=121
     col=0 
     for y in c:  
       sheet8.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet8.write(180, 2, 'ед.изм', style=style)
     row=181
     col=0 
     for y in d:  
       sheet8.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet8.write(240, 2, 'ед.изм', style=style)
     row=241
     col=0 
     for y in e:  
       sheet8.write(row, 2, y, style=style)
       row+=1 
#____________________________________


if len(batch_mass_item_f_6)>5:
   a=batch_mass_item_f_6[0:55]
   b=batch_mass_item_f_6[55:114]
   c=batch_mass_item_f_6[114:173]
   d=batch_mass_item_f_6[173:232]
   e=batch_mass_item_f_6[232:291]
   row=5
   col=0
   for y in a:
     sheet8.write(row, 1, '', style=style1)
     sheet8.write(row, 3, y, style=style)
     row+=1  
   if len(b)!=0:   
     sheet8.write(60, 3, 'Всего', style=style)
     row=61
     col=0 
     for y in b:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1 
   if len(c)!=0:   
     sheet8.write(120, 3, 'Всего', style=style)
     row=121
     col=0 
     for y in c:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1 
  
   if len(d)!=0:   
     sheet8.write(180, 3, 'Всего', style=style)
     row=181
     col=0 
     for y in d:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1 
   if len(e)!=0:   
     sheet8.write(240, 3, 'Всего', style=style)
     row=241
     col=0 
     for y in e:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1
wb.save('Main_list.xls')
wb=0
#_________________________________________
#__________Накладные______________________
"""
spreadsheet_id='1XwKY9L_t9FkakTzWEJZvmI2wrvF3Gq02OOztQbfojzI' #  Кор Горького 6
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data, batch_id, batch_mass_t1) 
#________________________________________
time.sleep(12)
spreadsheet_id='1xbuMghlorBXudBndZc__lwSCl_6_OCjGgAmCfd1FnOU' #  Пушкино Московский проспект
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t2, batch_id_t2, batch_mass_t2) 
#________________________________________
time.sleep(12)
spreadsheet_id='1q5DYpGOvf9rHWlnddcGANT6FXyySfnH06FZwUOOGXOg' #  Пушкино Чехова 40(7)
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t3, batch_id_t3, batch_mass_t3)    
#________________________________________
time.sleep(12)
spreadsheet_id='1hCeOXfzX1hXLhuXH0JNPr9NJTpFmdlVzEGF3cSKmPz8' # Пушкино Облака
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t4, batch_id_t4, batch_mass_t4) 
#_______________________________________
time.sleep(12)
spreadsheet_id='1vCfQOwsLesOebcmPpCibVSq3Scdpv-3rVLALShXC85Y' # С-Посад Скобянка
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t5, batch_id_t5, batch_mass_t5) 
#________________________________________
time.sleep(20)
spreadsheet_id='1rdPPa077Yq3ZqBV2fncc7rf2zDUmdo7XxpqCw5LuD1E' # С-Посад Кр-армии-203
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t6, batch_id_t6, batch_mass_t6) 
#_______________________________________
time.sleep(20)
spreadsheet_id='1MLW3APEbrgbsoOxflKys0BHkujxO8U4JUPCC9ur42bk' # С-Посад Кр-армии-182
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t7, batch_id_t7, batch_mass_t7) 
#_______________________________________
time.sleep(20)
spreadsheet_id='1KMIa8wr1CMDUjSa28718jfYHJl1Q1nqD6GMtV0q7Vpo' # С-Посад Кр-армии-185
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t8, batch_id_t8, batch_mass_t8)
#_______________________________________
time.sleep(20)
spreadsheet_id='12tg-IxF-BSolMO6zkw-Ln_C8cufoojlnKmemdfwLJ3E' # Ивантеевка Н-Слобода
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t9, batch_id_t9, batch_mass_t9)
#______________________________________
time.sleep(30)
spreadsheet_id='1_tq3s2I7VKtyGKdoM50LLWCejBhkZ4QMHFAU3_PhVL4' # Черноголовка
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t10, batch_id_t10, batch_mass_t10)
#_______________________________________
time.sleep(15)
spreadsheet_id='1PR73zzrZo_rwdwJ6COsDZRNnRnIhS1h0Vc0w9n0UB4o' # Черноголовка-2
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t11, batch_id_t11, batch_mass_t11)

#_______________________________________
time.sleep(12)
spreadsheet_id='1J7RpwGemREI8Tg30lxNiaMtLC9eUgd2S0I0RgujlGcg' # Софрино
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t13, batch_id_t13, batch_mass_t13)
#_______________________________________
time.sleep(15)
spreadsheet_id='1mTZ1csIJLEuV-WhOtQTe3iwty2erzLqbrWKYMDBwot8' # С-Посад Кр-армии-3
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t14, batch_id_t14, batch_mass_t14)
#_______________________________________ 
time.sleep(15)
spreadsheet_id='1QGdTqZETBIEWEoZhFMv-1MLIf7gQmpfqJ2QdwQ30YaE' # С-Посад Кр-армии-5
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t21, batch_id_t21, batch_mass_t21)
#_______________________________________
time.sleep(15)
spreadsheet_id='1z6NtEqEqqjEFt_1D0q2LZvRe65SY4203emI6pm0lKJ4' # С-Посад Рыбная
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t18, batch_id_t18, batch_mass_t18)
#_______________________________________
time.sleep(15)
spreadsheet_id='1751ZxnYMd6DeF4N0cMJtgxdFJczsOSnNfR-htG7bKWo' # Интернет магазин
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t19, batch_id_t19, batch_mass_t19)
#_______________________________________
time.sleep(15)
spreadsheet_id='1uG2FeDbfVV_joigtBVSRrdCjcBrvAc92qIE2f81sU2s' # Заветы Ильича
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t15, batch_id_t15, batch_mass_t15)
#_______________________________________
time.sleep(15)
spreadsheet_id='1D6RZEaAQrjv1hvnGuXVRePd6Deai8GsDpn9qmBRcHFo' # Пушкино Арманд
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t16, batch_id_t16, batch_mass_t16)
#_______________________________________
time.sleep(15)
spreadsheet_id='1yyeZp-wepxKBKtNFf8GSBu-rOfRMdKEJfLfUdabgeUQ' # Пушкино Тургенева 18
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t17, batch_id_t17, batch_mass_t17)
#_______________________________________
time.sleep(15)
spreadsheet_id='1XImDDCBuiaj6MdS7wiDhyQs-34rkaRs2lKFCKCZtbPw' # Щелково
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t20, batch_id_t20, batch_mass_t20)
#_______________________________________ 
time.sleep(15)
spreadsheet_id='1zfncsIHfH8NpaxM3vwvmcJU3FYi3j7SwtXEmMYPLU00' # Ивантеевка - Бережок
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "B3:B").execute( )#format table
resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
google_table(batch_data_t12, batch_id_t12, batch_mass_t12)
"""
#_______________________________________
print ('Записываю списание')
#__________________________________________________________________connect to table -1
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1SRWBpWfQkYgpOwnsc_nCto5YfmF5gqcGhUdcPlzfuCg'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_csv(spreadsheet_id, 'Sheet1')
time.sleep(4)
#_______________________________________________________________________________
a=1
if a==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['korolev_gor_6']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-1']
    c = i['pay_cost-1']
    d = i['sum-1']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
b=1
if b==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['pushkino_mosk_prosp']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-2']
    c = i['pay_cost-2']
    d = i['sum-2']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
c=1
if c==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['pushkino_cheh_40']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-3']
    c = i['pay_cost-3']
    d = i['sum-3']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
d=1
if d==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['pushkino_oblaka']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-4']
    c = i['pay_cost-4']
    d = i['sum-4']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________

d1=1
if d1==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  goods=[]
  quantity_1=[]
  pay_cost_1=[]
  sum_1=[]
  quantity_2=[]
  pay_cost_2=[]
  sum_2=[]
  quantity_3=[]
  pay_cost_3=[]
  sum_3=[]
  quantity_4=[]
  pay_cost_4=[]
  sum_4=[]
  for i in reader:
    a = i['name_id']
    a1 = i['qauntity-1']
    b = i['pay_cost-1']
    c = i['sum-1']
    d = i['qauntity-2']
    e = i['pay_cost-2']
    f = i['sum-2']
    g = i['qauntity-3']
    h = i['pay_cost-3']
    j = i['sum-3']
    g = i['qauntity-4']
    h = i['pay_cost-4']
    j = i['sum-4']
    if a1!='':
      goods.append(str(a))
      quantity_1.append(str(a1))
      pay_cost_1.append(str(b))
      sum_1.append(str(c))
      quantity_2.append(str(d))
      pay_cost_2.append(str(e))
      sum_2.append(str(f))
      quantity_3.append(str(g))
      pay_cost_3.append(str(h))
      sum_3.append(str(j))
      quantity_4.append(str(g))
      pay_cost_4.append(str(h))
      sum_4.append(str(j))
  spreadsheet_id='1TBIiujdDhwqAwGUtTE8DqHNUICLNuUhHIED4XLybox8' # Списание маршрут -1 
  credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
  httpAuth = credentials.authorize(httplib2.Http())
  service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "D3:D").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "E3:E").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "F3:F").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "G3:G").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "H3:H").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "I3:I").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "J3:J").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "K3:K").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "L3:L").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "M3:M").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "N3:N").execute( )#format table
  time.sleep(4)
  google_table_post_scrapp_4_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2, quantity_3, pay_cost_3, sum_3, quantity_4, pay_cost_4, sum_4)

#__________________________________________________________________connect to table -2
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1zQ4KY_9baq1Zb-EiYSx3ca9h2-CZSdA1XWX2NsvjA0Q'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_csv(spreadsheet_id, 'Sheet1')
time.sleep(4)
#_______________________________________________________________________________
g=1
if g==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sposad_skobyanka']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-5']
    c = i['pay_cost-5']
    d = i['sum-5']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
h=1
if h==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sposad_armii_182']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-6']
    c = i['pay_cost-6']
    d = i['sum-6']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
i=1
if i==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sposad_armii_203']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-7']
    c = i['pay_cost-7']
    d = i['sum-7']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
j=1
if j==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sposad_armii_185']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-8']
    c = i['pay_cost-8']
    d = i['sum-8']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
e1=1
if e1==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  goods=[]
  quantity_1=[]
  pay_cost_1=[]
  sum_1=[]
  quantity_2=[]
  pay_cost_2=[]
  sum_2=[]
  quantity_3=[]
  pay_cost_3=[]
  sum_3=[]
  quantity_4=[]
  pay_cost_4=[]
  sum_4=[]
  for i in reader:
    a = i['name_id']
    a1 = i['qauntity-5']
    b = i['pay_cost-5']
    c = i['sum-5']
    d = i['qauntity-6']
    e = i['pay_cost-6']
    f = i['sum-6']
    g = i['qauntity-7']
    h = i['pay_cost-7']
    j = i['sum-7']
    g = i['qauntity-8']
    h = i['pay_cost-8']
    j = i['sum-8']
    if a1!='':
      goods.append(str(a))
      quantity_1.append(str(a1))
      pay_cost_1.append(str(b))
      sum_1.append(str(c))
      quantity_2.append(str(d))
      pay_cost_2.append(str(e))
      sum_2.append(str(f))
      quantity_3.append(str(g))
      pay_cost_3.append(str(h))
      sum_3.append(str(j))
      quantity_4.append(str(g))
      pay_cost_4.append(str(h))
      sum_4.append(str(j))
  spreadsheet_id='1TfNIuyO1iAviof6zpyVtZfcMWNLy1MBaveqXp9XITxA' # Списание маршрут -2 
  credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
  httpAuth = credentials.authorize(httplib2.Http())
  service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "D3:D").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "E3:E").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "F3:F").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "G3:G").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "H3:H").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "I3:I").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "J3:J").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "K3:K").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "L3:L").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "M3:M").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "N3:N").execute( )#format table
  time.sleep(4)
  google_table_post_scrapp_4_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2, quantity_3, pay_cost_3, sum_3, quantity_4, pay_cost_4, sum_4)

#__________________________________________________________________connect to table -3
print('api - пауза на 30 sec')
time.sleep(30)
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1aGZmaFtx9HD8oItEW_at4vu7oWI7nGQtsjW3h_KdPCA'# table adress
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http()) # api validation
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_csv(spreadsheet_id, 'Sheet1') # dowload csv
time.sleep(4)
#________________________________________________________________ # parse for trade shop
e=1
if e==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['ivanteevka_sloboda']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']          # table to db - 1 element (record)
    b = i['qauntity-9']
    c = i['pay_cost-9']
    d = i['sum-9']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
f=1
if f==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['chernogolovka']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-10']
    c = i['pay_cost-10']
    d = i['sum-10']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
h=1
if h==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['chernogolovka_2']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-11']
    c = i['pay_cost-11']
    d = i['sum-11']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
o1=1
if o1==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  goods=[]
  quantity_1=[]
  pay_cost_1=[]
  sum_1=[]
  quantity_2=[]
  pay_cost_2=[]
  sum_2=[]
  quantity_3=[]
  pay_cost_3=[]
  sum_3=[]
  for i in reader:
    a = i['name_id']
    a1 = i['qauntity-9']
    b = i['pay_cost-9']
    c = i['sum-9']
    d = i['qauntity-10']
    e = i['pay_cost-10']
    f = i['sum-10']
    g = i['qauntity-11']
    h = i['pay_cost-11']
    j = i['sum-11']
    if a1!='':
      goods.append(str(a))
      quantity_1.append(str(a1))
      pay_cost_1.append(str(b))
      sum_1.append(str(c))
      quantity_2.append(str(d))
      pay_cost_2.append(str(e))
      sum_2.append(str(f))
      quantity_3.append(str(g))
      pay_cost_3.append(str(h))
      sum_3.append(str(j))
  spreadsheet_id='1Kn-ezNTCRNadYgQbnhB9NzJBINwdyjtiAqAoiO0mmAY' # Списание маршрут -3 
  credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
  httpAuth = credentials.authorize(httplib2.Http())
  service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "D3:D").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "E3:E").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "F3:F").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "G3:G").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "H3:H").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "I3:I").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "J3:J").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "K3:K").execute( )#format table
  time.sleep(4)
  google_table_post_scrapp_3_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2, quantity_3, pay_cost_3, sum_3)
#__________________________________________________________________connect to table -4
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1oWVD-JwJvV8eh-0jE1z8cZA7vN4OManYJWpCeeeIdfA'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_csv(spreadsheet_id, 'Sheet1')
time.sleep(4)

#_______________________________________________________________________________
k=1
if k==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sofrino']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-12']
    c = i['pay_cost-12']
    d = i['sum-12']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
l=1
if l==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sposad_ribnay']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-13']
    c = i['pay_cost-13']
    d = i['sum-13']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
m=1
if m==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sposad_armii_5']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-14']
    c = i['pay_cost-14']
    d = i['sum-14']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
n=1
if n==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['sposad_armii_3']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-15']
    c = i['pay_cost-15']
    d = i['sum-15']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________

o1=1
if o1==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  goods=[]
  quantity_1=[]
  pay_cost_1=[]
  sum_1=[]
  quantity_2=[]
  pay_cost_2=[]
  sum_2=[]
  quantity_3=[]
  pay_cost_3=[]
  sum_3=[]
  quantity_4=[]
  pay_cost_4=[]
  sum_4=[]
  for i in reader:
    a = i['name_id']
    a1 = i['qauntity-12']
    b = i['pay_cost-12']
    c = i['sum-12']
    d = i['qauntity-13']
    e = i['pay_cost-13']
    f = i['sum-13']
    g = i['qauntity-14']
    h = i['pay_cost-14']
    j = i['sum-14']
    g = i['qauntity-15']
    h = i['pay_cost-15']
    j = i['sum-15']
    if a1!='':
      goods.append(str(a))
      quantity_1.append(str(a1))
      pay_cost_1.append(str(b))
      sum_1.append(str(c))
      quantity_2.append(str(d))
      pay_cost_2.append(str(e))
      sum_2.append(str(f))
      quantity_3.append(str(g))
      pay_cost_3.append(str(h))
      sum_3.append(str(j))
      quantity_4.append(str(g))
      pay_cost_4.append(str(h))
      sum_4.append(str(j))
  spreadsheet_id='1wB5LYk5I49AgL_7JLixkxF76ORGPFEhAqx4DKwP2Ztg' # Списание маршрут -4 
  credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
  httpAuth = credentials.authorize(httplib2.Http())
  service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "D3:D").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "E3:E").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "F3:F").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "G3:G").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "H3:H").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "I3:I").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "J3:J").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "K3:K").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "L3:L").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "M3:M").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "N3:N").execute( )#format table
  time.sleep(4)
  google_table_post_scrapp_4_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2, quantity_3, pay_cost_3, sum_3, quantity_4, pay_cost_4, sum_4)

#__________________________________________________________________connect to table -5
print('api - пауза на 30 sec')
time.sleep(30)
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1xKahPITIZ0Q5C6j7XWPBOQVD_zDhahbbL7BMhp3WojY'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_csv(spreadsheet_id, 'Sheet1')
time.sleep(4)
#_______________________________________________________________________________
o=1
if o==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['internet_magazin']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-16']
    c = i['pay_cost-16']
    d = i['sum-16']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
p=1
if p==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['zavety_ilicha']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-17']
    c = i['pay_cost-17']
    d = i['sum-17']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
q=1
if q==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['pushkino_armand']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-18']
    c = i['pay_cost-18']
    d = i['sum-18']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
r=1
if r==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['pushkino_turgeneva']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-19']
    c = i['pay_cost-19']
    d = i['sum-19']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
o1=1
if o1==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  goods=[]
  quantity_1=[]
  pay_cost_1=[]
  sum_1=[]
  quantity_2=[]
  pay_cost_2=[]
  sum_2=[]
  quantity_3=[]
  pay_cost_3=[]
  sum_3=[]
  quantity_4=[]
  pay_cost_4=[]
  sum_4=[]
  for i in reader:
    a = i['name_id']
    a1 = i['qauntity-16']
    b = i['pay_cost-16']
    c = i['sum-16']
    d = i['qauntity-17']
    e = i['pay_cost-17']
    f = i['sum-17']
    g = i['qauntity-18']
    h = i['pay_cost-18']
    j = i['sum-18']
    g = i['qauntity-19']
    h = i['pay_cost-19']
    j = i['sum-19']
    if a1!='':
      goods.append(str(a))
      quantity_1.append(str(a1))
      pay_cost_1.append(str(b))
      sum_1.append(str(c))
      quantity_2.append(str(d))
      pay_cost_2.append(str(e))
      sum_2.append(str(f))
      quantity_3.append(str(g))
      pay_cost_3.append(str(h))
      sum_3.append(str(j))
      quantity_4.append(str(g))
      pay_cost_4.append(str(h))
      sum_4.append(str(j))
  spreadsheet_id='1V6_st3TWHkI9khPyFsxi_5-r1hWt5ZL_1YvjOHUvCTY' # Списание маршрут -5 
  credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
  httpAuth = credentials.authorize(httplib2.Http())
  service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "D3:D").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "E3:E").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "F3:F").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "G3:G").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "H3:H").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "I3:I").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "J3:J").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "K3:K").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "L3:L").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "M3:M").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "N3:N").execute( )#format table
  time.sleep(4)
  google_table_post_scrapp_4_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2, quantity_3, pay_cost_3, sum_3, quantity_4, pay_cost_4, sum_4)

#__________________________________________________________________connect to table -6
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1XJCQR75Trv2c4tFP1kLYZ29sWyXVekGefSovixvLG4A'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_csv(spreadsheet_id, 'Sheet1')
time.sleep(2)
#_______________________________________________________________________________
s=1
if s==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['schelkovo']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-20']
    c = i['pay_cost-20']
    d = i['sum-20']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
t=1
if t==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  name=['ivanteevka_beregok']
  data_time=[date]
  goods=[]
  quantity=[]
  pay_cost=[]
  sum_1=[]
  for i in reader:
    a = i['name_id']
    b = i['qauntity-21']
    c = i['pay_cost-21']
    d = i['sum-21']
    if b !='':
      goods.append(str(a))
      quantity.append(str(b))
      pay_cost.append(str(c))
      sum_1.append(str(d))
  f_1.close()
  sqlite_insert(name,data_time,goods,quantity,pay_cost,sum_1)
#_________________________________________________________________________________
o1=1
if o1==1:
  f_1 = open('scrapping_values.csv', 'r', newline='')
  reader= csv.DictReader(f_1, delimiter=',')
  goods=[]
  quantity_1=[]
  pay_cost_1=[]
  sum_1=[]
  quantity_2=[]
  pay_cost_2=[]
  sum_2=[]
  for i in reader:
    a = i['name_id']
    a1 = i['qauntity-20']
    b = i['pay_cost-20']
    c = i['sum-20']
    d = i['qauntity-21']
    e = i['pay_cost-21']
    f = i['sum-21']
    if a1!='':
      goods.append(str(a))
      quantity_1.append(str(a1))
      pay_cost_1.append(str(b))
      sum_1.append(str(c))
      quantity_2.append(str(d))
      pay_cost_2.append(str(e))
      sum_2.append(str(f))
  spreadsheet_id='1VbQVcLmiHy8uZuMBaEX65kEhZfV83shGi5jhhoIC9Cw' # Списание маршрут -6 
  credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
  httpAuth = credentials.authorize(httplib2.Http())
  service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "A3:A").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "C3:C").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "D3:D").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "E3:E").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "F3:F").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "G3:G").execute( )#format table
  resultClear = service.spreadsheets( ).values( ).clear(spreadsheetId=spreadsheet_id, range= "H3:H").execute( )#format table
  time.sleep(4)
  google_table_post_scrapp_2_elem(goods, quantity_1, pay_cost_1, sum_1, quantity_2, pay_cost_2, sum_2)
#_________________________________________________________________________________


print ('ver.: 1.41')
print ('Готово, записал все значения в таблицу и сформировал заказ на склад!!!')
print ('Маршруты и накладные сформированны')
print ('Общий заказ в и расчет по маршрутам в Main_list.xls')
print ('Накладные в GoogleDocs')

