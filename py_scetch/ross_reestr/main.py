import requests
import re
import xlwt
from xlwt import Workbook

'https://reestr.rublacklist.net/?status=1&gov=30&paginate_by=100' #Rossmolodezh
'https://reestr.rublacklist.net/?status=1&gov=33&paginate_by=100' #Rosszdrav
'https://reestr.rublacklist.net/?status=1&gov=5&paginate_by=100' #Genprok
'https://reestr.rublacklist.net/?status=1&gov=28&paginate_by=100' #Rossalcogol
'https://reestr.rublacklist.net/?status=1&gov=all&paginate_by=500' #allsite

'https://reestr.rublacklist.net/ru/?page=1&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=2&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=3&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=4&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=5&status=1&gov=all&paginate_by=100',

'https://reestr.rublacklist.net/ru/?page=6&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=7&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=8&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=9&status=1&gov=all&paginate_by=100',

urls =[
'https://reestr.rublacklist.net/ru/?page=1&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=2&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=3&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=4&status=1&gov=all&paginate_by=100',
'https://reestr.rublacklist.net/ru/?page=5&status=1&gov=all&paginate_by=100',






      ]


list_to_href = []
list_rec = []

def url_lists(urls):
    for url in urls:
        r = requests.get(url)
        resp = r.text
        resp_fragment = re.findall('<div class="table_td td_site">\n.*\n.*\n<\/a>\n<\/div>\n.*\n.*\n.*\n<\/a>', resp)
        str_fragmet = ''.join(resp_fragment)
        rec = re.findall('<a href=\"\/ru\/record\/\d{7}\/', str_fragmet)
        for item_rec in rec:
            str_rec = ''.join(item_rec).replace('<a href="','')
            record = 'https://reestr.rublacklist.net' + str_rec
            list_to_href.append(record)
            [list_rec.append(x) for x in list_to_href if x not in list_rec]

title_list = []
cont_tel_1 =[]
cont_tel_2 =[]
cont_tel_3 =[]
cont_tel_4 =[]
cont_tel_5 =[]
cont_mail = []
cont_gramm =[]
meta_root_link=[]
meta_stat = []  #status
meta_domain = [] #domain
meta_ip = []  # ip
meta_state = [] # resolution
meta_date = []  #date des


def get_contact(lst_url):
    for url in lst_url:
         try:
            r = requests.get(url)
            resp = r.text
            getDomain = re.findall('ON\n.*\n.*\n.*\n.*\n.*\n.*\n.*\n.*\n.*\n.*\n.*',resp)
            str_getDomain = ''.join(getDomain).replace('Статус','').replace('<div class="td_title">','').replace('Домен</div>\n<div class="td_content">','').replace('</div>','').replace('<div class="table_wrapper">','').replace('ON','').replace('<div class="table_row">','').replace('<div class="td_title">Статус','').replace('<div class="td_content">Заблокирован','').replace('<div class="table_row">','').replace('\n','')
            full_url = 'https://'+str_getDomain
            print(full_url)
            res = requests.get(full_url)
            dom_res = res.text
            get_title = re.findall('<title>.*<\/title>',dom_res)
            get_title_str = ''.join(get_title).replace('</title>', '').replace('<title>', '->')
            get_tel_1 = re.findall('[:]\d\d\d{9}\"[ ]', dom_res)
            get_t_1 = ''.join(get_tel_1)[0:20].replace(':','').replace('"',' ')
            get_tel_2 = re.findall('\d.[(]\d*[)].{10}', dom_res)
            get_t_2 =''.join(get_tel_2)[:30].replace('(999) 999-99-999', '').replace('9','').replace('(999)', '')
            get_tel_5 = re.findall(r'\d.\(\d..\).{10}', dom_res)
            get_t_5 =''.join(get_tel_5)[:30].replace('(999) 999-99-999', '').replace('9','').replace('(999)', '')
            get_tel_3 = re.findall('tel:\d{11}\">', dom_res)
            get_t_3 = ''.join(get_tel_3)[0:20].replace('class','').replace('tel:','').replace('">tel','')
            get_tel_4 = re.findall(':[+]\d.*\"', dom_res)
            get_t_4 = ''.join(get_tel_4)[0:20].replace('tel:','').replace('">tel','').replace('class','')
            get_mail_1 = re.findall('mailto[:].*[@].*[.].{5}.*\">', dom_res)
            get_mail_res = ''.join(get_mail_1)[7:30]
            get_gramm = re.findall('https:\/\/t[.]me\/.{25}', dom_res)
            get_telegramm = ''.join(get_gramm)[:40].replace('<',' ').replace('target', '  ').replace('<span class','  ').replace('</a>',' ').replace('target="_blaht',' ').replace('targetht','')
            # # ____________________________________________
            status_fact = re.findall('Статус.*\n.*\n', resp)
            str_status_fact = ''.join(status_fact[0]).replace('</div>\n','').replace('<div class="td_content">','').replace('Статус','')
            dom_ip = re.findall('IP-адрес<\/div>\n.*\n.*<br>\d{1}',resp)
            d_IP = ''.join(dom_ip).replace('IP-адрес</div>\n<div class="td_content">\n','')
            ar_dep_state = re.findall('Орган власти.*\n.*',resp)
            str_dep_state = ''.join(ar_dep_state[0]).replace('Орган власти</div>\n<div class="td_content">', '').replace('</div>','')
            ar_data_of_res = re.findall('Дата первой .*\n.*',resp)
            str_date = ''.join(ar_data_of_res[0]).replace('Дата первой блокировки</div>\n<div class="td_content">','').replace('</div>','')

            meta_root_link.append(url)
            title_list.append(get_title_str)
            meta_stat.append(str_status_fact)
            meta_domain.append(full_url)
            meta_ip.append(d_IP)
            meta_state.append(str_dep_state)
            meta_date.append(str_date)
            cont_tel_1.append(get_t_1)
            cont_tel_2.append(get_t_2)
            cont_tel_3.append(get_t_5)
            cont_tel_4.append(get_t_3)
            cont_tel_5.append(get_t_4)
            cont_mail.append(get_mail_res)
            cont_gramm.append(get_telegramm)
         except:
             err = 'access_error'


url_lists(urls)
get_contact(list_rec)

wb = Workbook()
wb=Workbook()
style_string = 'font:bold on; borders: bottom dashed; borders: top dotted; borders: right dashed;'
style=xlwt.easyxf(style_string)
#____________________________________________________________
sheet1 = wb.add_sheet('Block sheet')
sheet1.write(0, 0, 'Список доменов')
sheet1.write(0, 1, 'Статус')
sheet1.write(0, 2, ' IP адрес')
sheet1.write(0, 3, 'Орган Власти')
sheet1.write(0, 4, 'Дата Блока')
sheet1.write(0, 5, 'возм.Тел:')
sheet1.write(0, 6, 'возм.Тел:')
sheet1.write(0, 7, 'возм.Тел:')
sheet1.write(0, 8, 'возм.Тел:')
sheet1.write(0, 9, 'возм.Тел:')
sheet1.write(0, 10, 'возм.Email:')
sheet1.write(0, 11, 'Телеграмм:')
sheet1.write(0, 12, 'Заголовок сайта ->')
sheet1.write(0, 18, 'Ссылка в реестре')

row=3
col=1
for x in meta_domain:
  sheet1.write(row, 0, x, style =style)
  row+=1
row=3
col=1
for x in meta_stat:
  sheet1.write(row, 1, x, style =style)
  row+=1
row=3
col=1
for x in meta_ip:
  sheet1.write(row, 2, x, style =style)
  row+=1
row=3
col=1
for x in meta_state:
  sheet1.write(row, 3, x, style =style)
  row+=1
row=3
col=1
for x in meta_date:
  sheet1.write(row, 4, x, style =style)
  row+=1
row=3
col=1

for x in cont_tel_1:
  sheet1.write(row, 5, x, style =style)
  row+=1
row=3
col=1
for x in cont_tel_2:
  sheet1.write(row, 6, x, style =style)
  row+=1
row=3
col=1
for x in cont_tel_3:
  sheet1.write(row, 7, x, style =style)
  row+=1
row=3
col=1
for x in cont_tel_4:
  sheet1.write(row, 8, x, style =style)
  row+=1
row=3
col=1
for x in cont_tel_5:
  sheet1.write(row, 9, x, style =style)
  row+=1
row=3
col=1
for x in cont_mail:
  sheet1.write(row, 10, x, style =style)
  row+=1
row=3
col=1
for x in cont_gramm:
  sheet1.write(row, 11, x, style =style)
  row+=1
row=3
col=1
for x in title_list:
  sheet1.write(row, 12, x, style =style)
  row+=1
row=3
col=1
for x in meta_root_link:
  sheet1.write(row, 18, x, style =style)
  row+=1
row=3
col=1

wb.save('Block_list.xls')
wb=0
