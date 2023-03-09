import xlwt
import re
from xlwt import Workbook
from datetime import date

def writeData():

    now = date.today()
    wb = Workbook()
    style_string = 'font:bold on; borders: bottom dashed; borders: top dotted; borders: right dashed;'
    style=xlwt.easyxf(style_string)
    """create sheet"""
    sheet = wb.add_sheet('Links Result', cell_overwrite_ok=True)
    sheet.col(0).width = 13000
    sheet.col(1).width = 3500
    sheet.col(2).width = 8000
    sheet.col(3).width = 8000
    sheet.col(4).width = 6000
    sheet.col(5).width = 6000
    sheet.col(6).width = 6000

    sheet.write(0, 0, 'Ссылка', style = style)
    sheet.write(0, 1, 'Дата мониторинга', style = style)
    sheet.write(0, 2, 'Форма WB', style = style)
    sheet.write(1, 2, 'дата об отправке/результат')
    sheet.write(0, 3, 'Email WB', style = style)
    sheet.write(1, 3, 'дата об отправке/результат')
    sheet.write(0, 4, 'Претензия', style = style)
    sheet.write(1, 4, 'номер')
    sheet.write(0, 5, 'Претензия нарушителю', style = style)
    sheet.write(1, 5, 'дата об отправке/результат')
    sheet.write(0, 6, 'Заказное нарушителю', style = style)
    sheet.write(1, 6, 'дата об отправке/результат')

    with open("resultData.json", "r") as reader:
        while True:
            info = reader.readline()
            crudeCount = re.findall("nt':.\d*,", info)
            crudeFile = re.findall("le':.*\.pdf',", info)
            crudeLink = re.findall("nk':..*\.aspx", info)
            newCount = ''.join(crudeCount).replace('"', '').replace('nt\': ','').replace(',','')
            newFile = ''.join(crudeFile).replace('le\': ', '').replace('\'','').replace(',','')
            newLink = ''.join(crudeLink).replace('nk\': ', '').replace('\'','').replace(',','')

            if len(newCount) != 0:
                count = int(newCount)
                file = newFile
                link = newLink

            sheet.write(count, 0, link)
            sheet.write(count, 1, f'{now}')
            sheet.write(count, 2, f'{now}/fill-success')
            sheet.write(count, 3, f'{now}/send-success')
            sheet.write(count, 4, file)
            if not info:
                break

    """write result"""

    # sheet.write(count, 5, f'{now}/send-success')


    """save result"""
    wb.save('Local_result.xls')



