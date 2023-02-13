import requests
import re
import docx
from datetime import datetime
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

from mailSend import sendEmail


ulrList = [
    'https://www.wildberries.ru/catalog/90053004/detail.aspx',
    'https://www.wildberries.ru/catalog/98529036/detail.aspx',
    'https://www.wildberries.ru/catalog/140970297/detail.aspx'
]


def createClaim(item):
    myDoc = docx.Document()
    current_datetime = datetime.now()
    linkIndexPattern = re.findall('\d.*/', item)
    linkIndex = ''.join(linkIndexPattern).replace('/','')
    apiCard = f'https://card.wb.ru/cards/detail?spp=0&regions=80,64,83,4,38,33,70,68,69,86,75,30,40,48,1,22,66,31,71&pricemarginCoeff=1.0&reg=0&appType=1&emp=0&locale=ru&lang=ru&curr=rub&couponsGeo=12,3,18,15,21&dest=-1029256,-102269,-2162196,-1257786&nm={linkIndex}'
    resp = requests.get(apiCard)
    result = resp.json()
    brandName = result['data']['products'][0]['brand']
    myDoc.add_heading('                                   Претензия.', 2)
    myDoc.add_paragraph(f' ')
    myDoc.add_paragraph(f"Претензия о неправомерном использовании товарного знака Мне (Нам) стало известно"
                        f" о незаконном использовании Вами/Вашей организацией зарегистрированного мной (нами)"
                        f" товарного знака в виде "
                        f"{brandName} (далее - товарный знак).Товарный знак был обнаружен \"{item}\". "
                        f"Данные обстоятельства подтверждаются [вписать нужное].")

    myDoc.add_paragraph(f"В соответствии со ст. 1229,1484ГК РФ лицу, на имя которого зарегистрирован товарный знак, "
                        f"принадлежит исключительное право использования товарного знака любым не противоречащим закону способом. "
                        f"Использование товарного знака без согласия правообладателя является незаконным и влечет ответственность. "
                        f"За действия по незаконному использованию товарного знака предусмотрена гражданско-правовая (ст.1515ГК РФ), "
                        f"административная (ст. 14.10 КоАП РФ) и уголовная (ст. 180 УК РФ) ответственность. "
                        f"[Наименование организации/ИП - правообладателя] является правообладателем товарного знака "
                        f"на основании свидетельства N [значение]. При этом Вам не было предоставлено право использования товарного знака "
                        f"по лицензионному договору.")

    myDoc.add_paragraph(f'На основании ст. 1252 ГК РФ требую (требуем) прекратить использование товарного знака'
                        f' [указать способ использования]. Кроме того, требую (требуем) возместить мне'
                        f' (нашей организации) убытки в размере [сумма цифрами и прописью] рублей. '
                        f'Расчет суммы произведен на основании следующих данных: [вписать нужное]. '
                        f'Просим направить мотивированный ответ на претензию по адресу:'
                        f' [указать почтовый адрес/адрес электронной почты] в срок [указать срок].')
    myDoc.add_paragraph(f'В случае полного или частичного отказа в удовлетворении претензии в [указать срок], '
                        f'я буду вынужден (будем вынуждены) обратиться с иском в суд за защитой своих законных'
                        f' прав и интересов. При этом мной/нами также будет заявлено требование о '
                        f'взыскании с Вас/Вашей организации судебных расходов по оплате государственной '
                        f'пошлины и судебных издержек, связанных с рассмотрением спора в суде.')
    myDoc.add_paragraph(f'Приложения:')
    myDoc.add_paragraph(f"1. \"{item}\".")
    # myDoc.add_paragraph(f'Расчет суммы возмещения ущерба.')
    myDoc.add_paragraph(f'{current_datetime.day}.{current_datetime.month}.{current_datetime.year}'
                        f'                            [должность]/________________/______________/')
    style = myDoc.styles.add_style('UserHead1', WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = 'Cambria'
    style.font.size = Pt(9)
    myDoc.add_paragraph(f'                                                                                                подпись      инициалы', style = style)

    myDoc.save(f"email/ClaimFile{linkIndex}.docx")
    sendEmail(f'ClaimFile{linkIndex}.docx')
    print(brandName)


for i in ulrList:
    createClaim(i)

'https://www.ozon.ru/api/entrypoint-api.bx/page/json/v2?url=%2Fproduct%2Famway-l-o-c-zhidkost-dlya-mytya-stekol-loc-amvey-sredstvo-dlya-stekla-lok-amvey-699409218%2F%3Flayout_container%3DpdpPage2column%26layout_page_index%3D2%26sh%3DioWT5K2JUg'
'https://www.ozon.ru/api/entrypoint-api.bx/page/json/v2?url=%2Fproduct%2Fpyatnovyvoditel-amvey-sprey-dlya-predvaritelnogo-vyvedeniya-pyaten-amway-sa8-prewash-spray-400-837491259%2F%3Flayout_container%3DpdpPage2column%26layout_page_index%3D2%26sh%3DioWT5OXGjg'