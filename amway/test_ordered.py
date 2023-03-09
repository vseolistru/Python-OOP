import re
from playwright.sync_api import Page, expect
import os
from dotenv import load_dotenv
load_dotenv()

user = os.getenv('GOS_USLUGI_USER')
password = os.getenv('GOS_USLUGI_PASSWORD')


def test_sendemail(page: Page):
    page.goto("https://zakaznoe.pochta.ru/mail/personal/compose")

    """login to pochta"""
    page.wait_for_load_state("load")
    page.locator("//input[@id='login']").fill(user)
    page.locator("//input[@id='password']").fill(password)
    page.locator("//button[@class = 'plain-button plain-button_wide']").click()

    """fill fields"""
    page.wait_for_timeout(8000)
    page.wait_for_load_state("load")
    page.locator("//form/div/div/div/div/div/div/div/div/div[1]/div/div/div/input").fill('ООО "ВАЙЛДБЕРРИЗ СК"')
    page.locator("//form/div/div/div/div/div/div/div/div/div[1]/div/div/div/input").click()
    page.locator("text = 121205, г Москва, Можайский р-н, тер Сколково инновационного центра, ул Нобеля, д 1, помещ 2 ком 14").click()
    # page.locator("input").fill('ООО "ВАЙЛДБЕРРИЗ СК"')

    current_data = os.getcwd()
    current_file = ''
    with open('output.txt', 'r') as data:
        current_file = data.read()

    print(f'{current_data}/pdf/{current_file}')
    file = os.path.join(f'{current_data}/pdf/{current_file}')
    # page.locator('text = Загрузить письмо').click()
    # upload = page.locator('text = Загрузить письмо').set_input_files(file)
    with page.expect_file_chooser() as file_chosser:
        page.locator('text = Загрузить письмо').click()

    fc = file_chosser.value
    fc.set_files(file)
    # page.wait_for_event('filechooser')
    page.wait_for_timeout(25000)

