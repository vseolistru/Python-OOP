from playwright.sync_api import Page
import os


def test_sendClaimBySiteForm(page: Page):
    page.goto('https://seller.wildberries.ru/appeal-copyright/')

    """fill filelds"""
    page.wait_for_load_state('load')
    page.locator("//input[@class = 'e06176']").click()
    page.locator("//li[@class='c3a9e9'][2]").click()
    page.locator("//input[@id='theme_id']").click()
    page.locator("//li[@class='c3a9e9'][3]").click()

    """fill good`s art"""
    page.wait_for_timeout(4000)
    with open('output.txt', 'r') as data:
        current_file = data.read()
    myName = current_file[9:-4]
    page.locator("//input[@id='articles.0']").fill(myName)

    """fill checkBox"""
    page.locator("//div[@class = 'a5015e'][2]/div/label[@class ='c38474 cb8e91 b07ff2']").click()

    """fill Contacts"""
    page.locator("//input[@id='cro_company']").fill('Правообладатель - Товарищество с ограниченной '
                  'ответственностью "Эмвэй Казахстан"')
    page.locator("//input[@id='cro_brand']").fill('Amway')
    page.locator("//input[@name='cro_phone']").fill('74950232022')
    page.locator("//input[@id='cro_fio']").fill('Баутин Михаил Владимирович')
    page.locator("//input[@id='cro_email']").fill('some@mail.com')

    """files upload"""
    current_data = os.getcwd()
    file = os.path.join(f'{current_data}/pdf/{current_file}')
    """first file"""
    with page.expect_file_chooser() as file_chooser:
        page.locator("//div[@class='dec352 bd5d1c']/label[text()='Подтверждение вашего права']").click()
    fc = file_chooser.value
    fc.set_files(file)

    """second file"""
    with page.expect_file_chooser() as second_file_chooser:
        page.locator("//div[@class='dec352 bd5d1c']/label[text()='Подтверждение нарушения права']").click()
    fc = second_file_chooser.value
    fc.set_files(file)

    """thirth file"""
    with page.expect_file_chooser() as next_file_chooser:
        page.locator("//div[@class='dec352 bd5d1c']/label[text()='Текст жалобы']").click()
    fc = next_file_chooser.value
    fc.set_files(file)

    """last file"""
    with page.expect_file_chooser() as last_file_chooser:
        page.locator("//div[@class='dec352 bd5d1c']/label[text()='Полномочия']").click()
    fc = last_file_chooser.value
    fc.set_files(file)

    # page.wait_for_timeout(14000)