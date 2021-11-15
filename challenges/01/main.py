from selenium.webdriver import Chrome
from selenium.webdriver.support.ui import Select
import openpyxl
import time


class Bot01:
    def __init__(self):
        self.driver = Chrome(executable_path=r'../chromedriver.exe')
        self.driver.get('https://desafiosrpa.com.br/inputforms.html?')
        self.driver.maximize_window()

    def getData(self):
        # Read file
        book = openpyxl.load_workbook("data.xlsx")

        # Set page
        dataPage = book["Dados"]

        # Datas
        for rows in dataPage.iter_rows(min_row=2):
            result = [cell.value for cell in rows]
            print(result)
            self.fillForm(*result)

        return self.driver.quit()

    def fillForm(self, *args):
        try:
            # Insert ID
            self.driver.find_element_by_xpath("//input[@placeholder='ID0000']").send_keys(args[0])

            # Insert First Name
            self.driver.find_element_by_xpath("//input[@placeholder='Joaquim']").send_keys(args[1])

            # Insert Last Name
            self.driver.find_element_by_xpath("//input[@placeholder='Xavier']").send_keys(args[2])

            # Selected Country
            country = self.driver.find_element_by_tag_name('select')
            options = Select(country)
            options.select_by_visible_text(args[3])

            # Check options of input of type select (OPTIONAL)
            # countryOptions = country.find_elements_by_tag_name("option")
            # text = [option.text for option in countryOptions]
            # print(text)

            # Inset Date Birth
            birthDiv = self.driver.find_element_by_id('divNascimento')
            birthDiv.find_element_by_tag_name('input').send_keys(args[4].strftime('%d%m%Y'))

            # Insert Date Credits
            creditsDiv = self.driver.find_element_by_id('divDataCredito')
            creditsDiv.find_element_by_tag_name('input').send_keys(args[5].strftime('%d%m%Y'))

            time.sleep(1)
            self.driver.find_element_by_id('btao').submit()
        except Exception as e:
            print(e)
            self.driver.quit()


myBot = Bot01()
myBot.getData()
