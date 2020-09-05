from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, ElementNotVisibleException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located, presence_of_all_elements_located
from webdriver_manager.chrome import ChromeDriverManager
import time
import xlsxwriter
import platform
import requests

def get_detail_row(tooltip):
    return '//div[@data-tooltip="{}"]/div[@class="section-info-line"]/span[@class="section-info-text"]/span[@class="widget-pane-link"]'.format(tooltip)


def init():
    # keyword = input("Masukkan keyword untuk google maps : ")
    keyword = "smk di jakarta"

    link = "https://www.google.com/maps/search/{}".format(keyword)

    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument("--start-maximized")
    chromeOptions.add_experimental_option(
            "prefs", {"intl.accept_languages": "en,en_US"})

    driver = webdriver.Chrome(
        executable_path=ChromeDriverManager().install(), chrome_options=chromeOptions)
    wait = WebDriverWait(driver, 10)
    driver.get(link)

    # files = open('results/{}.csv'.format(keyword), 'a+')
    workbook = xlsxwriter.Workbook('results/{}.xlsx'.format(keyword))
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.write(0, 0, 'Name', bold)
    worksheet.write(0, 1, 'Phone', bold)
    worksheet.write(0, 2, 'Website', bold)
    worksheet.write(0, 3, 'Address', bold)
    row_cell = 1

    while True:
        try:
            time.sleep(2)
            index = 0
            rows = wait.until(presence_of_all_elements_located(
                (By.XPATH, '//*[@class="section-result"]')))

            for i in range(len(rows)):
                try:
                    rows = wait.until(presence_of_all_elements_located(
                        (By.XPATH, '//*[@class="section-result"]')))

                    for idx, row in enumerate(rows):
                        if idx != index:
                            continue

                        row.click()
                        title = ""
                        phone = ""
                        address = ""
                        website = ""

                        try:
                            xpath = '//*[@id="pane"]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]'

                            title = wait.until(
                                presence_of_element_located((By.XPATH, xpath))).text
                        except Exception as identifier:
                            pass

                        try:
                            xpath = '//*[@id="pane"]/div/div[1]/div/div/div[8]/button/div/div[2]/div[1]'

                            address = wait.until(
                                presence_of_element_located((By.XPATH, xpath))).text
                        except Exception as identifier:
                            pass

                        try:
                            xpath = '//*[@id="pane"]/div/div[1]/div/div/div[10]/button/div/div[2]/div[1]'

                            website = wait.until(
                                presence_of_element_located((By.XPATH, xpath))).text
                        except Exception as identifier:
                            pass

                        try:
                            xpath = '//*[@id="pane"]/div/div[1]/div/div/div[11]/button/div/div[2]/div[1]'

                            phone = wait.until(
                                presence_of_element_located((By.XPATH, xpath))).text
                        except Exception as identifier:
                            pass

                        print("nama", title)
                        print("telpon", phone)
                        print("website", website)
                        print("alamat", address)

                        base_url = "http://contact.maxcotton.id/api/v1/contacts"

                        print(base_url)

                        response = requests.post(base_url, data={
                            "name": title if title != "" else "-",
                            "phone": phone if phone != "" else "-",
                            "email": "-",
                            "website": website if website != "" else "-",
                            "address": address if address != "" else "-",
                            "tag": keyword
                        })
                        print(response)
                        print(response.json())

                        worksheet.write(row_cell, 0, title)
                        worksheet.write(row_cell, 1, phone)
                        worksheet.write(row_cell, 2, website)
                        worksheet.write(row_cell, 3, address)

                        print("Data with name {} has been saved".format(title))

                        row_cell += 1
                        index += 1

                        wait.until(presence_of_element_located(
                            (By.CLASS_NAME, "section-back-to-list-button"))).click()

                        time.sleep(2)
                except Exception as identifier:
                    # print (identifier)
                    pass

            wait.until(presence_of_element_located(
                (By.XPATH, '//*[@aria-label="Halaman berikutnya"]'))).click()
        except Exception as identifier:
            print(identifier)
            break

    workbook.close()

if __name__ == "__main__":
    # time.sleep(1)

    init()
