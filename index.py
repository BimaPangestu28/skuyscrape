from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, ElementNotVisibleException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located, presence_of_all_elements_located
import time
import xlsxwriter
import platform
import requests

chromeOptions = webdriver.ChromeOptions()
chromeOptions.add_argument("--start-maximized")

if platform.system() == "Windows":
    driverSelect = './driver/chromedriver.exe'
else:
    driverSelect = './driver/chromedriver'

def get_universitas_data():
    driver = webdriver.Chrome(
        executable_path=driverSelect, chrome_options=chromeOptions)
    driver.get('https://forlap.ristekdikti.go.id/perguruantinggi')

    code = driver.find_element_by_xpath(
        '//*[@id="searchPtForm"]/div[6]')
    code = code.text.split(' ')
    code = int(code[1]) + int(code[3])

    driver.find_element_by_id("kode_pengaman").send_keys(code)

    driver.find_element_by_xpath(
        '//*[@id="searchPtForm"]/div[7]/div/input').click()

    page = 1
    total_dosen = 0
    total_mahasiswa = 0

    while page < 275:
        lists = driver.find_elements_by_class_name("ttop")

        for index, l in enumerate(lists):
            lists_university = open('lists_univesity.txt', 'a+')
            td = l.find_elements_by_xpath("td")

            total_dosen += int(td[6].text.replace('.', ''))
            total_mahasiswa += int(td[7].text.replace('.', ''))

            try:
                lists_university.write("{}, {}, {} \n".format(
                    td[2].text, td[6].text, td[7].text))
            except Exception as identifier:
                print(identifier)

            if index >= len(lists) - 1:
                time.sleep(2)

                try:
                    if page > 2:
                        driver.find_element_by_xpath(
                            '//*[@class="pagination"]/ul/li[4]/a').click()
                    elif page > 1:
                        driver.find_element_by_xpath(
                            '//*[@class="pagination"]/ul/li[3]/a').click()
                    else:
                        driver.find_element_by_xpath(
                            '//*[@class="pagination"]/ul/li[2]/a').click()
                except Exception as identifier:
                    pass

                page += 1

    lists_university.write("{}, {}, {} \n".format(
        "total", total_dosen, total_mahasiswa))

    lists_university.close()


def get_detail_row(tooltip):
    return '//div[@data-tooltip="{}"]/div[@class="section-info-line"]/span[@class="section-info-text"]/span[@class="widget-pane-link"]'.format(tooltip)


def get_data_from_map(skuychat_token_secret):
    keyword = input("Masukkan keyword untuk google maps : ")
    # keyword = "smk di jakarta"

    link = "https://www.google.com/maps/search/{}".format(keyword)

    driver = webdriver.Chrome(
        executable_path=driverSelect, chrome_options=chromeOptions)
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
                            xpath = '//*[@class="section-hero-header-title"]/div[@class="section-hero-header-title-description"]/div[1]/h1'

                            title = wait.until(
                                presence_of_element_located((By.XPATH, xpath))).text
                        except Exception as identifier:
                            print (identifier)

                        try:
                            address = wait.until(presence_of_element_located(
                                (By.XPATH, get_detail_row("Salin alamat")))).text
                        except Exception as identifier:
                            pass

                        try:
                            website = wait.until(presence_of_element_located(
                                (By.XPATH, get_detail_row("Buka situs")))).text
                        except Exception as identifier:
                            pass

                        try:
                            phone = wait.until(presence_of_element_located(
                                (By.XPATH, get_detail_row("Salin nomor telepon")))).text
                        except Exception as identifier:
                            pass

                        if skuychat_token_secret != "":
                            base_url = "http://localhost:8000/api/v1/hook/add-contact?token_secret={}".format(skuychat_token_secret)
                            # base_url = "https://skuychat.com/api/v1/hook/add-contact?token_secret={}".format(skuychat_token_secret)"

                            response = requests.post(base_url, data={
                                "name": title,
                                "phone": phone
                            })

                        try:
                            worksheet.write(row_cell, 0, title)
                            worksheet.write(row_cell, 1, phone)
                            worksheet.write(row_cell, 2, website)
                            worksheet.write(row_cell, 3, address)
                            # files.write("{}, {}, {}, {} \n".format(title, phone, website, address))

                            print("Data with name {} has been saved".format(title))
                        except Exception as identifier:
                            print (identifier)

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


def init():
    print(
        """
            Gogarmen Scrapping Data
            Silahkan jenis data yang ingin kamu dapatkan

            1) Data perguruan tinggi di indonesia
            2) Data scrapping dari google maps
        """
    )

    select_option()


def select_option():
    option = input("Pilih jenis data : ")
    # option = 2

    print("""
        Skuyscrape adalah fitur tambahan dari skuychat
        Apakah kamu ingin langsung menambahkan kontak ke skuychat milik mu?
    """)

    option_integrate = input("Pilih (Y/N) : ")

    skuychat_token_secret = ""

    if option_integrate.lower() == "y":
        skuychat_token_secret = input("Masukkan token secret skuychat milik mu : ")

    if option == "1":
        get_universitas_data()
    elif option == "2":
        get_data_from_map(skuychat_token_secret)
    else:
        print("Option tidak ditemukan")
        select_option()


if __name__ == "__main__":
    # time.sleep(1)

    init()
