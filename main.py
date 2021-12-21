from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from pandas.io.excel import ExcelWriter
from webdriver_manager.chrome import ChromeDriverManager

s = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s)
driver.get("https://itdashboard.gov/")


def click_but():
    element = driver.find_element(By.XPATH, '//*[@id="node-23"]/div/div/div/div/div/div/div/a')
    element.click()
    time.sleep(5.5)


def find_agencies():
    click_but()
    make_list = list()
    for elem in driver.find_elements(By.XPATH, './/span[contains(@class, "h4 w200")]'):
        make_list.append(elem.text)
    del make_list[:27]
    return pd.DataFrame(make_list)


def go_to_agency():
    elem = driver.find_element(By.XPATH, '//*[@id="agency-tiles-widget"]/div/div[3]/div[2]/div/div/div/div[2]/a')
    elem.click()
    time.sleep(2)
    element = driver.find_element(By.XPATH, '//*[@id="block-itdb-custom--5"]/div/div/div/div[1]/div/h3')
    actions = ActionChains(driver)
    actions.move_to_element(element).perform()
    time.sleep(10)
    make = driver.find_element(By.XPATH, '//*[@id="investments-table-object_length"]/label/select')
    make.click()
    but = driver.find_element(By.XPATH, '//*[@id="investments-table-object_length"]/label/select/option[4]')
    but.click()
    time.sleep(30)


def find_investments():
    go_to_agency()
    make_list = list()
    for elem in driver.find_elements(By.XPATH, './/tr[contains(@role, "row")]'):
        make_list.append(elem.text)
    return pd.DataFrame(make_list)


def make_xlsx():
    first_list = find_agencies()
    second_list = find_investments()
    with ExcelWriter('teams.xlsx') as writer:
        first_list.sample(10).to_excel(writer, sheet_name="Agencies")
        second_list.sample(10).to_excel(writer, sheet_name="Individual Investments")


def download_pdf():
    if driver.find_element(By.XPATH, '//*[@id="investments-table-object"]/tbody/tr[1]/td[1]/a'):
        elem_first = driver.find_element(By.XPATH, '//*[@id="investments-table-object"]/tbody/tr[1]/td[1]/a')
        elem_first.click()
        time.sleep(5)
        elem_second = driver.find_element(By.XPATH, '//*[@id="business-case-pdf"]/a')
        elem_second.click()
    else:
        print("No link here")


def make_all():
    make_xlsx()
    download_pdf()


make_all()
