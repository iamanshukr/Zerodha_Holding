from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time

workbook = load_workbook("D:\\Python\\Projects\Zerodha\\UserData.xlsm")
sheet = workbook.active

cell = sheet["A2"].value
# print(cell)
i = 2
while i < sheet["F1"].value:
    a = "A"+str(i)
    b = "B"+str(i)
    c = "C"+str(i)
    username = sheet[a].value
    password = sheet[b].value
    PIN = sheet[c].value

    i = i+1

    driver = webdriver.Chrome("C:\Program Files\Chrome driver\\chromedriver.exe")
    driver.get("https://kite.zerodha.com/")

    username_textbox = driver.find_element_by_id("userid")
    username_textbox.send_keys(username)

    password_textbox = driver.find_element_by_id("password")
    password_textbox.send_keys(password)
    # time.sleep(3)
    password_textbox.send_keys(Keys.ENTER)

    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "pin"))
        )
        element.send_keys(PIN)

        element.click()
        element.send_keys(Keys.ENTER)
        
        # element = WebDriverWait(driver, 10).until(
        #     EC.presence_of_element_located((By.LINK_TEXT, "holdings"))
        # )
        # element.click()
        
    except:
        pass


    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.get("https://console.zerodha.com/portfolio/holdings")

    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.LINK_TEXT, "XLSX"))
        )
        element.click()

        time.sleep(15)

        # element = WebDriverWait(driver, 8).until(
        #     EC.presence_of_element_located((By.XPATH,"//a[@href = '/logout']"))
        # )
        # element.click()

    except:
        pass


    #This program downloads the list of stocks holding in Zerodha Demat account

print(f"Your holding statement of {i-2} accounts has been downloaded")