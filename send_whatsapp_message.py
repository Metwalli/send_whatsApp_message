from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl as excel

driver = webdriver.Chrome('./chromedriver')

file_name = "data.xlsx"
m = open("./message.txt", "r")
sending_message = m.read().splitlines()

# To convert python into exe file
# pyinstaller -w -F -i Whatsapp.ico send_whatsapp_message.py


driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 100)
wait5 = WebDriverWait(driver, 5)

admin_no = "+86 130 1893 8629" #input("Enter your Number:")
you_ = '//span[contains(@title,' + '"' + admin_no + '"' +')]'

file = excel.load_workbook(file_name)
sheet = file.active
for cell in range(1, len(sheet['B'])):
    if sheet['B'] != "":
        you_title = wait.until(EC.presence_of_element_located((
            By.XPATH, you_)))
        you_title.click()
        for c in range(ord('A'), ord('Z')+1):
            if sheet[chr(c)][0].value == 'mobile':
                send_link = "https://api.whatsapp.com/send?phone=" + sheet[chr(c)][cell].value + Keys.SHIFT + '\r\n' + Keys.SHIFT
                you_title = wait.until(EC.presence_of_element_located((By.XPATH, you_)))
                you_title.click()
                message = driver.find_element_by_class_name('_13mgZ')
                message.send_keys(send_link + Keys.ENTER)
                time.sleep(1)
                target_arg = '//a[contains(@href,' + '"' + "https://api.whatsapp.com/send?phone="\
                             + sheet[chr(c)][cell].value + '"' + ')]'
                break

        target_link = wait.until(EC.presence_of_element_located((
            By.XPATH, target_arg)))
        target_link.click()

        msg = Keys.SHIFT
        for line in sending_message:
            # find_field = False
            for c in range(ord('A'), ord('Z') + 1):
                variable = "$" + str(sheet[chr(c)][0].value) + "$"
                if line.find(variable) != -1:
                    line = line.replace(variable, str(sheet[chr(c)][cell].value))

            msg = msg + Keys.SHIFT + line + Keys.SHIFT + "\r\n"

        message = driver.find_element_by_class_name('_13mgZ')
        message.send_keys(msg)

        sendbutton = driver.find_element_by_class_name('_3M-N-')
        sendbutton.click()

        time.sleep(1)
