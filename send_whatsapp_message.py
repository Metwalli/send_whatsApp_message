"""
Title: Send whatsApp message to list of numbers
Author: Metwalli Mohsen

"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl as excel
from tkinter import *
from tkinter import filedialog, messagebox
from urllib.request import urlopen
import datetime
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials


# Read the date from website
# import requests
# import bs4 as bs


class SendMessage():
    def __init__(self, am, dataDir, messageDir):
        self.dataDir = dataDir
        self.messagDir = messageDir
        self.admin_no = am
        self.success_counter = 0
        self.fail_counter = 0

    def license_check(self):
        res = urlopen('http://just-the-time.appspot.com/')
        result = res.read().strip()
        result_str = result.decode('utf-8')
        if datetime.datetime.strptime(result_str, '%Y-%m-%d %H:%M:%S') > datetime.datetime.strptime('2020-06-30 23:59:00', '%Y-%m-%d %H:%M:%S'):
            exit()

    def send_massage(self):
        try:

            # driver = webdriver.Chrome('chromedriver')
            # driver.get("https://web.whatsapp.com/")
            # wait = WebDriverWait(driver, 100)
            # javaScript = "var div = document.getElementsByClassName('z_tTQ')[0]; " \
            #              "var a = document.createElement('a'); a.href = 'https://api.whatsapp.com/send?phone=+8613018938629'; " \
            #              "a.text = 'https://api.whatsapp.com/send?phone=+8613018938629'; div.appendChild(a);"
            # driver.execute_script(javaScript)

            admin_no = self.admin_no
            you_ = '//span[contains(@title,' + '"' + admin_no + '"' + ')]'

            file_name = self.dataDir
            m = open(self.messagDir, "r")
            sending_message = m.read().splitlines()
            file = excel.load_workbook(file_name)
            sheet = file.active

            # To convert python into exe file
            # pyinstaller -w -F -i icon.ico send_whatsapp_message.py

            driver = webdriver.Chrome('ASWMDriver')
            driver.get("https://web.whatsapp.com/")
            wait = WebDriverWait(driver, 50)

        except:
            messagebox.showerror("Error", "Missing Dependencies")
            exit()
        noOfRecords = 1
        for cell in range(1, len(sheet['B'])):
            if str(sheet['A'][cell].value).strip() != "":
                noOfRecords += 1
            else:
                break

        for cell in range(1, noOfRecords):
            try:
                you_title = wait.until(EC.presence_of_element_located((
                    By.XPATH, you_)))
                you_title.click()
                for c in range(ord('A'), ord('Z')+1):
                    if sheet[chr(c)][0].value == 'mobile':
                        send_link = "https://wa.me/" + str(sheet[chr(c)][cell].value) + Keys.SHIFT + '\r\n' + Keys.SHIFT
                        you_title = wait.until(EC.presence_of_element_located((By.XPATH, you_)))
                        you_title.click()
                        message = driver.find_element_by_class_name('_2UL8j')

                        message.send_keys(send_link)

                        sendbutton = driver.find_element_by_class_name('_1U1xa')
                        sendbutton.click()
                        time.sleep(3)

                        target_arg = '//a[contains(@href,' + '"' + "https://wa.me/"\
                                     + str(sheet[chr(c)][cell].value) + '"' + ')]'
                        driver.create_web_element(target_arg)
                        break

                target_link = wait.until(EC.presence_of_element_located((
                    By.XPATH, target_arg)))
                target_link.click()
                time.sleep(3)
                try:
                    urlerror = driver.find_element_by_class_name('FV2Qy')
                    urlerror.click()
                    sheet["L" + str(cell + 1)] = "fail"
                    self.fail_counter += 1
                    continue
                except:
                    pass
                msg = Keys.SHIFT
                for line in sending_message:
                    # find_field = False
                    for c in range(ord('A'), ord('Z') + 1):
                        variable = "$" + str(sheet[chr(c)][0].value) + "$"
                        if line.find(variable) != -1:
                            line = line.replace(variable, str(sheet[chr(c)][cell].value))

                    msg = msg + Keys.SHIFT + line + Keys.SHIFT + "\r\n"

                message = driver.find_element_by_class_name('_2UL8j')
                message.send_keys(msg)

                sendbutton = driver.find_element_by_class_name('_1U1xa')
                sendbutton.click()
                sheet["L" + str(cell + 1)] = "success"
                self.success_counter += 1
                time.sleep(3)

                pass
            except:
                sheet["L" + str(cell + 1)] = "fail"
                self.fail_counter += 1
                time.sleep(3)
                # urlerror = driver.find_element_by_class_name('_1WZqU')
                # urlerror.click()
                pass
        try:
            file.save(self.dataDir)
            messagebox.showinfo("Info",
                                "Success Messages:" + str(self.success_counter) + "\n" + "Fail Messages:" + str(self.fail_counter) + "\n You can find the the sending results in data.xlsx file")
        except:
            file.save("output.xlsx")
            messagebox.showinfo("Info", "Success Messages:" + str(self.success_counter) + "\n" + "Fail Messages:" + str(self.fail_counter) + "\n You can find the the sending results in " + "output.xlsx file")


