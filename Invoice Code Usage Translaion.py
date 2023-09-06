import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from element_manager import *
from selenium.webdriver.common.keys import Keys
import pandas as pd
from tkinter import *
from tkinter import filedialog, ttk
from tkinter.filedialog import askopenfile
import os

filepath = ""
username_var = None
password_var = None
root = Tk()


def files():
    global filepath

    file = filedialog.askopenfile(mode='r', filetypes=[('EXCEL', '*.xlsx')])
    if file:
        filepath = os.path.abspath(file.name)
        Label(root, text="The File is located at : " + str(filepath), font=('Aerial 11')).pack()
        return filepath


ttk.Button(root, text="Select file", command=files).pack(pady=20)


def start():
    global filepath
    global username_var
    global password_var
    if filepath:
        username = username_var.get()
        password = password_var.get()

        if not username or not password:
            print("Please enter both username and password.")
            return
        driver = webdriver.Chrome()
        # to open the url in browser
        driver.get('https://rhenus.okta.com/')
        excel_file = filepath
        data = pd.read_excel(excel_file)
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(
            driver, '126YJaa68GerZcM')).send_keys(username)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'f1GMPJOh2_4aGTs')).click()

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(
            driver, 'dVdU1X5TlD1mXYy')).send_keys(password)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'GtyGHu4AyVcmHk0')).click()

        # to click on the element found
        driver.find_element(By.XPATH, get_xpath(driver, 'ESHkRyzqZ41400A')).click()

        time.sleep(15)
        # to switch to another window
        driver.get('https://rocs2-americas.rocs.live/tmsclient/')
        time.sleep(5)
        # to click on the element found
        driver.find_element(By.XPATH, get_xpath(driver, 'mwAVIwjRPQgwEJ4')).click()
        time.sleep(1)
        # to click on the element(Invoice code usage) found
        driver.find_element(By.XPATH, get_xpath(driver, 'WYnSHOxxIGwDBrB')).click()
        time.sleep(1)
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(
            driver, 'UOI81OkmSb3dClU')).send_keys('99')
        time.sleep(1)
        # press Enter key
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        # to click on the element(Seafreight Managemen...) found
        driver.find_element(
            By.XPATH, get_xpath(driver, 'TUM_pTlvV22oDpx')).click()
        time.sleep(1)

        for index, row in data.iterrows():
            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[3]/div/div[2]/div/div[3]/div[1]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div/div[5]/input').clear()
            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[3]/div/div[2]/div/div[3]/div[1]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div/div[5]/input').send_keys(
                row['Invoicecode'])
            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[3]/div/div[2]/div/div[3]/div[1]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div/div[5]/input').send_keys(
                Keys.ENTER)
            # to click on the element() found
            time.sleep(1)
            driver.find_element(By.XPATH, get_xpath(driver, 'grUTdmI4fTg5idY')).click()
            time.sleep(1)
            # to type content in input field
            driver.find_element(By.XPATH, get_xpath(driver, 'BGeYF3LxwArp8Ks')).send_keys(row['LanguageCode'])
            time.sleep(1)
            # to click on the element found
            driver.find_element(By.XPATH, get_xpath(driver, 't0n4BhEZb_Gpxi6')).click()
            time.sleep(1)
            # to click on the element found
            driver.find_element(By.XPATH, get_xpath(driver, 'X2ypjED3Wn67uKu')).click()
            time.sleep(1)
            # to type content in input field
            driver.find_element(By.XPATH, get_xpath(driver, 'DxG94fNNP1VsAvl')).send_keys(row['TranslationDescription'])
            driver.find_element(By.XPATH, get_xpath(driver, 'DxG94fNNP1VsAvl')).send_keys(Keys.ENTER)
            time.sleep(1)
            driver.find_element(By.XPATH,
                                '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div[3]/div[1]/div[3]/div[1]/div/div[1]').click()
            time.sleep(1)
    else:
        print("No file selected.")


root.title("Invoice code usage Translation")
root.geometry("600x600")
username_label = Label(root, text="Username:")
username_label.pack()
username_var = StringVar()
username_entry = Entry(root, textvariable=username_var)
username_entry.pack()

password_label = Label(root, text="Password:")
password_label.pack()
password_var = StringVar()
password_entry = Entry(root, textvariable=password_var, show="*")  # Show * for password input
password_entry.pack()

ttk.Button(root, text="Start Automation", command=start).pack(pady=20)
root.mainloop()
