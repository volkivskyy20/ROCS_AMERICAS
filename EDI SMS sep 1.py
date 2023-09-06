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
    driver.get('https://rhenus.okta.com/')

    excel_file = filepath
    data = pd.read_excel(excel_file)
    # to type content in input field
    driver.find_element(By.XPATH, get_xpath(driver, '126YJaa68GerZcM')).send_keys(username)

    # to click on input field
    driver.find_element(By.XPATH, get_xpath(driver, 'f1GMPJOh2_4aGTs')).click()

    # to type content in input field
    driver.find_element(By.XPATH, get_xpath(driver, 'dVdU1X5TlD1mXYy')).send_keys(password)

    # to click on input field
    driver.find_element(By.XPATH, get_xpath(driver, 'GtyGHu4AyVcmHk0')).click()

    # to click on the element found
    driver.find_element(By.XPATH, get_xpath(driver, 'ESHkRyzqZ41400A')).click()

    time.sleep(15)
    # to switch to another window
    driver.get('https://rocs2-americas.rocs.live/tmsclient/')
    time.sleep(5)
    driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/div').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div[2]/div[3]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, get_xpath(driver, 'LQYZGc2I2aI22aP')).click()
    time.sleep(1)
    driver.find_element(By.XPATH, get_xpath(driver, '6AV6hVACqgDOqL3')).click()
    time.sleep(1)
    for index, row in data.iterrows():
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div/div[5]/input').click()
        time.sleep(2)
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div/div[5]/input').send_keys(
            row['EDIclient'])
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div/div[5]/input').send_keys(
            Keys.ENTER)

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div[3]/div[1]/input').send_keys(row['Checkreceivercode'])

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div[3]/div[2]/input').clear()

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div[3]/div[2]/input').send_keys(row['Useownmatchtables'])

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div[3]/div[3]/input').clear()

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div[3]/div[3]/input').send_keys(row['Usepartyidentification'])

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div[3]/div[4]/input').clear()

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/div[3]/div[4]/input').send_keys(row['Addtolastopenconsolidation'])

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/input').clear()

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/div[3]/div[1]/input').send_keys(row['Skipduplicates'])

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/input').clear()

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/div[3]/div[2]/input').send_keys(row['Autostatustransfer'])
        # time.sleep(1)

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[1]/div[3]/div[2]/input').send_keys(row['Matchtablesof'])

        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[1]/div[3]/div[4]/input').clear()
        time.sleep(1)
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[1]/div[3]/div[4]/input').send_keys(
            row['Relationcode'])
        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[1]/div[3]/div[4]/input').send_keys(Keys.ENTER)
        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[1]/div[3]/div[7]/input').send_keys(row['Officecode'])
        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[1]/div[3]/div[7]/input').send_keys(Keys.ENTER)
        # driver.find_element(By.XPATH,'/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[3]/div[3]/div[2]/div[3]/div[1]/div[3]/div[9]/input').send_keys(row['OfficeEDIcode'])

        # driver.find_element(By.XPATH,get_xpath(driver,'7LNKiOOfMjA9nfd')).click()
        # time.sleep(1)
        # driver.find_element(By.XPATH,get_xpath(driver,'CeQQzEaMmUxqlyI')).send_keys(row["EDIclient"])
        # time.sleep(1)
        # driver.find_element(By.XPATH,get_xpath(driver,'8x5dFCVA4NBfGdt')).click()
        # time.sleep(1)
        # driver.find_element(By.XPATH,get_xpath(driver,'x3NXbHXzjLJX4MC')).send_keys(row["EDIcode"])
        # time.sleep(1)
        # driver.find_element(By.XPATH,get_xpath(driver,'jmsFl5X7NOzATYN')).click()
        # time.sleep(1)
        # driver.find_element(By.XPATH,get_xpath(driver,'6K4NXYj88R3uG15')).send_keys(row["Targetappl."])
        # time.sleep(1)

        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[3]/div/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div[1]/button').click()
        time.sleep(2)


root.title("EDI SMS Step 1")
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
