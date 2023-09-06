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
        time.sleep(1)
        driver.find_element(By.XPATH, get_xpath(driver, 'o_Eoif0BDwJnyLr')).click()
        time.sleep(1)
        driver.find_element(By.XPATH, get_xpath(driver, 'oZp_YI0ZuZKAyI5')).clear()
        time.sleep(1)
        driver.find_element(By.XPATH, get_xpath(driver, 'oZp_YI0ZuZKAyI5')).send_keys(row['EDIClient'])
        time.sleep(1)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        driver.find_element(By.XPATH, get_xpath(driver, 'Wzo_UnrNwzIgse1')).click()
        time.sleep(1)
        shipper_value = int(row['Shipper'])
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[6]/div[6]/div[2]/div[2]/div[3]/div[4]/div[2]/div/div[3]/div/div[3]/div[2]/input').send_keys(
            str(shipper_value))
        time.sleep(1)

        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[6]/div[6]/div[2]/div[2]/div[3]/div[4]/div[2]/div/div[3]/div/div[3]/div[2]/input').send_keys(
            Keys.ENTER)
        time.sleep(1)
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[6]/div[6]/div[2]/div[2]/div[3]/div[4]/div[2]/div/div[3]/div/div[3]/div[4]/input').send_keys(
            row['ShipperQualifier'])
        time.sleep(1)
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[6]/div[6]/div[2]/div[2]/div[3]/div[4]/div[2]/div/div[3]/div/div[3]/div[4]/input').send_keys(
            Keys.ENTER)
        time.sleep(1)

        driver.find_element(By.XPATH, get_xpath(driver, 'IfaB8B1XdXVuAdB')).send_keys(row['RefQualifier'])
        time.sleep(1)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        # to type content in input field
        dep = int(row['Subdept'])

        driver.find_element(By.XPATH, get_xpath(driver, 'h_u4z6bwZwT6mYm')).send_keys(str(dep))
        time.sleep(1)
        # press Enter key
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        # to click on the element(OCEAN IMPORT LCL) found

        driver.find_element(By.XPATH, get_xpath(driver, 'imWrjk5Om3fxbpa')).click()
        time.sleep(1)
        # to click on the element(OK) found
        driver.find_element(By.XPATH, get_xpath(driver, 'JN4uc22vsfzLeUv')).click()
        time.sleep(1)
        # to type content in input field
        pol_quali = int(row['POLQualifier'])
        driver.find_element(By.XPATH, get_xpath(driver, 'ZQ0PAKGknSXj93r')).send_keys(str(pol_quali))
        time.sleep(1)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        pod_quali = int(row['PODQualifier'])
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'yujjcBVmpSNnzUi')).send_keys(str(pod_quali))
        time.sleep(1)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        etd_quali = int(row['ETDQualifier'])
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'Si5J6ktoDe5IqFX')).send_keys(str(etd_quali))
        time.sleep(1)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        eta_quali = int(row['ETAQualifier'])
        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'hViLupjH9UIj0yv')).send_keys(eta_quali)
        time.sleep(1)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        # to click on input field
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div[6]/div[6]/div[2]/div[2]/div[3]/div[4]/div[1]/div[3]').click()
        time.sleep(2)
        # to type content in input field
        con = int(row['Consignee'])
        driver.find_element(By.XPATH, get_xpath(driver, 'mpY7JQC4xdgnRxi')).send_keys(con)
        time.sleep(1)

        # press Enter key
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'DKfRdxhikceLFli')).click()
        time.sleep(1)

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'NwHD0uW6iSDEnkI')).send_keys(row['ConsigneeQualifier'])
        time.sleep(1)

        # press Enter key
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)

        # to click on input field
        driver.find_element(By.XPATH, get_xpath(driver, 'hUMuzad96gWkGZ9')).click()
        time.sleep(1)

        # to type content in input field
        driver.find_element(By.XPATH, get_xpath(driver, 'MpOCvT9B7Ev_tSa')).send_keys(row['RefQualifier2'])
        time.sleep(1)

        # press Enter key
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div[6]/div[3]/div[1]/div[2]').click()
        time.sleep(5)


root.title("EDI SMS Step 2")
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
