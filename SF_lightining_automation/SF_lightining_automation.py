#@By Mahmoud sayid

#import modules    
def page_is_loading(driver):
    while True:
        x = driver.execute_script("return document.readyState")
        if x == "complete" :
            return True
        else:
            return False


import os

import time
Main_path=os.path.expanduser('~\\MTools')

def imports(package):
    try:
        __import__(package)
        print(package+ " imported successfully ")
    except ImportError:
        print("installing "+package+"....")
        os.system("pip install "+ package+" --user")   
    
#global browser
imports("selenium")
imports("pywin32")

import tkinter
from selenium import webdriver
from selenium.webdriver.support.ui import Select
# Import smtplib for the actual sending function
import smtplib
import win32com.client as win32
# Import the email modules we'll need
from email.message import EmailMessage
from tkinter import *
from tkinter import messagebox
from tkinter import simpledialog
from tkinter.filedialog import askopenfile 
from selenium.common.exceptions import NoSuchElementException     
from pathlib import Path
try:
    import urllib.request as url
except ImportError:
    import urllib as url
    print("Downloading compoments....")


def download_images():
# In order to fetch the image online  
    Path(Main_path).mkdir(parents=True, exist_ok=True)
    if (Path(Main_path+"\\Mellanox.png")).exists():
        print ("File exist")
    else:
        print ("File not exist downloading file...")
        url.urlretrieve("https://upload.wikimedia.org/wikipedia/he/3/38/Nvidia_logo.png", Main_path+"\\Mellanox.png")
    
    
    if (Path(Main_path+"\\chromedriver.exe")).exists():
        print ("File exist")
    else:
        print ("File not exist downloading Driver...")
        url.urlretrieve("https://github.com/mhmodsayid/Mellanox_Tool/raw/master/chromedriver.exe",Main_path+"\\chromedriver.exe")
        print("Downloading completed.")


def chrome_driver_lunch(mode=None):
    global main_url
    global browser
    main_url = "https://mellanox.lightning.force.com/lightning/r/Case/5001T00001QnhArQAJ/view"
    
    options = webdriver.ChromeOptions()
    options.arguments.append("--no-sandbox")
    options.arguments.append("--log-level=3")
    options.arguments.append("--disable-dev-shm-usage")
    #options.arguments.append("--window-position=-32000,-32000")# to hide the browser

    try:
        options.arguments.append("--user-data-dir="+Main_path+"\\selenium") 
        browser = webdriver.Chrome(executable_path=Main_path+"\\chromedriver.exe", chrome_options=options)
        browser.get(main_url)
        while not page_is_loading(browser):
                continue
        browser.implicitly_wait(10)

    except Exception as identifier:
        print(str(identifier))
 
    
# Designing window for registration
 
def chrome_possion(mode=None):
    if mode=="hide":
        browser.set_window_position(-32000,-32000)
    if mode=="show":
        browser.set_window_position(0,0)
     

download_images()

# Designing window for login 
def close():
    register_screen.destroy()


def login():
    chrome_driver_lunch()
    case_number="00908203"

    browser.find_element_by_xpath('//div[@class="uiInput uiAutocomplete uiInput--default"] //input[@title="Search..."]').click()
    browser.find_element_by_xpath('//div[@class="uiInput uiAutocomplete uiInput--default"] //input[@title="Search..."]').send_keys(case_number)
    browser.find_element_by_xpath('//mark[contains(text(),'+case_number+')]/..').click()

    while not page_is_loading(browser):
        continue
    time.sleep(1)
    Case_id=browser.find_element_by_xpath('//div[@class="windowViewMode-maximized active lafPageHost"] //lightning-formatted-text[contains(text(),"ref")]').text
    initializeOutlook()
    mail.Subject = Case_id
    mail.Body = "Hello World, Test"
    mail.Send()
    print(Case_id)

    #RMA dequeue
    
    browser.find_element_by_xpath('//div[@class="windowViewMode-maximized active lafPageHost"] //button[@title="Change Owner"]').click()
    browser.find_element_by_xpath('//a[@aria-label="Enter new owner name—Current Selection: , Pick an object"]').click()
    browser.find_element_by_xpath('//*[@title="Queues"]').click()
    browser.find_element_by_xpath('//div[@class="changeOwnerContentBody forceRecordLayout"]//input[@title="Search Queues"]').send_keys("Mellanox Technical Support")
    browser.find_element_by_xpath('//Ul[@class="lookup__list  visible"]//div[@title="Mellanox Technical Support"]').click()
    browser.find_element_by_xpath('//div[@class="modal-footer slds-modal__footer"]//button[contains(text(),"Change Owner")]').click()



def initializeOutlook(Mail=None):
    
    global mail
    global outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'networking-support@nvidia.com'
    if(Mail is not None):
        mail.To=Mail

    for account in outlook.Session.Accounts:
        if("@nvidia.com" in account.DisplayName):
            username=account.DisplayName
            print(account.DisplayName)

    return  StringVar(register_screen, value=username)


# Designing Main(first) window
 
def main_account_screen():
    try:
        global register_screen
        register_screen = Tk()
        register_screen.title("Mellanox AutoAgent")
        register_screen.geometry("420x500")
        global username
        global password
        global username_entry
        global password_entry
        username = StringVar(register_screen, value='')
        password = StringVar(register_screen, value='')
        username=initializeOutlook()
   
    
        Label(register_screen, text="").pack()
        username_lable = Label(register_screen, text="Email:")
        username_lable.pack()
        username_entry = Entry(register_screen, textvariable=username)
        username_entry.pack()
        password_lable = Label(register_screen, text="\nPassword: ")
        password_lable.pack()
        password_entry = Entry(register_screen, textvariable=password,show='*')
        password_entry.pack()
        Label(register_screen, text="").pack()
        Button(text="Login", height="1", width="10", command = login).pack()
        widget = Label(register_screen, compound='top')
        widget.lenna_image_png=PhotoImage(file=Main_path+"\\Mellanox.png")
    
        widget['text'] = "Ver: 1.00"
        widget['image'] = widget.lenna_image_png
        widget.pack()
        register_screen.mainloop()
    finally:
        browser.close()
        browser.quit()
       
main_account_screen()

