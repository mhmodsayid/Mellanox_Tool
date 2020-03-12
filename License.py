from selenium import webdriver
from selenium.webdriver.support.ui import Select
# Import smtplib for the actual sending function
import smtplib

# Import the email modules we'll need
from email.message import EmailMessage

from tkinter import *

#import modules
import os
options = webdriver.ChromeOptions()
"""options.arguments.append("headless")"""
browser = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=options)
# Designing window for registration
 

# Designing window for login 
 
def login():
    

    global login_screen
    login_screen = Toplevel(register_screen)
    login_screen.title("Login")
    login_screen.geometry("300x250")
    Label(login_screen, text="Please enter SO number").pack()
    Label(login_screen, text="").pack()
    global username_verify
    username_verify = StringVar()
    global username_login_entry
    Label(login_screen, text="SO number").pack()
    username_login_entry = Entry(login_screen, textvariable=username_verify)
    username_login_entry.pack()
    Label(login_screen, text="").pack()

    
    Label(login_screen, text="").pack()
    Button(login_screen, text="Dequeue", width=10, height=1, command = login_verify).pack()
 
# Implementing event on register button
 

 
# Implementing event on login button 
 
def login_verify():
    so_number = username_verify.get()
   


    # Define the URL's we will open and a few other variables 
    main_url = 'https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?str=00751864+&searchType=2&isdtp=vw&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce#!/fen=a4F&initialViewMode=detail&str=1086773'


    # Open main window with URL A
    
   
    browser.get(main_url)

    browser.implicitly_wait(5)
    browser.find_element_by_id("i0116").send_keys("mahmouds@mellanox.com")
    browser.find_element_by_id("idSIButton9").click()
    browser.implicitly_wait(30)

    """k = input("Please enter your case\n")"""
    k=so_number
    while k!=0:
        
    
      #  browser.find_element_by_id("secondSearchText").clear()
       # browser.find_element_by_id("secondSearchText").send_keys(k)
        #browser.find_element_by_id("secondSearchButton").click()
        #browser.find_element_by_link_text(k).click()
        """
        browser.find_element_by_link_text('[Change]').click()
        select = Select(browser.find_element_by_id("newOwn_mlktp"))
        select.select_by_index(1)
        browser.find_element_by_id("newOwn").send_keys("Mellanox Technical Support")
        browser.find_element_by_xpath('//*[@title="Save"]').click()"""

        browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[2]/td[2]/div/a').click()


                
        Data = browser.find_elements_by_class_name('dataCell')
        output=""
        
        

        index=[0,1,2,3,4,8]
        
        for i in range(int(assets_Size_int)):
            for j in index:
                output+=Data[j+(10*i)].text
                output+="\n"
            output+="------------------------- \n"   

        browser.execute_script("window.history.go(-1)")
        
    
        
        print("Current Page Title is : %s" %browser.title)
        Case_ID=browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[1]/td[4]/div').text

        
        print(output)
       


# Designing Main(first) window
 
def main_account_screen():
    global register_screen
    register_screen = Tk()
    register_screen.title("Register")
    register_screen.geometry("300x250")
 
    global username
    
    global username_entry
    
    username = StringVar()
 
    Label(register_screen, text="").pack()
    username_lable = Label(register_screen, text="Email")
    username_lable.pack()
    username_entry = Entry(register_screen, textvariable=username)
    username_entry.pack()
    
    Label(register_screen, text="").pack()
    Button(text="Login", height="2", width="30", command = login).pack()
    
 
    register_screen.mainloop()
 
 
main_account_screen()







k = input("Please enter your case\n")
browser.stop_client()
browser.close()
browser.quit()