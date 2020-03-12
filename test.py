class Asset:
    def __init__(self, serial_number,part_number,warranty,dOA,pD,aD):
        self.serial_number = serial_number
        self.part_number = part_number
        self.warranty = warranty
        self.dOA = dOA
        self.pD = pD
        self.aD = aD

#
    def __str__(self):
        return "Serial Number:  {0}\nPart Number:  {1}\nWarranty:  {2}\nDOA:  {3}\nproblem description:\n{4}\nPart description:\n{5}\n===========================================================\n".format(self.serial_number, self.part_number, self.warranty, self.dOA, self.pD,self.aD)
    

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
options.arguments.append("headless")#hide the browser
browser = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=options)
main_url = 'https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com'
# Designing window for registration
 

# Designing window for login 
 
def login():
   

    # Send the message via our own SMTP server.
    
    

# Define the URL's we will open and a few other variables 
    


    # Open main window with URL A
    
   
    
    browser.get(main_url)
    browser.implicitly_wait(5)
    browser.find_element_by_id("i0116").send_keys("mahmouds@mellanox.com")
    browser.find_element_by_id("idSIButton9").click()
    browser.implicitly_wait(30)
    browser.find_element_by_id("secondSearchText").clear()


    global login_screen
    login_screen = Toplevel(register_screen)
    login_screen.title("Mellanox AutoAgent")
    login_screen.geometry("450x450")
    Label(login_screen, text="Please enter RMA Details bellow:").pack()

    global username_verify
    username_verify = StringVar()
    global username_login_entry
    Label(login_screen, text="RMA case number:").pack()
    username_login_entry = Entry(login_screen, textvariable=username_verify)
    username_login_entry.pack()
    
    widget = Label(login_screen, compound='top')
    widget.lenna_image_png=PhotoImage(file="Mellanox.png")
    widget['image'] = widget.lenna_image_png
   
   

    Button(login_screen, text="Dequeue, Add summary comment", width=30, height=1, command = login_verify).pack()
    widget.pack()
    register_screen.withdraw()
# Implementing event on register button
 

 
# Implementing event on login button 
 
def login_verify():
    
    rma_Number = username_verify.get()
    browser.get(main_url)
    msg = EmailMessage()
    msg['From'] = username.get()
    msg['To'] = 'supportadmin@mellanox.com'
    server = smtplib.SMTP('smtp.office365.com.', 587)
    server.ehlo()
    server.starttls()
    #Next, log in to the server
    server.login(username.get(), password.get())
    

    """k = input("Please enter your case\n")"""
    k=rma_Number
    
        
    
    browser.find_element_by_id("secondSearchText").clear()
    browser.find_element_by_id("secondSearchText").send_keys(k)
    browser.find_element_by_id("secondSearchButton").click()
    browser.find_element_by_link_text(k).click()
    Case_ID=browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[1]/td[4]/div').text

    
    browser.find_element_by_link_text('[Change]').click()
    select = Select(browser.find_element_by_id("newOwn_mlktp"))
    select.select_by_index(1)
    browser.find_element_by_id("newOwn").send_keys("Mellanox Technical Support")
    browser.find_element_by_xpath('//*[@title="Save"]').click()

    browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[2]/td[2]/div/a').click()


                                
    assets_Size_int=browser.find_element_by_id("00N50000002RhrW_ileinner").text
    if(int(assets_Size_int)>5):
        try:
            browser.find_element_by_partial_link_text('Go to list (').click()
        except:
            print("No need for extend the list\n")

    
    Data = browser.find_elements_by_class_name('dataCell')
    
    sn=""
    index=[0,1,2,3,4,8]
    assetss = []  
    if(int(assets_Size_int)>24):
        for i in range(25):
            assetss.append(Asset(Data[index[0]+(10*i)].text,Data[index[1]+(10*i)].text,Data[index[2]+(10*i)].text,Data[index[3]+(10*i)].text,Data[index[4]+(10*i)].text,Data[index[5]+(10*i)].text))
            sn=sn+Data[index[0]+(10*i)].text+"\n"
        browser.find_element_by_partial_link_text('Next Page>').click()
        Data = browser.find_elements_by_class_name('dataCell')
        for i in range(int(assets_Size_int)-25):
            assetss.append(Asset(Data[index[0]+(10*i)].text,Data[index[1]+(10*i)].text,Data[index[2]+(10*i)].text,Data[index[3]+(10*i)].text,Data[index[4]+(10*i)].text,Data[index[5]+(10*i)].text))
            sn=sn+Data[index[0]+(10*i)].text+"\n"
    else:
        for i in range(int(assets_Size_int)):
            assetss.append(Asset(Data[index[0]+(10*i)].text,Data[index[1]+(10*i)].text,Data[index[2]+(10*i)].text,Data[index[3]+(10*i)].text,Data[index[4]+(10*i)].text,Data[index[5]+(10*i)].text))
            sn=sn+Data[index[0]+(10*i)].text+"\n"
          

    
    print("Current Page Title is : %s" %browser.title)
    

    my_lst_str = ''.join(map(str, assetss))
    header="Number of assets: "+assets_Size_int+"\n"
    sn=sn+"\n"
    msg.set_content(header+sn+my_lst_str)
    msg['Subject'] =Case_ID
    server.send_message(msg)
    server.quit()


def download_images():
# In order to fetch the image online  
    try:
        import urllib.request as url
    except ImportError:
        import urllib as url
    url.urlretrieve("https://avatars2.githubusercontent.com/u/5813145?s=280&v=4", "Mellanox.png")



# Designing Main(first) window
 
def main_account_screen():
    global register_screen
    register_screen = Tk()
    register_screen.title("Mellanox AutoAgent")
    register_screen.geometry("450x480")
    global username
    global password
    global username_entry
    global password_entry
    username = StringVar(register_screen, value='mahmouds@mellanox.com')
    password = StringVar(register_screen, value='Hppaviliondv_2')
 
    Label(register_screen, text="").pack()
    username_lable = Label(register_screen, text="Email:")
    username_lable.pack()
    username_entry = Entry(register_screen, textvariable=username)
    username_entry.pack()
    password_lable = Label(register_screen, text="\nPassword: ")
    password_lable.pack()
    password_entry = Entry(register_screen, textvariable=password, show='*')
    password_entry.pack()
    Label(register_screen, text="").pack()
    Button(text="Login", height="1", width="10", command = login).pack()
    download_images()
    widget = Label(register_screen, compound='top')
    widget.lenna_image_png=PhotoImage(file="Mellanox.png")
   
    widget['text'] = "Ver: 1.00"
    widget['image'] = widget.lenna_image_png
    widget.pack()

    register_screen.mainloop()
 
 
main_account_screen()







k = input("Please enter your case\n")
browser.stop_client()
browser.close()
browser.quit()