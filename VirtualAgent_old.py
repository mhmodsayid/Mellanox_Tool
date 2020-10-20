#@By Mahmoud sayid
class Asset:
    def __init__(self, serial_number,part_number,warranty,dOA,pD,aD):
        self.serial_number = serial_number
        self.part_number = part_number
        self.warranty = warranty
        self.dOA = dOA
        self.pD = pD
        self.aD = aD

#bbtt
    def __str__(self):
        return "\nSerial Number:  {0}\nPart Number:  {1}\nWarranty:  {2}\nDOA:  {3}\nproblem description:\n{4}\nPart description:\n{5}\n===========================================================\n".format(self.serial_number, self.part_number, self.warranty, self.dOA, self.pD,self.aD)
#import modules    
def page_is_loading(driver):
    while True:
        x = driver.execute_script("return document.readyState")
        if x == "complete":
            return True
        else:
            return False



import pickle
import os

import time
package="selenium"

Main_path=os.path.expanduser('~\\MTools')

try:
    __import__(package)
    print(package+ " imported successfully ")
except ImportError:
    print("installing selenium....")
    os.system("pip install "+ package+" --user")   
package="cryptography"
try:
    __import__(package)
    print(package+ " imported successfully")
except ImportError:
    print("installing cryptography....")
    os.system("pip install "+ package+" --user")   
package="pywin32"
try:
    __import__(package)
    print(package+ " imported successfully")
except ImportError:
    print("installing pywin32....")
    os.system("pip install "+ package+" --user")   
import tkinter
from cryptography.fernet import Fernet
from selenium import webdriver
from selenium.webdriver.support.ui import Select
# Import smtplib for the actual sending function
import smtplib
import win32com.client as win32
# Import the email modules we'll need
from email.message import EmailMessage
from tkinter import *
from tkinter import simpledialog
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


def chrome_driver_lunch(mode):
    global main_url
    main_url = "https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com"
    global browser
    options = webdriver.ChromeOptions()
    options.arguments.append("--log-level=3")
    if mode:
        options.arguments.append("--headless")#hide the browser

    try:
        options.arguments.append("user-data-dir="+Main_path+"\\selenium") 
        browser = webdriver.Chrome(executable_path=Main_path+"\\chromedriver.exe", chrome_options=options)
        browser.get(main_url)
        browser.implicitly_wait(5)
    except Exception as identifier:
        print(str(identifier))

    
# Designing window for registration
 

download_images()


def saveCredential():
    def write_key():
        try:
            os.makedirs((os.path.expanduser('~')+'\\MTools'))
        except FileExistsError:
            # directory already exists
            pass
        key = Fernet.generate_key()
        with open((os.path.expanduser('~')+'\\MTools'+'\\key.key'), "wb") as key_file:
            key_file.write(key)
        
    def load_key():
        return open((os.path.expanduser('~')+'\\MTools'+'\\key.key'), "rb").read()
    write_key()
    key = load_key()
    f = Fernet(key)
    bpas=bytes(password.get(), 'ascii')
    encrypted = f.encrypt(bpas)
    path=os.path.join(os.path.expanduser('~'),'MTools',"Tool.bin")
    try:
        os.makedirs((os.path.expanduser('~')+'\\MTools'))
    except FileExistsError:
    # directory already exists
        pass
    newFile = open(path, "wb")
    newFile.write(encrypted)
    newFile.close()
    path=os.path.join(os.path.expanduser('~'),'MTools',"Username.bin")
    newFile = open(path, "wt")
    newFile.write(username.get())
    newFile.close()
   
def LoadCredential(password,username):
    def load_key():
        return open((os.path.expanduser('~')+'\\MTools'+'\\key.key'), "rb").read()
        
    # directory already exists
        
    try:
        path=os.path.join(os.path.expanduser('~'),'MTools',"Tool.bin")
        key = load_key()
        f = Fernet(key)
        newFile = open(path, "rb")
        temp=newFile.read()
        decrypted_encrypted = f.decrypt(temp)
        decrypted_encrypted=decrypted_encrypted.decode("utf-8") 
        password = StringVar(register_screen, value=decrypted_encrypted)
        path=os.path.join(os.path.expanduser('~'),'MTools',"Username.bin")
        newFile = open(path, "rt")
        username = StringVar(register_screen, value=newFile.read())


        newFile.close()
    except OSError:

        print("no key avalible")
    return password,username


def Cookie_check():
    pass


# Designing window for login 
def close():
    register_screen.destroy()


def login():

    saveCredential()
    
    if (Path(Main_path+"\\selenium")).exists():#if not first time lunch 
        chrome_driver_lunch(0)
        try:

            browser.find_element_by_xpath("/html/body/div/form[1]/div/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div").click()

            if ("Sign in" in browser.page_source):
                
                print("Error not saved")
     
        except Exception as identifier:
            #login from mellanox no need for passowrd and username
            browser.get_screenshot_as_file("screenshot.png")    

            print(str(identifier))
            print("cookie avalible")

            
    else:#first time lunch
        chrome_driver_lunch(0)
        try:
            browser.find_element_by_id("i0116").send_keys(username.get())#email address
            browser.find_element_by_id("idSIButton9").click()#next 
            browser.implicitly_wait(10)#to check if the elemnt exists fast
            browser.find_element_by_id("passwordInput").send_keys(password.get())
            browser.find_element_by_id("submitButton").click()
           
        except Exception as identifier:
            #login from mellanox no need for passowrd and username
            browser.get_screenshot_as_file("screenshot.png")    
            print(str(identifier))
            print("cookie avalible")

            
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
    widget.lenna_image_png=PhotoImage(file=Main_path+"\\Mellanox.png")
    widget['image'] = widget.lenna_image_png


    Button(login_screen, text="Dequeue, Add summary comment", width=30, height=1, command = DQ_Sumary_OEM).pack()
    widget.pack()
    register_screen.withdraw()
    login_screen.protocol("WM_DELETE_WINDOW", close)
            
# Implementing event on register button
 
def search_for_related_RMA(SNlist):
    Duplicate=FALSE
    try:
        for SN in SNlist:
            if(SN==''):
                continue
            rma_search_url="https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com#!/str="+SN+"&initialViewMode=detail&fen=a0F&collapse=1"
            browser.get(rma_search_url)
            resultss=browser.find_element_by_class_name("searchFirstCell")

            while not page_is_loading(browser):
                continue
            time.sleep(1)
            resultss=browser.find_element_by_xpath('/html/body/div[2]/table/tbody/tr/td[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr/td[1]/h3/span')
            #resultss=browser.find_element_by_id("ext-gen48")
            #resultss=browser.find_element_by_id("ext-gen46")
            number_of_search_in_RMA=resultss.text
            temp=number_of_search_in_RMA.split("(")
            temp=temp[1].split(")")
            number_of_RMA=temp[0]

            if int(number_of_RMA)>1:
                Duplicate=True
                mail.Body="Duplicate SN: "+SN+"\n"+mail.Body

        if Duplicate:
            mail.Body="There is Previous RMA for the assets\n"+mail.Body
            
        else:
            mail.Body="There is NO Previous RMA for the assets\n"+mail.Body
    

    except Exception as EX:
    
        print(str(EX))
        pass


def search_for_previews_cases(SNlist):
    Duplicate=False
    try:
        for SN in SNlist:
            if(SN==''):
                continue

            rma_search_url="https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com#!/fen=500&initialViewMode=detail&collapse=1&str="+SN
            browser.get(rma_search_url)
            while not page_is_loading(browser):
                continue
            time.sleep(1)
            resultss=browser.find_element_by_id("sidetabLink-500")
            related_Case=resultss.text
            if(related_Case!=''):
                temp=related_Case.split("(")
                temp=temp[1].split(")")
                temp[0]=temp[0].replace('+', '')
                number_of_RMA=temp[0]

                if int(number_of_RMA)>1:
                    Duplicate=True
                    mail.Body="There is Related case for SN:= "+SN+"\n"+mail.Body

        if Duplicate:
            mail.Body="There is Previous case for the assets\n"+mail.Body
            
        else:
            mail.Body="There is NO Previous case for the assets\n"+mail.Body


    except Exception as EX:
        print(EX.with_traceback())
        print(str(EX))
        pass

def search_if_OEM(accountName):
    Wikinox_OEM="https://wikinox.mellanox.com/pages/viewpage.action?spaceKey=FLS&title=OEM+accounts+and+Procedure"
    browser.get(Wikinox_OEM)
    if (accountName in browser.page_source):
        print("OEM")
        mail.Body="OEM Account \n"+mail.Body
    else:
        mail.Body="Not a OEM Account \n"+mail.Body
    

# Implementing event on login button 

def DQ():
    browser.find_element_by_link_text('[Change]').click()
    select = Select(browser.find_element_by_id("newOwn_mlktp"))
    select.select_by_index(1)
    browser.find_element_by_id("newOwn").send_keys("Mellanox Technical Support")
    browser.find_element_by_xpath('//*[@title="Save"]').click()
 
def DQ_Sumary_OEM():
    browser.implicitly_wait(10)
    rma_Number = username_verify.get()
    if rma_Number!="":
        browser.get(main_url)
        
        k=rma_Number
        k=k.strip()
        
        time.sleep(5)
        browser.get_screenshot_as_file("screenshot.png")    
        browser.find_element_by_id("secondSearchText").clear()
        browser.find_element_by_id("secondSearchText").send_keys(k)
        browser.find_element_by_id("secondSearchButton").click()
    



        browser.find_element_by_link_text(k).click()
        accountName=browser.find_element_by_id("cas4_ileinner").text
        FA_request=browser.find_element_by_id("00N50000002SCEj_ileinner").text
        related_case=browser.find_element_by_id("CF00N50000002SFLs_ilecell").text

        

        Case_ID=browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[1]/td[4]/div').text

       
        DQ()

        #get RMA summery
        browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[2]/td[2]/div/a').click()                
        assets_Size_int=browser.find_element_by_id("00N50000002RhrW_ileinner").text
        if(int(assets_Size_int)>5):
            try:
                browser.find_element_by_partial_link_text('Go to list (').click()
            except:
                print("No need for extend the list\n")

        
        
        browser.find_element_by_xpath("//*[text()='RMA Type final']").click()
        Data = browser.find_elements_by_class_name('dataCell')
        sn=""
        index=[0,1,2,3,4,8]
        assetss = []  
        print(len(Data))
        numofassets=0
        i=0
        lastwarranty=""
        if(len(Data)/10)<int(assets_Size_int):
            while numofassets<int(assets_Size_int):#i is the number of assets
                for i in range(int(len(Data)/10)):
                    numofassets+=1

                    if Data[index[4]+(10*i)].text.find('...')!= -1:
                        browser.find_element_by_xpath('//*[@title='+Data[index[1]+(10*i)].text+']').click()
                        Data[index[4]+(10*i)].text=browser.find_element_by_id("00N50000001vo8N_ilecell").text
                        browser.back()
                        pass

                    assetss.append(Asset(Data[index[0]+(10*i)].text,Data[index[1]+(10*i)].text,Data[index[2]+(10*i)].text,Data[index[3]+(10*i)].text,Data[index[4]+(10*i)].text,Data[index[5]+(10*i)].text))
                    sn=sn+Data[index[0]+(10*i)].text+"\n"
                    if lastwarranty=="" or Data[index[2]+(10*i)].text!=lastwarranty :
                        lastwarranty=Data[index[2]+(10*i)].text
                        sn=sn+"\nWarranty: "+lastwarranty+"\n\n"
                        
                if numofassets<int(assets_Size_int):

                    browser.find_element_by_partial_link_text('Next Page>').click()
                    browser.find_element_by_xpath("//*[text()='RMA Type final']").click()
                    Data = browser.find_elements_by_class_name('dataCell')



        else:
            for i in range(int(assets_Size_int)):
                if Data[index[4]+(10*i)].text.find('...')!= -1:
                    Data[index[0]+(10*i)].click()
                    New_Problem_Description=browser.find_element_by_id("00N50000001vo8N_ileinner").text
                    browser.back()
                    Data = browser.find_elements_by_class_name('dataCell')
                    assetss.append(Asset(Data[index[0]+(10*i)].text,Data[index[1]+(10*i)].text,Data[index[2]+(10*i)].text,Data[index[3]+(10*i)].text,New_Problem_Description,Data[index[5]+(10*i)].text))
                    
                else:
                    assetss.append(Asset(Data[index[0]+(10*i)].text,Data[index[1]+(10*i)].text,Data[index[2]+(10*i)].text,Data[index[3]+(10*i)].text,Data[index[4]+(10*i)].text,Data[index[5]+(10*i)].text))
                sn=sn+Data[index[0]+(10*i)].text+"\n"
            

        
        print("Current Page Title is : %s" %browser.title)
        

        my_lst_str = ''.join(map(str, assetss))
        header="Number of assets: "+assets_Size_int+"\n"
        sn=sn+"\n"
        #msg.set_content(header+sn+my_lst_str)
        #msg['Subject'] =Case_ID
        initializeOutlook()
        mail.Subject = Case_ID
        mail.Body = header+sn+my_lst_str
        
        #server.send_message(msg)
        #set if FA needed
        mail.Body="Failure Analysis Request: "+FA_request+"\n"+mail.Body
        if related_case!="":
            mail.Body="Related case: "+related_case+"\n"+mail.Body
        ###-----------------------------------------------find RMA relate
        # d
        SNlist=sn.split("\n")

        search_for_previews_cases(SNlist)

        search_for_related_RMA(SNlist)

        search_if_OEM(accountName)
        
        print(mail.Body)
        mail.Send()        

       



def initializeOutlook():
    
    global mail
    global outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'supportadmin@mellanox.com'
    #mail.To = 'mh_mouds@hotmail.com'

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
        password,username=LoadCredential(password,username)
    
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


#browser.stop_client()
#browser.close()
#browser.quit()