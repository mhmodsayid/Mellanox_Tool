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
        return "Serial Number:  {0}\nPart Number:  {1}\nWarranty:  {2}\nDOA:  {3}\nproblem description:\n{4}\nPart description:\n{5}\n===========================================================\n".format(self.serial_number, self.part_number, self.warranty, self.dOA, self.pD,self.aD)
#import modules    


import pickle
import os
package="selenium"
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
def download_images():
# In order to fetch the image online  
    try:
        import urllib.request as url
    except ImportError:
        import urllib as url
    print("Downloading compoments....")
    if os.path.isfile('Mellanox.png'):
        print ("File exist")
    else:
        print ("File not exist downloading file...")
        url.urlretrieve("https://avatars2.githubusercontent.com/u/5813145?s=280&v=4", "Mellanox.png")
    
    try:
        import urllib.request as url
    except ImportError:
        import urllib as url
    if os.path.isfile('chromedriver.exe'):
        print ("File exist")
    else:
        print ("File not exist downloading Driver...")
        url.urlretrieve("https://raw.githubusercontent.com/mhmodsayid/Mellanox_Tool/master/chromedriver.exe?token=AHSZ7S4PCNBASIBMTPSBS6S6TY5BI","chromedriver.exe")
        print("Downloading completed.")

#https://dl.boxcloud.com/d/1/b1!BWEj6KfqzFcPnfTFN6ouZ4LVngba7vhISr-U2RvORFH1uPxBVeWit6hngiQ2LAvt-0ZnT_0u4mD4qRA3BQ_Q6Mr5Tsf10U79i3648NWX___VziFjU38XzLdCP6Lwku0O9P2JdvjkBC7Z-ZgKaeOGwMrtXeqRgiH4nGOqaxYjA7y3A2dJeWrjiZ5QKWWdosjkNdBEncomooqOjOGSqy3dMzIp6aB00Re47UBT-QuszoIWrPtVx1cuAHnL-v-F8HDMQzO9mYLf3skUwV0gQ1xA-TJnG91VMS9hh2eX1Xl80phZPbs72Ii1EwcWDgdkGlpH9ciUrtGqjzETw_F0KoGNu0ifNfp9F-TN83xcCWale5qkFbeF8_5sqOJBuxY5zIaHoLF7Q03t9ioIcqLBHVx2RyuZxC5hf1W8zhJGf3fet3XuOAEPwqzKhCIw_oYV7bIfnb0L2J2Hv6BLUG4qtLynk-3IiiO7JN9UFaOa0yyIXS1QIWRjwMtr5yZx52tCGMLa5diHgHqNiaoXoG35coGo0GT-TuQmN87dVJWqYtN76aFEBSzbEIZrNtHV-jXBy7-F7_u1xPZgU6Sav8q2HlVbxka3U12E7q9CRuSHONZyxl_bT5ctFsWwBk_ic5Mm8MwRFX6SYC3qXdiRflkIGaWYb1clpIrllHYUFrwiJX2F8rp0kbQ-n68sk_AMXUO8jpyqKtUBKSD-3Xr5KZRsZQJxIqjHLNVlUscISjtVLbbfyVMq6noa6TLPUBqFlojKrBPiYQNaS6CQY0qfS1aAi2D0ztDChHKYKLnagu12w6UVpnrq-Kmx5vuqsd2kg7Pn00AYAte5fByGlQsWRTKEZLaKYZ9iwPIskl5W16ltydYLUebCoicBq0q9Ik021LZXr1pMMbHM-iy3SsIzkUX-qvzu7tnOvJTq1darpTDx70GBSBvlSV2Hzk2cTmHpk-asfam0WyhEBGoclqJ7xDhHoDKBTtSMLm8Hoh1brRunxg1SurxvWhRZoJ1TdlCQuH2ThZYcW8XWZeWb6tZdAeySVHLWsDVMV8LrBYLTKdPEfJecvr-EcedbPKqpBVBuT2gaWGW12ZmXU-UpBGYwxJDKMXuEFo9fIShpSJQiEOx6iz1DieU91b8omcgSgGLzSQ-Q7z5qnMIiXG4nbb5MOAqvcf_1gxCGtsXTVUKyTGzkXj6su0eB9YWosGumVA935inv-mG8bCo73MBMtCGeNl-bCuB58wrOuJATyTN2kZecE20ulTYqnIFwFSA_4589SGfmoWFnqzQlK7KzECr3MIPeETKFqM0BLRNcig../download


download_images()


options = webdriver.ChromeOptions()
options.arguments.append("headless")#hide the browser
browser = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=options)
main_url = 'https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com'
# Designing window for registration
 



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
        pass
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














# Designing window for login 
 
def login():

    saveCredential()
    browser.get(main_url)
    browser.implicitly_wait(5)
    browser.find_element_by_id("i0116").send_keys(username.get())
    browser.find_element_by_id("idSIButton9").click()
    

    try:
        browser.implicitly_wait(10)#to check if the elemnt exists fast
        browser.find_element_by_id("passwordInput").send_keys(password.get())
        browser.find_element_by_id("submitButton").click()
        code = simpledialog.askstring(title="Test",prompt="Please enter the code from your phone ")
        browser.find_element_by_id("idTxtBx_SAOTCC_OTC").send_keys(code)
        browser.find_element_by_id("idSubmit_SAOTCC_Continue").click()
        browser.find_element_by_id("idSIButton9").click()

        
    except NoSuchElementException:
        #login from mellanox no need for passowrd and username
        print("elements password not found")
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
    browser.implicitly_wait(50)
    rma_Number = username_verify.get()
    browser.get(main_url)

   # msg = EmailMessage()
    #msg['From'] = username.get()
    #msg['To'] = 'supportadmin@mellanox.com'
    #server = smtplib.SMTP('smtp.office365.com.', 587)
    #server.ehlo()
    #server.starttls()
    #Next, log in to the server
    #server.login(username.get(), password.get())

    
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    

    
    k=rma_Number
    k=k.strip()
        
    
    browser.find_element_by_id("secondSearchText").clear()
    browser.find_element_by_id("secondSearchText").send_keys(k)
    browser.find_element_by_id("secondSearchButton").click()
    browser.find_element_by_link_text(k).click()
    Case_ID=browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[1]/td[4]/div').text

    #RMA DEQ:
    browser.find_element_by_link_text('[Change]').click()
    select = Select(browser.find_element_by_id("newOwn_mlktp"))
    select.select_by_index(1)
    browser.find_element_by_id("newOwn").send_keys("Mellanox Technical Support")
    browser.find_element_by_xpath('//*[@title="Save"]').click()

    browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[2]/td[2]/div/a').click()

    #get RMA summery
                                
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
    #msg.set_content(header+sn+my_lst_str)
    #msg['Subject'] =Case_ID
    initializeOutlook()
    mail.Subject = Case_ID
    mail.Body = header+sn+my_lst_str
    
    #server.send_message(msg)

    ###-----------------------------------------------find RMA related
    
    try:
        rma_search_url="https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com#!/str='219801A1E98193W00002'&initialViewMode=detail&fen=a0F&collapse=1"
        browser.get(rma_search_url)
        AssetsSn=sn.replace("\n"," or ",int(assets_Size_int)-1)
        browser.find_element_by_id("secondSearchText").clear()
        browser.find_element_by_id("secondSearchText").send_keys(AssetsSn)
        browser.find_element_by_id("secondSearchButton").click()#search
        resultss=browser.find_element_by_id("ext-gen46")
        number_of_search_in_RMA=resultss.text
        temp=number_of_search_in_RMA.split("(")
        temp=temp[1].split(")")
        number_of_RMA=temp[0]

        if (number_of_RMA.find("+") != -1):
            print("+too many")
        else:
            if int(number_of_RMA)>=int(assets_Size_int):
                mail.Body="There is NO previews RMA for the assets\n"+mail.Body
                print("There is NO previews RMA for the assets")
            else:
                mail.Body="There is previews RMA for the assets\n"+mail.Body
                print("there is previews RMA for the assets")

    except Exception:
        print(Exception)
        pass
    
    try:
        Case_search_url="https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com#!/str=.&initialViewMode=detail&fen=500&collapse=1&searchAll=true"
        browser.get(Case_search_url)
        browser.find_element_by_id("secondSearchText").clear()
        browser.find_element_by_id("secondSearchText").send_keys(AssetsSn)
        browser.find_element_by_id("secondSearchButton").click()#search
        resultss=browser.find_element_by_id("ext-gen52")
        related_Case=resultss.text
        temp=related_Case.split("(")
        temp=temp[1].split(")")
        number_of_RMA=temp[0]
        if (number_of_RMA.find("+") != -1):
            print("+too many")
        else:
            if int(number_of_RMA)>=int(assets_Size_int):
                mail.Body="There is NO previews RMA case for the assets\n"+mail.Body
                print("There is NO previews RMA case for the assets")
            else:
                mail.Body="There is previews RMA case for the assets\n"+mail.Body
                print("there is previews RMA case for the assets")

    except Exception:
        print(Exception)
        pass

    mail.Send()        

    #need to check if assets allready in preview RMA
    #SnList=sn.splitlines()
    
    #need if RMA related to case

#

def initializeOutlook():
    
    global mail
    global outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'supportadmin@mellanox.com'

    for account in outlook.Session.Accounts:
        if("@mellanox.com" in account.DisplayName):
            username=account.DisplayName
            print(account.DisplayName)

    return  StringVar(register_screen, value=username)


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
    widget.lenna_image_png=PhotoImage(file="Mellanox.png")
   
    widget['text'] = "Ver: 1.00"
    widget['image'] = widget.lenna_image_png
    widget.pack()

    register_screen.mainloop()
 
 
main_account_screen()







k = input("Please enter your case\n")

#browser.stop_client()
#browser.close()
#browser.quit()