#@By Mahmoud sayid
class Task:

    def __init__(self,case_number,command) -> None:
        self.case_number = case_number
        self.command= command
    def __str__(self):
        return "Serial number: {0}   |    GUID: {1}\n".format(self.case_number,self.command)

    def __repr__(self):
        return "Task has been done {0} |  {1}\n".format(self.case_number,self.command)


class Basic_asset:
    def __init__(self,serial_number,GUID=None):
        self.serial_number = serial_number
        self.GUID= GUID


    def __str__(self):
        return "Serial number: {0}   |    GUID: {1}\n".format(self.serial_number,self.GUID)

    def __repr__(self):
        return "{0}   {1}\n".format(self.serial_number,self.GUID)

class Asset(Basic_asset):
    def __init__(self, serial_number,part_number,warranty,dOA,pD,aD,GUID=None):
        super().__init__(serial_number,GUID)
        self.part_number = part_number
        self.warranty = warranty
        self.dOA = dOA
        self.pD = pD
        self.aD = aD
#bbtt
    def __str__(self):
        return "\nSerial Number:  {0}\nPart Number:  {1}\nWarranty:  {2}\nDOA:  {3}\nproblem description:\n{4}\nPart description:\n{5}\n===========================================================\n".format(self.serial_number, self.part_number, self.warranty, self.dOA, self.pD,self.aD)
    def __repr__(self):
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
imports("cryptography")
imports("pywin32")
import threading
from queue import Queue
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
cases_Q = Queue()
import pythoncom
Tracebility_URL="https://360.mellanox.com/mquery/trace.php"


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
    main_url = "https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com"
    
    options = webdriver.ChromeOptions()
    options.arguments.append("--no-sandbox")
    options.arguments.append("--log-level=3")
    options.arguments.append("--disable-dev-shm-usage")
    options.arguments.append("--window-position=-32000,-32000")# to hide the browser

    try:
        options.arguments.append("--user-data-dir="+Main_path+"\\selenium") 
        browser = webdriver.Chrome(executable_path=Main_path+"\\chromedriver.exe", chrome_options=options)
        browser.get(main_url)
        while not page_is_loading(browser):
                continue
        browser.implicitly_wait(10)

        if ("No matches found" and "Search Feeds" not in browser.page_source):
            chrome_possion("show")
            messagebox.showinfo(title="Warning ", message="Please sign in")

    except Exception as identifier:
        print(str(identifier))
 
    
# Designing window for registration
 
def chrome_possion(mode=None):
    if mode=="hide":
        browser.set_window_position(-32000,-32000)
    if mode=="show":
        browser.set_window_position(0,0)
     

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


# Designing window for login 
def close():
    register_screen.destroy()

def Traceability(path,data):
    def write_list(result_list):
        with open(path+'\\Results.txt', 'w') as f:
            for asset in result_list:
                f.write("%s" % asset)
    chrome_possion("hide")
    result_list=[]
    browser.get(Tracebility_URL)

    
    assets_list=data.split("\n")

    for asset in assets_list:
        if asset!='':
            browser.find_element_by_xpath("/html/body/div/form/div/input").send_keys(asset)
            browser.find_element_by_xpath('//button[text()="Search"]').click()
            GUID=browser.find_element_by_xpath("/html/body/div/div[4]/div[2]/li[2]/b").text
            browser.find_element_by_xpath("/html/body/div/form/div/input").clear()
            result_list.append(Basic_asset(asset,GUID))
        pass
    write_list(result_list)
    messagebox.showinfo(title="Warning ", message="Done, Plesae find result.txt")
def case_comment(task:Task):
    #chrome_possion("show")
    command=task.command
    case_Number=task.case_number
    SF_lightning = "https://mellanox.lightning.force.com/lightning/r/Case/5001T00001QnhArQAJ/view"
    browser.get(SF_lightning)
    while not page_is_loading(browser):
        continue
    time.sleep(1)
    browser.find_element_by_xpath('//div[@class="uiInput uiAutocomplete uiInput--default"] //input[@title="Search..."]').click()
    while not page_is_loading(browser):
        continue
    time.sleep(1)
    browser.find_element_by_xpath('//div[@class="uiInput uiAutocomplete uiInput--default"] //input[@title="Search..."]').send_keys(case_Number)
    while not page_is_loading(browser):
        continue
    time.sleep(1)
    browser.find_element_by_xpath('//mark[contains(text(),'+case_Number+')]/..').click()
    while not page_is_loading(browser):
        continue
    time.sleep(1)
    Case_id=browser.find_element_by_xpath('//div[@class="windowViewMode-maximized active lafPageHost"] //lightning-formatted-text[contains(text(),"ref")]').text
    contact=browser.find_element_by_xpath('//div[@class="windowViewMode-maximized active lafPageHost"] //p[contains(text(),"Contact Name")]//..//a').text
    result={
        'case First reply':lambda x:0,
        'case LOGS': lambda x: 2,
        'case Department': lambda x: 3,
        'case Reviewing': lambda x: 1,
        'case Reviewing replying': lambda x: 4
    }[task.command](None)
    


    message= get_message(result)
    message="Hello "+contact+",\n\n\n"+message
    public_comment(Case_id,message)
    browser.find_element_by_xpath('//header[@class="slds-global-header_container branding-header slds-no-print oneHeader"] //button[@title="Close '+case_Number+'"]').click()

def menu_page():
    global Main_menu
    Main_menu = Toplevel(register_screen)
    Main_menu.title("Mellanox AutoAgent")
    Main_menu.geometry("450x450")
    Label(Main_menu, text="Please select you agent").pack()
    global username_login_entry
    widget = Label(Main_menu, compound='top')
    widget.lenna_image_png=PhotoImage(file=Main_path+"\\Mellanox.png")
    widget['image'] = widget.lenna_image_png
    Button(Main_menu, text="RMA Agent", width=30, height=1, command = lambda: swap_page("RMA_screen")).pack()
    Button(Main_menu, text="SN to GUID", width=30, height=1, command = lambda: swap_page("Traceability_screen")).pack()
    Button(Main_menu, text="Case Agent", width=30, height=1, command = lambda: swap_page("Case Agent")).pack()
    widget.pack()
    register_screen.withdraw()
    Main_menu.protocol("WM_DELETE_WINDOW", close)
def Traceability_page():
    #chrome_possion("show")#need to deeeletee
        global file_path
        content=""
        file_path=StringVar()

        def open_file(): 
            global file_location
            nonlocal content
            file = askopenfile(mode ='r', filetypes =[('input text', '*.txt')]) 
            if file is not None: 
                content = file.read() 
                file_path.set(file.name)
                file_location=os.path.dirname(file.name)
                browser.get(Tracebility_URL)
                while not page_is_loading(browser):
                    continue
                if ("Sign In" in browser.page_source):
                    chrome_possion("show")
                    messagebox.showinfo(title="Warning ", message="Please sign in")
    
        global Traceability_screen
        Traceability_screen = Toplevel(Main_menu)
        Traceability_screen.title("Mellanox AutoAgent")
        Traceability_screen.geometry("450x450")
        Label(Traceability_screen, text="Please load input file text").pack()
        global username_login_entry
        widget = Label(Traceability_screen, compound='top')
        widget.lenna_image_png=PhotoImage(file=Main_path+"\\Mellanox.png")
        widget['image'] = widget.lenna_image_png
        btnBrowseD = Button(Traceability_screen, text ='Browse input file',justify=CENTER,activebackground="light blue" ,command = lambda:open_file())
        btnBrowseD.place(relx = 0.12, rely = 0.75, anchor = CENTER)

        e1=Entry(Traceability_screen, textvariable=file_path,relief=SUNKEN, width=45)
        e1.place(relx = 0.60, rely = 0.75, anchor = CENTER)
        Button(Traceability_screen, text="Convert file", command = lambda: Traceability(file_location,content)).place(relx = 0.12, rely = 0.85, anchor = CENTER)
        Button(login_screen, text="Back", command = lambda:[login_screen.withdraw(),menu_page()]).pack()
        widget.place( relx = 0.5, rely = 0.35,anchor = CENTER)
        Main_menu.withdraw()
        Traceability_screen.protocol("WM_DELETE_WINDOW", close)
def RMA_page():
    global login_screen
    login_screen = Toplevel(Main_menu)
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
    Button(login_screen, text="Dequeue+Summary+First_reply", width=30, height=1, command = lambda: RMA_task("all")).pack()
    Button(login_screen, text="Dequeue", width=30, height=1, command = lambda: RMA_task("DQ") ).pack()
    Button(login_screen, text="summary", width=30, height=1, command = lambda: RMA_task("comment")).pack()
    Button(login_screen, text="Dequeue+Summary", width=30, height=1, command = lambda: RMA_task()).pack()
    Button(login_screen, text="First reply", width=30, height=1, command = lambda: RMA_task("review")).pack()
    Button(login_screen, text="Back", width=30, height=1, command = lambda:[login_screen.withdraw(),menu_page()]).pack()
    widget.pack()
    Main_menu.withdraw()
    login_screen.protocol("WM_DELETE_WINDOW", close)
def Case_page():
    global login_screen
    login_screen = Toplevel(Main_menu)
    login_screen.title("Mellanox AutoAgent")
    login_screen.geometry("450x450")
    Label(login_screen, text="Please enter case Details bellow:").pack()
    global username_verify
    username_verify = StringVar()
    global username_login_entry
    Label(login_screen, text="case number:").pack()
    username_login_entry = Entry(login_screen, textvariable=username_verify)
    username_login_entry.pack()
    widget = Label(login_screen, compound='top')
    widget.lenna_image_png=PhotoImage(file=Main_path+"\\Mellanox.png")
    widget['image'] = widget.lenna_image_png
    Button(login_screen, text="First reply", width=30, height=1, command = lambda: case_task("case First reply")).pack()
    Button(login_screen, text="Reviewing after reply", width=30, height=1, command = lambda: case_task("case Reviewing")).pack()
    Button(login_screen, text="Reviewing after LOGS", width=30, height=1, command = lambda: case_task("case LOGS") ).pack()
    Button(login_screen, text="relavent Department after info", width=30, height=1, command = lambda: case_task("case Department")).pack()
    Button(login_screen, text="Department after replying", width=30, height=1, command = lambda: case_task("case Reviewing replying")).pack()
    Button(login_screen, text="Back", width=30, height=1, command = lambda:[login_screen.withdraw(),menu_page()]).pack()

    widget.pack()
    Main_menu.withdraw()
    login_screen.protocol("WM_DELETE_WINDOW", close)
def case_task(command):
    cases_Q.put(Task(username_verify.get(),command))
    
def swap_page(page):
    
    #chrome_possion("hide")

    if page=="Main_menu":
        menu_page()

    if page=="Traceability_screen":
        Traceability_page()

    if page=="RMA_screen":
        RMA_page()

    if page=="Case Agent":
        Case_page()


def login():

    saveCredential()
    chrome_driver_lunch()
    swap_page("Main_menu")

    
            
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
 

def search_related_issue():
    Data=[]
    RMA_page="https://wikinox.mellanox.com/display/FLS/RMA+page+for+script"
    browser.get(RMA_page)
    rows = browser.find_elements_by_tag_name("tr")
    for row in range(2,len(rows)-1):
        Data.append(rows[row].text.split("."))
    print(Data)

def get_message(x=0):
    #chrome_possion("show")
    wikinox_page="https://wikinox.mellanox.com/display/FLS/RMA+page+for+script"
    browser.get(wikinox_page)
    table = browser.find_elements_by_class_name("confluenceTd")
    message=table[x*4+3].text
    return message

def public_comment(case_id,messange):
   
    initializeOutlook()
    mail.Subject = "external "+case_id
    #mail.Subject = case_id
    mail.Body = messange
    print(mail.Body)
    mail.Send()    
def threadmanage():
    while(True):
        if not cases_Q.empty():
            task:Task=cases_Q.queue[0]
            if "case" not in task.command:
                process = threading.Thread(target=RMA, args=[task])
                process.start()
            else:
                process = threading.Thread(target=case_comment, args=[task])
                process.start()
            while(process.is_alive()):
                time.sleep(2)#sleep for 1 sec
                print("wait for previews thread to finish")
            print("previews thread finished")
            cases_Q.get()
            print(task)
        time.sleep(1)#sleep for 1 sec
        
    
def RMA_task(command=None):
    cases_Q.put(Task(username_verify.get(),command))
        
    

def RMA(task:Task):
    command=task.command
    rma_Number=task.case_number
    
    chrome_possion("hide")
    #chrome_possion("show")
    #chrome_driver_lunch()
    if rma_Number!="":
        
        browser.get(main_url)
        
        k=rma_Number
        k=k.strip()
        
        time.sleep(5)
        #browser.get_screenshot_as_file("screenshot.png")    
        browser.find_element_by_id("secondSearchText").clear()
        browser.find_element_by_id("secondSearchText").send_keys(k)
        browser.find_element_by_id("secondSearchButton").click()
    
        browser.find_element_by_link_text(k).click()
        accountName=browser.find_element_by_id("cas4_ileinner").text
        contact=browser.find_element_by_id("cas3_ileinner").text
        
        FA_request=browser.find_element_by_id("00N50000002SCEj_ileinner").text
        related_case=browser.find_element_by_id("CF00N50000002SFLs_ilecell").text

        

        Case_ID=browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/table/tbody/tr[1]/td[4]/div').text

        if command==None or command=="DQ" or command=="all":
            DQ()
        if command==None or command=="comment" or command=="all":
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
            search_related_issue()
            print(mail.Body)
            mail.Send()      

        if command=="all" or command=="review":
            message=get_message(x=0)
            message="Hello "+contact+",\n\n\n"+message
            public_comment(Case_ID,message)
            


def initializeOutlook():
    pythoncom.CoInitialize()
    global mail
    global outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'networking-support@nvidia.com'
    #mail.To = 'mh_mouds@hotmail.com'

    for account in outlook.Session.Accounts:
        if("@nvidia.com" in account.DisplayName):
            username=account.DisplayName
            print(account.DisplayName)

    return  StringVar(register_screen, value=username)


# Designing Main(first) window
 
def main_account_screen():
    process = threading.Thread(target=threadmanage, args=[])
    process.start()
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

