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
except ImportError:
    print("installing selenium....")
    os.system("pip install "+ package)   
package="cryptography"
try:
    __import__(package)
except ImportError:
    print("installing cryptography....")
    os.system("pip install "+ package)   
from cryptography.fernet import Fernet
from selenium import webdriver
from selenium.webdriver.support.ui import Select
# Import smtplib for the actual sending function
import smtplib

# Import the email modules we'll need
from email.message import EmailMessage
from tkinter import *

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
        url.urlretrieve("https://ucb1d9488fa18564d2e9e0cd572c.dl.dropboxusercontent.com/cd/0/get/A0pvCQ4jHHt9nsbsVgjORq34Gj0yzi1pEVXG1YrCKVdqR-5siXliNVI9S5z3kDt8jFdMaCjMxeCDF4heEzYJ78F-CNiof00yKHPWTy_t741ADAua4InXC5L10C4lczbPCuE/file?_download_id=52500573738052191167474377859756385038390858859054808382686637678&_notify_domain=www.dropbox.com&dl=1", "chromedriver.exe")
        print("Downloading completed.")



download_images()


options = webdriver.ChromeOptions()
#options.arguments.append("headless")#hide the browser
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
    msg = EmailMessage()
    msg['From'] = username.get()
    msg['To'] = 'supportadmin@mellanox.com'
    server = smtplib.SMTP('smtp.office365.com.', 587)
    server.ehlo()
    server.starttls()
    #Next, log in to the server
    #server.login(username.get(), password.get())
    
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
    msg.set_content(header+sn+my_lst_str)
    msg['Subject'] =Case_ID
    #server.send_message(msg)

    #-----------------------------------------------find RMA related
    rma_search_url="https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com#!/str='219801A1E98193W00002'&initialViewMode=detail&fen=a0F&collapse=1"
    browser.get(rma_search_url)
    AssetsSn=sn.replace("\n"," or ")
    browser.find_element_by_id("secondSearchText").send_keys(AssetsSn)
    browser.find_element_by_id("secondSearchButton").click()#search
    resultss=browser.find_elements_by_class_name('searchFirstCell')
    number_of_search_in_RMA=resultss[0].text
    temp=number_of_search_in_RMA.split("(")
    temp=temp[1].split(")")
    number_of_RMA=temp[0]

    if(number_of_RMA.find("+")):
        print("+too many")
    else:
        if int(number_of_RMA)>int(assets_Size_int):
            print("there is related case")
            pass

    Case_search_url="https://mellanox.my.salesforce.com/_ui/search/ui/UnifiedSearchResults?searchType=2&sen=ka&sen=00O&sen=01t&sen=00T&sen=00U&sen=a1G&sen=00a&sen=a43&sen=0F9&sen=a2v&sen=a26&sen=a28&sen=a2z&sen=0TO&sen=02s&sen=001&sen=a4F&sen=068&sen=003&sen=a4H&sen=005&sen=500&sen=a0F&sen=a0E&sen=a0H&sen=a0M&str=00751864%20&isdtp=vw&isWsVw=true&isWsVw=true&nonce=29c527472e4db3b0434215c49cac45bcaf9a8ac6f5473c5acccb8b58d751b8ce&sfdcIFrameOrigin=https%3A%2F%2Fmellanox.my.salesforce.com#!/str=.&initialViewMode=detail&fen=500&collapse=1&searchAll=true"
    browser.get(Case_search_url)
    browser.find_element_by_id("secondSearchText").send_keys(AssetsSn)
    browser.find_element_by_id("secondSearchButton").click()#search
    resultss=browser.find_elements_by_class_name('searchFirstCell')
    related_Case=resultss[0].text
    temp=related_Case.split("(")
    temp=temp[1].split(")")
    if(temp[0].find("+")):
        print("+too many")
    else:
        if int(temp[0])!=0:
            print("there is related case" )
            pass
    



    #need to check if assets allready in preview RMA
    #SnList=sn.splitlines()
    
    #need if RMA related to case




    server.quit()

#



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