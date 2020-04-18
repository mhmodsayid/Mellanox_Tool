import win32com.client as win32
from datetime import datetime
from datetime import timedelta
import time

outlook = win32.Dispatch('outlook.application').GetNamespace("MAPI")
#Data=outlook.mailItem
inbox  = outlook.GetDefaultFolder(6)
root_folder = inbox.Folders(4)
messages = root_folder.Items


print(root_folder.name)
while (True):
    messages = inbox.Items
    currentDT = datetime.now()
    print (str(currentDT))
    currentDT=currentDT-timedelta(minutes=20)

    print (currentDT.strftime("%m/%d/%Y %H:%M:%S"))

    message=messages.find("[ReceivedTime] >'"+ str(currentDT.strftime("%m/%d/%Y %H:%M"))+"' And [From]='Mahmoud Sayid' And [Subject]='& QUANTA &'")



    #message = messages.GetFirst ()
    try:
        received_time = str(message.ReceivedTime)
        print  (message.Subject+"time "+received_time )
        messages.Sort 
        message = messages.GetFirst ()
        received_time = str(message.ReceivedTime)
        print  (message.Subject+"time "+received_time )

        print ('total messages: ', len(messages))

    except Exception as err:
        print(err)
    time.sleep(2.4)



#Mellanox Support Admin <supportadmin@mellanox.com>

