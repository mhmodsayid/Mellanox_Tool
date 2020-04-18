import win32com.client as win32
from datetime import datetime
from datetime import timedelta

outlook = win32.Dispatch('outlook.application').GetNamespace("MAPI")
#Data=outlook.mailItem
inbox  = outlook.GetDefaultFolder(6)
messages = inbox.Items


currentDT = datetime.now()
print (str(currentDT))
currentDT=currentDT-timedelta(minutes=20)

print (currentDT.strftime("%m/%d/%Y %H:%M:%S"))

message=messages.find("[ReceivedTime] >'"+ str(currentDT.strftime("%m/%d/%Y %H:%M"))+"' And [From]='Mahmoud Sayid'")



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



#Mellanox Support Admin <supportadmin@mellanox.com>

