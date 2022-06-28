import win32com.client

#other libraries to be used in this script
import os
import datetime as dt
import pandas as pd
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace('MAPI')
for account in mapi.Accounts:
    print(account.DeliveryStore.DisplayName)

# for idx, folder in enumerate(mapi.Folders("yasin.mohammed@d.com").Folders):
#     print(idx+1, folder)
# or using index to access the folder
# for idx, folder in enumerate(mapi.Folders(1).Folders):
#     print(idx+1, folder)

messages = mapi.Folders("yasin.mohammed@d.com").Folders("Inbox").Items

# start:seperate with timeweek
lastWeekDateTime = dt.datetime.now() - dt.timedelta(minutes = 190)
lastWeekDateTime = lastWeekDateTime.strftime('%m/%d/%Y %H:%M %p')  #<-- This format compatible with "Restrict"
messages = messages.Restrict("[ReceivedTime] >= '" + lastWeekDateTime +"'")
# end

#start: check with sub
messages = messages.Restrict("[Subject] = 'CHECKTHISOUT'")

# or
#messages = mapi.Folders(1).Folders(2).Items
for msg in list(messages)[:10]:
    print(msg.body)
    for attachment in msg.Attachments:
        fileName = 'file_' + attachment.FileName
        print(attachment.FileName)
        print(fileName)
        #attachment.SaveAsFile(os.path.join(r"C:\Users\S974009\Desktop\file_" + attachment.FileName))

df = pd.read_excel(r"C:\Users\S974\Desktop\file_AWS_CTS_LZ_Unused_Resources_16thJune2022.xlsx", sheet_name=[5])

print(df)
