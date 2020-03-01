##Step 1 : Get access token by follow link => https://oauthplay.azurewebsites.net
from pyOutlook import OutlookAccount
#account_one = OutlookAccount('[YOUR_ACCOUNT_ACCESS_TOKEN]')

##Step 2 : Uncomment below section, run and copy your inbox folder "id"
folders = account_one.get_folders()
for folder in folders:
    print(folder.name+","+folder.id)

##Step 3 : Uncomment below section, run and copy and put SC2 Folder "id"
from pyOutlook import *
from bs4 import BeautifulSoup
import re
import pandas as pd
import numpy as np
import xlsxwriter
#InboxFolder = account_one.get_folder_by_id('[YOUR_INBOX_FOLDER_ID]')
#print(InboxFolder.total_items)
#InboxSubFolders = InboxFolder.get_subfolders()
#for InboxSubFolder in InboxSubFolders:
#     print(InboxSubFolder.name+","+InboxSubFolder.id)
#SC2Folder = account_one.get_folder_by_id('YOUR_SC2_FOLDER_ID')
#print(SC2Folder.total_items)

df = pd.DataFrame(["a","b","c","d","e"],columns=['dummy'])
dataindex = 0
for message in SC2Folder.messages():
    # print(message.body)
    data = []
    soup = BeautifulSoup(message.body, 'html.parser')
    link = soup.find('a').get('href')
    data.append(link)
    for tag in soup.find_all("td",class_="Body1_1"):
        # print(tag.string)
        data.append(tag.string)
    df[str(dataindex)] = pd.Series(data, index=df.index)
    dataindex = dataindex + 1
writer = pd.ExcelWriter('book1.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
