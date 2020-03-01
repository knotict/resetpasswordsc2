##Step 1 : Get access token by follow link => https://oauthplay.azurewebsites.net
from pyOutlook import OutlookAccount
#account_one = OutlookAccount('[YOUR_ACCOUNT_ACCESS_TOKEN')
account_one = OutlookAccount('eyJ0eXAiOiJKV1QiLCJub25jZSI6ImhVcUd5QUpRTHg0NlBuTDN3TnFJWmo2ZVJxd2Y5XzVzbEFzX2QwMW9lZ2MiLCJhbGciOiJSUzI1NiIsIng1dCI6IkhsQzBSMTJza3hOWjFXUXdtak9GXzZ0X3RERSIsImtpZCI6IkhsQzBSMTJza3hOWjFXUXdtak9GXzZ0X3RERSJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2Q5Y2Q0ODVlLTM5YmQtNGNjOS05NTM5LThjOTc2MzFmYmI3MS8iLCJpYXQiOjE1ODMwNzQxNDAsIm5iZiI6MTU4MzA3NDE0MCwiZXhwIjoxNTgzMDc4MDQwLCJhY2N0IjowLCJhY3IiOiIxIiwiYWlvIjoiQVRRQXkvOE9BQUFBL0JIVWRyNmwxYm9OUVpRYzk4YTVqTERLQVlDa0p1dDdIK1pGcHM2NWN6c3J3S0UzT0dIS1BZL0kzVjQ3c2lwZyIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiT0F1dGggU2FuZGJveCIsImFwcGlkIjoiMzI2MTNmYzUtZTdhYy00ODk0LWFjOTQtZmJjMzljOWYzZTRhIiwiYXBwaWRhY3IiOiIxIiwiZW5mcG9saWRzIjpbXSwiZmFtaWx5X25hbWUiOiJLYW5qYW5hcGEiLCJnaXZlbl9uYW1lIjoiVGhpdGlwb25nIiwiaXBhZGRyIjoiNDkuMjI4LjEzNy4xMzkiLCJuYW1lIjoiVGhpdGlwb25nIEthbmphbmFwYSAoVERFTSkiLCJvaWQiOiJjZTZkYWZlNi1mNjI5LTQ5MzctOWVmMi01MTA1Zjk2NDMwZTMiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMTE4ODQ5NDI0My0zMzUwNzU0NTkzLTQxMTc4Njk5NjMtMTU5OTEiLCJwdWlkIjoiMTAwMzIwMDA4MDc2M0M3MyIsInNjcCI6IkNhbGVuZGFycy5SZWFkV3JpdGUgQ2FsZW5kYXJzLlJlYWRXcml0ZS5TaGFyZWQgQ29udGFjdHMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZS5TaGFyZWQgTWFpbC5SZWFkV3JpdGUgTWFpbC5SZWFkV3JpdGUuU2hhcmVkIE1haWwuU2VuZCBNYWlsLlNlbmQuU2hhcmVkIE1haWxib3hTZXR0aW5ncy5SZWFkV3JpdGUgUGVvcGxlLlJlYWQgVGFza3MuUmVhZFdyaXRlIFRhc2tzLlJlYWRXcml0ZS5TaGFyZWQgVXNlci5SZWFkQmFzaWMuQWxsIiwic2lkIjoiOTIxNjg3ZGMtMmQ1YS00YjczLThmZTctNTczNzU3ZmU0ZWZkIiwic3ViIjoiM1pKdF9iSUtrV0wwMUYzNFh3djlJY3d6eDRnaENQWWk2Vko4VVlfLXEwUSIsInRpZCI6ImQ5Y2Q0ODVlLTM5YmQtNGNjOS05NTM5LThjOTc2MzFmYmI3MSIsInVuaXF1ZV9uYW1lIjoidGhpdGlwb25nX2thbkB0ZGVtLnRveW90YS1hc2lhLmNvbSIsInVwbiI6InRoaXRpcG9uZ19rYW5AdGRlbS50b3lvdGEtYXNpYS5jb20iLCJ1dGkiOiJablRrUHJsdVEwbUVreWlyQWFtVEFBIiwidmVyIjoiMS4wIn0.tRkMYIRJqh8TtgCDMZeGuSDVEgAES9NRoNC5M0ZapjxnaZqn9iqg5M3tjJrfIuNFw0GA0g3YPFlQF8kD5I64mCEZG2wjZE-KKxky_nVS3-PAl6X5wLuwX_B3y1j0rzdws96y2snkjhqXBxDdpEnjah2vVWGBLZsQNHUeNzrIuHz7MJs0tMCvy191VBtrWV7g-p7LXE68NVPhCll_R0Z8syO7dQxPIRJRAtuK6hJ2YUfqZzYa5nNkp7zYxcgA_StWO4wAKVbfbqI2C1SZVImIwFE03Y5jHK9owEMFmrneAhwqwvkoW7w8oSQj4xsAay_BYQaegg2a3_n2qIwcwUiajw')

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
#InboxFolder = account_one.get_folder_by_id('AAMkADY2YWNiMGUyLTU1ZGItNDE1OS05YjA2LWMzYWY0NTUyMjU0OQAuAAAAAAAlGM1Zqru-QZBcVwv5AIrFAQDezCz2ks0zTo46VfEhvVg_AAAAAAENAAA=')
#print(InboxFolder.total_items)
#InboxSubFolders = InboxFolder.get_subfolders()
#for InboxSubFolder in InboxSubFolders:
#     print(InboxSubFolder.name+","+InboxSubFolder.id)
#SC2Folder = account_one.get_folder_by_id('YOUR_SC2_FOLDER_ID')
SC2Folder = account_one.get_folder_by_id('AAMkADY2YWNiMGUyLTU1ZGItNDE1OS05YjA2LWMzYWY0NTUyMjU0OQAuAAAAAAAlGM1Zqru-QZBcVwv5AIrFAQDezCz2ks0zTo46VfEhvVg_AAKCF6LIAAA=')
print(SC2Folder.total_items)

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
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
