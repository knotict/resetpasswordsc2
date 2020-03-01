import rpa as r
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

df = pd.read_excel('pandas_simple.xlsx', sheet_name='Sheet1')
count = 1
while (count < 11):     
    count = count + 1
    r.init()
    sc2url = df.iloc[0, count]
    sc2user = df.iloc[1, count]
    temppass = df.iloc[2, count]
    newpass = "P@ssw0rd"
    r.url(sc2url)
    r.click('//*[@id="internal-panel"]/label[1]/span/a')
    r.type('#username', '[clear]')
    r.type('#username', sc2user)
    r.type('#password', '[clear]')
    r.type('#password', temppass+'[enter]')
    r.type('#newPassword','[clear]')
    r.type('#newPassword',newpass)
    r.type('#confirmNewPassword','[clear]')
    r.type('#confirmNewPassword',newpass+'[enter]')
    r.snap('page', 'results'+str(count)+'.png')
    r.close()