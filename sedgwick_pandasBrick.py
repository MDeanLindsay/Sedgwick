import pandas as pd
import numpy as np
import datetime as dt
from pandas import DataFrame 
from functools import partial
from pandas import ExcelWriter

date_parser = partial(pd.to_datetime, format = '%Y/%m/%d', errors = 'coerce')

df = pd.read_excel (r'C:\\Users\\michael_lindsay\\SedgwickTXT.xlsx',
    usecols = ['ClaimNumber','dateopened', 'dateclosed','datereopened', 'claimstatus', 'claimsubstatus'],
    parse_dates = ['dateopened', 'dateclosed', 'datereopened'],
    date_parser=date_parser)

'''
#Print XLSX statement.
df.to_excel(r'C:\\Users\\michael_lindsay\\OutputTest.xlsx', index = None, header = True)
print(df)
'''

#SETUP

counts = {}

def rolling_count(val):
    if val in counts:
        counts[val] += 1
    else:
        counts[val] = 1
    return

#JANUARY

jan1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-01-31')]
jan2 = df[(df['dateopened'] <= '2018-01-31') & (df['dateclosed'] > '2018-01-31')]
jan3 = df[(df['datereopened'] <= '2018-01-31') & (df['dateclosed'] > '2018-01-31')]
jan4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-01-31')]
print (jan4)

jan = jan1.append([jan2,jan3,jan4])
jan['Month'] = 'January' 
jan['Month'].apply(rolling_count)

print(jan)

#FEBRUARY

feb1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-02-28')]
feb2 = df[(df['dateopened'] <= '2018-02-28') & (df['dateclosed'] > '2018-02-28')]
feb3 = df[(df['datereopened'] <= '2018-02-28') & (df['dateclosed'] > '2018-02-28')]
feb4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-02-28')]


feb = feb1.append([feb2,feb3,feb4])
feb['Month'] = 'February' 
feb['Month'].apply(rolling_count)

#MARCH

mar1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-03-31')]
mar2 = df[(df['dateopened'] <= '2018-03-31') & (df['dateclosed'] > '2018-03-31')]
mar3 = df[(df['datereopened'] <= '2018-03-31') & (df['dateclosed'] > '2018-03-31')]
mar4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-03-31')]



mar = mar1.append([mar2,mar3,mar4])
mar['Month'] = 'March' 
mar['Month'].apply(rolling_count)

#APRIL

apr1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-04-30')]
apr2 = df[(df['dateopened'] <= '2018-04-30') & (df['dateclosed'] > '2018-04-30')]
apr3 = df[(df['datereopened'] <= '2018-04-30') & (df['dateclosed'] > '2018-04-30')]
apr4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-04-30')]

apr = apr1.append([apr2,apr3,apr4])
apr['Month'] = 'April' 
apr['Month'].apply(rolling_count)

#MAY

may1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-05-31')]
may2 = df[(df['dateopened'] <= '2018-05-31') & (df['dateclosed'] > '2018-05-31')]
may3 = df[(df['datereopened'] <= '2018-05-31') & (df['dateclosed'] > '2018-05-31')]
may4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-05-31')]

may = may1.append([may2,may3,may4])
may['Month'] = 'May' 
may['Month'].apply(rolling_count)

#JUNE

jun1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-06-30')]
jun2 = df[(df['dateopened'] <= '2018-06-30') & (df['dateclosed'] > '2018-06-30')]
jun3 = df[(df['datereopened'] <= '2018-06-30') & (df['dateclosed'] > '2018-06-30')]
jun4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-06-30')]

jun = jun1.append([jun2,jun3,jun4])
jun['Month'] = 'June' 
jun['Month'].apply(rolling_count)

#JULY 

jul1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-07-31')]
jul2 = df[(df['dateopened'] <= '2018-07-31') & (df['dateclosed'] > '2018-07-31')]
jul3 = df[(df['datereopened'] <= '2018-07-31') & (df['dateclosed'] > '2018-07-31')]
jul4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-07-31')]

jul = jul1.append([jul2,jul3,jul4])
jul['Month'] = 'July' 
jul['Month'].apply(rolling_count)


#AUGUST

aug1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-08-31')]
aug2 = df[(df['dateopened'] <= '2018-08-31') & (df['dateclosed'] > '2018-08-31')]
aug3 = df[(df['datereopened'] <= '2018-08-31') & (df['dateclosed'] > '2018-08-31')]
aug4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-08-31')]

aug = aug1.append([aug2,aug3,aug4])
aug['Month'] = 'August' 
aug['Month'].apply(rolling_count)

#SEPTEMBER

sep1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-09-30')]
sep2 = df[(df['dateopened'] <= '2018-09-30') & (df['dateclosed'] > '2018-09-30')]
sep3 = df[(df['datereopened'] <= '2018-09-30') & (df['dateclosed'] > '2018-09-30')]
sep4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-09-30')]

sep = sep1.append([sep2,sep3,sep4])
sep['Month'] = 'September' 
sep['Month'].apply(rolling_count)

#OCTOBER

octo1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-10-31')]
octo2 = df[(df['dateopened'] <= '2018-10-31') & (df['dateclosed'] > '2018-10-31')]
octo3 = df[(df['datereopened'] <= '2018-10-31') & (df['dateclosed'] > '2018-10-31')]
octo4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-10-31')]

octo = octo1.append([octo2,octo3,octo4])
octo['Month'] = 'October' 
octo['Month'].apply(rolling_count)

#NOVEMBER

nov1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-11-30')]
nov2 = df[(df['dateopened'] <= '2018-11-30') & (df['dateclosed'] > '2018-11-30')]
nov3 = df[(df['datereopened'] <= '2018-11-30') & (df['dateclosed'] > '2018-11-30')]
nov4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-11-30')]

nov = nov1.append([nov2,nov3,nov4])
nov['Month'] = 'November' 
nov['Month'].apply(rolling_count)

#DECEMBER

dec1 = df[(df['claimstatus'] == 'O') & (df['dateopened'] <= '2018-12-31')]
dec2 = df[(df['dateopened'] <= '2018-12-31') & (df['dateclosed'] > '2018-12-31')]
dec3 = df[(df['datereopened'] <= '2018-12-31') & (df['dateclosed'] > '2018-12-31')]
dec4 = df[(df['claimstatus'] == 'R') & (df['datereopened'] <= '2018-12-31')]

dec = dec1.append([dec2,dec3,dec4])
dec['Month'] = 'December' 
dec['Month'].apply(rolling_count)


openclaims = pd.DataFrame([counts], columns= ['January', 'February', 'March', 'April','May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'])
print(openclaims)
