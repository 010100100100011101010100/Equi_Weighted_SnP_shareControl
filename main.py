#flow-> Importing modules -> reading all the 505 stocks from the csv -> getting the data using api call(single then batch) -> parsing the data and creating a dataframe ->calculating number of shares to buy ->formatting our excel file ->saving and finish off

#importing the modules
import pandas as pd
import numpy as np
import math as m
import xlsxwriter as x
import requests
import json
from secrets import IEX_CLOUD_API_TOKEN


#reading the tickers from the csv file
stocks=pd.read_csv('sp_500_stocks.csv')


#our test api call
symbol='AAPL'
url=f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}'
data=requests.get(url).json()

#defining the columns in our dataframe
m=['Ticker','Stock Price','Market Cap','Number of shares to buy']

#defining chunks function to allow multiple data request from a single api call
def chunks(l,n):
    for i in range(0,len(l),n):
        yield l[i:i+n]

#creating the list of all tickers in chunks of 100 and creating a dataframe
symbol_groups=list(chunks(stocks,100))
symbol_string=[]
final_dataframe=pd.DataFrame(columns=m)
for i in range(0,len(symbol_groups)):
    symbol_string.append(','.join(symbol_groups[i]))
for s in symbol_string:
    batch_url=f'https://sandbox.iexapis.com/stable/stock/market/batch&symbols={s}?types=quote&token={IEX_CLOUD_API_TOKEN}'
    data=requests.get(batch_url).json()
    final_dataframe=final_dataframe.append(pd.Series([
        s,
        data[s]['quote']['latestPrice'],
        data[s]['quote']['marketCap']
        'N/A',index=m
    ]),ignore_index=True)

final_dataframe 

#now that we created the dataframe , we need to calculate the number of shares for each company based on the portfolio
#testing whether the portfolio value entered is numerical
portfolio=input()
try:
    val=float(portfolio)
except ValueError:
    print("please enter a numerical value")
    portfolio = input()

#creating the position size along with number of shares for each company
position_size=float(val/len(final_dataframe.index))
for i in range(len(final_dataframe.index)):
    final_dataframe.loc[i,'Number of shares to buy'] = m.floor(position_size/final_dataframe.loc[i,'Stock Price'])


#now that we have the number of shares added to our dataframe time to create the xlsx file , we will do that with pandas library itself
writer=pd.ExcelWriter('recommended trades',engine='xlsxwriter')
final_dataframe.to_excel(writer,sheet_name='recommended trades',index=False)

#we also initialize the bg and font color 
bg='#0a0a23'
font='#ffffff'


#we need format for each data entry in our excel file and we need three types

string=writer.book.add_format({
    'font_color':font,
    'bg_color':bg,
    'border':20
})

dollar=writer.book.add_format({
    'num_format':'$0.00',
    'font_color':font,
    'bg_color':bg,
    'border':20
})

integer=writer.book.add_format({
    'num_format':'0',
    'font_color':font,
    'bg_color':bg,
    'border':20
})

#now we need to apply the formats along with header adjustments our excel file
col_format={
    'A':['Ticker',string],
    'B':['Stock',dollar],
    'C':['Market Cap',dollar],
    'D':['Number of shares to buy',integer]
}
for c in col_format.keys():
    writer.sheets['recommended trades'].set_column(f'{c}:{c}',20,col_format[c][1])
    writer.sheets['recommended trades'].write(f'{c}1',col_format[c][0],string)

writer.save()




