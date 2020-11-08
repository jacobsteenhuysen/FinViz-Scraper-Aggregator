#!/usr/bin/python3

from finviz.helper_functions.save_data import export_to_db, export_to_csv
from finviz.screener import Screener
from finviz.main_func import *
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta, date


#filters = []
#filters = ['geo_usa']
filters = ['fa_div_pos']  # Shows companies in the S&P500
print("Filtering stocks..")
stock_list = Screener(filters=filters, order='ticker')
print("Parsing every stock..")
stock_list.get_ticker_details()


#df = pd.DataFrame(data=stock_list)

#print(df.head())

# Export the screener results to CSV file


stock_list.to_csv(r'C:/Users/Jacob Steenhuysen/Downloads/all_world_yields6.csv')

df =  pd.read_csv(r'C:/Users/Jacob Steenhuysen/Downloads/all_world_yields6.csv')



tickers_list = df['Ticker'].tolist()
data = pd.DataFrame(columns=tickers_list)

import yfinance as yf
#for ticker in tickers_list:
#    data[ticker] = yf.download(ticker, period="5d", interval="1d") ["Close"]

for ticker in tickers_list:
    data[ticker] = yf.download(ticker, start=datetime.now()-timedelta(days=366), end=date.today()) ["Adj Close"]


allWorldAnnualizedVol = data.apply(lambda x: x.pct_change().rolling(252).std()*(252**0.5))
                    
worldAnnualizedVol = allWorldAnnualizedVol.tail(1).transpose()

worldAnnualizedVol = worldAnnualizedVol.set_index(df.index)

df['Annualized Volatility'] = worldAnnualizedVol



sectors = df.groupby(df.Sector)

Basic_Materials = sectors.get_group("Basic Materials")
Communication_Services = sectors.get_group("Communication Services")
Consumer_Cyclical = sectors.get_group("Consumer Cyclical")
Consumer_Defensive = sectors.get_group("Consumer Defensive")
Energy = sectors.get_group("Energy")
Financials = sectors.get_group("Financial")
Healthcare = sectors.get_group("Healthcare")
Industrials = sectors.get_group("Industrials")
Real_Estate = sectors.get_group("Real Estate")
Tech = sectors.get_group("Technology")
Utilities = sectors.get_group("Utilities")



"""
to(Basic_Materials, file=r"C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx", sheetName="Basic Materials Tab")
write.xlsx(Communication_Services, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Communication Services", append=TRUE)
write.xlsx(Consumer_Cyclical, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Consumer Cyclical", append=TRUE)
write.xlsx(Consumer_Defensive, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Consumer Defensive", append=TRUE)
write.xlsx(Energy, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Energy", append=TRUE)
write.xlsx(Financials, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Financials", append=TRUE)
write.xlsx(Healthcare, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Healthcare", append=TRUE)
write.xlsx(Industrials, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Industrials", append=TRUE)
write.xlsx(Real_Estate, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Real_Estate", append=TRUE)
write.xlsx(Tech, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Tech", append=TRUE)
write.xlsx(Utilities, file=r'C:/Users/Jacob Steenhuysen/Downloads/FV_Test.xlsx', sheetName="Utilities", append=TRUE)
"""


with pd.ExcelWriter(r'C:/Users/Jacob Steenhuysen/Downloads/Income and Opportunity Model 9-16-2020.xlsx') as writer: 
    Basic_Materials.to_excel(writer, sheet_name="Basic Materials")
    Communication_Services.to_excel(writer, sheet_name="Communication Services")
    Consumer_Cyclical.to_excel(writer, sheet_name="Consumer Cyclical")
    Consumer_Defensive.to_excel(writer, sheet_name="Consumer Defensive")
    Energy.to_excel(writer, sheet_name="Energy")
    Financials.to_excel(writer, sheet_name="Financials")
    Healthcare.to_excel(writer, sheet_name="Healthcare")
    Industrials.to_excel(writer, sheet_name="Industrials")
    Real_Estate.to_excel(writer, sheet_name="Real_Estate")
    Tech.to_excel(writer, sheet_name="Tech")
    Utilities.to_excel(writer, sheet_name="Utilities")




