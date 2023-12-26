import requests
import pandas as pd
import os
import traceback
from datetime import date
import json

current_date = date.today()

print("Current Date:", current_date)

try:

    cookie='_ga=GA1.1.1416523229.1701883786; _ga_PJSKY6CFJH=GS1.1.1701890785.2.1.1701890785.60.0.0; _ga_QJZ4447QD3=GS1.1.1702884292.15.0.1702884292.0.0.0; nsit=2r4D_sSA5agGTfCs1rGlgPxt; nseappid=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJhcGkubnNlIiwiYXVkIjoiYXBpLm5zZSIsImlhdCI6MTcwMjk5MzUyNiwiZXhwIjoxNzAzMDAwNzI2fQ.t_o1A6jhQs-kXyuLU-27xX0zjyqLLbKEUYppuNBG3FI; defaultLang=en; _ga_87M7PJ3R97=GS1.1.1702993527.37.0.1702993527.0.0.0; ak_bmsc=35FCC96457E5019630A77F897D5B94EC~000000000000000000000000000000~YAAQpXEsMTr3S1qMAQAACghTghb/1Sh+tDTN+zqYQAQCUjUuTG2xVe0cYL0nVns/CcI+hgggVPE5osvtBbiRt8UTEMAIfRzLx6YL8Z0pd/6ddLIP5rbRpobOOchRucFCP6fiZwXbHMbg4LT5x3x2egXifzBE/FEUTnnWuuB2xUdgBfz03noBUWmRzDChK5ESr8E3cbUFcMw+0KuA55dmTOZ+xh8GuzBBa99HOQ4xkZMmMUCzyyA0ffDdEcyYnu8deyrY9KUrBLE672qQZu+3h/ue110YGQQskF6qDk57FRigKPkqLDb9eUURf005seZ6q/b0GEw/s2mXnQjDr8J5LhBpTqc20aHANe3sA2NBwrLGUIsZY96Oe905XnI3fPHkhxZv4Bayi10DYX8psbC8t4/QodHWI+WDKQqNJxyHCv6RtRX2b/BiChMtTmEy22EBKqWcb7G/eTVqYmSX9+294hg8YCH587+zHiGItrYGXqjGW8qBjigVU9tT2p8U2qGb3Q==; RT="z=1&dm=nseindia.com&si=19ab3b89-93b4-417c-84c4-0b273bf7ad6f&ss=lqcedhff&sl=1&se=8c&tt=23n&bcn=%2F%2F684d0d47.akstat.io%2F&ld=4dd"; bm_sv=40BADBFB4B37B60927B8AB1BA4D30763~YAAQpXEsMez6S1qMAQAAyzZTghZz1UgMFezxNIMpvKOay9EUjWIUFNaEpDqJB16UOpsV2hFQWF70rcFtVPcuvhTPQN9mHzzcXkYyNeJtmOU37y4rC1qxSa8tzcw2jjxF7fgddnmt+9pBixAEF5EgwYoft5rg9AThWhG3dhyPktbo1WztUEncOgpml96knqlvT3hy2n8Ry2dib+hU0N4Z+Lic8I+59urslJRm8lhO6L2gQTXL2TnRFcNCcjPEP116Osc=~1'
     #####################################################

    header={
    'Accept':'application/json, text/javascript, */*; q=0.01',
    'Accept-Encoding':'gzip, deflate, br',
    'Accept-Language':'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7,te;q=0.6',
    'Cookie':f"{cookie}",
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'X-Requested-With':'XMLHttpRequest'
    }

    

    # Open the JSON file in read mode
    with open('stocknames.json', 'r') as file:
        stocks = json.load(file)
        print(stocks)

    for stock in stocks["stocks"]:
        stockName=stock

        try:
            url=f'https://www.nseindia.com/api/quote-equity?symbol={stockName}'
            response=requests.get(url=url,headers=header)
            print("Quote Equity Status",response)
            priceInfo=response.json()['priceInfo']

            url=f"https://www.nseindia.com/api/quote-equity?symbol={stockName}&section=trade_info"
            response=requests.get(url=url,headers=header)
            print("Trade Info Status",response)

            url=f"https://www.nseindia.com/api/quote-derivative?symbol={stockName}"
            derivative_response=requests.get(url=url,headers=header)
            print("quote derivative Status",derivative_response)

            tradeInfo=response.json()["marketDeptOrderBook"]["tradeInfo"]
            quantityDetails=response.json()["marketDeptOrderBook"]
            valueAtRisk=response.json()["marketDeptOrderBook"]['valueAtRisk']
            securityWiseDP=response.json()["securityWiseDP"]

            derivative_metadata=derivative_response.json()["stocks"][0]["metadata"]
            derivative_marketDeptOrderBook=derivative_response.json()["stocks"][0]['marketDeptOrderBook']
            # print("derivative_metadata:",derivative_metadata)
        #     "securityWiseDP": {
        #     "quantityTraded": 9470239,
        #     "deliveryQuantity": 4780117,
        #     "deliveryToTradedQuantity": 50.48,
        #     "seriesRemarks": null,
        #     "secWiseDelPosDate": "08-DEC-2023 EOD"
        # }
            # print("price info:",priceInfo)

        except requests.RequestException as e:
            print(f"Request Exception: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

        data = dict()
        
        data['DATE']=[current_date] 
        data['change']=[priceInfo['change']]
        data['pChange']=[priceInfo['pChange']]
        data['LAST PRICE']=[priceInfo['lastPrice']]
        data['PREV. CLOSE']=[priceInfo['previousClose']]
        data['OPEN']=[priceInfo['open']]
        data['HIGH']=[priceInfo['intraDayHighLow']['min']]
        data['LOW']=[priceInfo['intraDayHighLow']['max']]
        data['CLOSE*']=[priceInfo['close']]
        data['VWAP']=[priceInfo['vwap']]

        data['52 Week High']=[priceInfo['weekHighLow']['max']]
        data['52 Week Low']=[priceInfo['weekHighLow']['min']]
        data['52 Week High date']=[priceInfo['weekHighLow']['maxDate']]
        data['52 Week Low data']=[priceInfo['weekHighLow']['minDate']]
        data['Upper Band']=[priceInfo['upperCP']]
        data['Lower Band']=[priceInfo['lowerCP']]
        data['Price Band']=[priceInfo['pPriceBand']]

        data['Buy Quantity']=[quantityDetails['totalBuyQuantity']]
        data['Sell Quantity']=[quantityDetails["totalSellQuantity"]]

        data['Traded Volume (Shares)']=[tradeInfo['totalTradedVolume']]
        data['Traded Value (₹ Lakhs)']=[tradeInfo['totalTradedValue']]
        data['Total Market Cap (₹ Lakhs)']=[tradeInfo['totalMarketCap']]
        data['Free Float Market Cap (₹ Lakhs)']=[tradeInfo['ffmc']]
        data['Impact cost']=[tradeInfo['impactCost']]
        data['Daily Volatility']=[tradeInfo['cmDailyVolatility']]
        data['Annualised Volatility']=[tradeInfo['cmAnnualVolatility']]

        data['Security VaR']=[valueAtRisk['securityVar']]
        data['Index VaR']=[valueAtRisk['indexVar']]
        data['VaR Margin']=[valueAtRisk['varMargin']]
        data['Extreme Loss Rate']=[valueAtRisk['extremeLossMargin']]
        data['Adhoc Margin']=[valueAtRisk['adhocMargin']]
        data['Applicable Margin Rate']=[valueAtRisk['applicableMargin']]

        data['Quantity Traded']=[securityWiseDP["quantityTraded"]]
        data['Deliverable Quantity (gross across client level)']=[securityWiseDP["deliveryQuantity"]]
        data[" percent of Deliverable Quantity to Traded Quantity"]=[securityWiseDP["deliveryToTradedQuantity"]]


        data['INSTRUMENT TYPE']=[derivative_metadata['instrumentType']]
        data['EXPIRY DATE']=[derivative_metadata['expiryDate']]
        data['OPEN']=[derivative_metadata['openPrice']]
        data['HIGH']=[derivative_metadata['highPrice']]
        data['LOW']=[derivative_metadata['lowPrice']]
        data['CLOSE']=[derivative_metadata['closePrice']]
        data['PREV.CLOSE']=[derivative_metadata['prevClose']]
        data['LAST']=[derivative_metadata['lastPrice']]
        data['CHANGE']=[derivative_metadata['change']]
        data['%CHANGE']=[derivative_metadata['pChange']]
        data['Volume(contracts)']=[derivative_metadata['numberOfContractsTraded']]
        data['Value(Lakhs)']=[derivative_metadata['totalTurnover']]

        data['Buy(Total Quantity)']=[derivative_marketDeptOrderBook['totalBuyQuantity']]
        data['Sell(Total Quantity)']=[derivative_marketDeptOrderBook['totalSellQuantity']]

        data['Traded Volume (Contracts)']=[derivative_marketDeptOrderBook['tradeInfo']['tradedVolume']]
        data['Traded Value (₹ Lakhs )']=[derivative_marketDeptOrderBook['tradeInfo']["value"]]
        data['VWAP']=[derivative_marketDeptOrderBook['tradeInfo']["vmap"]]
        # data['Underlying Value']=[derivative_marketDeptOrderBook['tradeInfo'][]]
        data['Market Lot']=[derivative_marketDeptOrderBook['tradeInfo']["marketLot"]]
        data['Open Interest (Contracts)']=[derivative_marketDeptOrderBook['tradeInfo']["openInterest"]]
        data['Change in Open Interest']=[derivative_marketDeptOrderBook['tradeInfo']["changeinOpenInterest"]]
        data['% Change in Open Interest']=[derivative_marketDeptOrderBook['tradeInfo']["pchangeinOpenInterest"]]

        data['Daily Volatility']=[derivative_marketDeptOrderBook["otherInfo"]["dailyvolatility"]]
        data['Annualised Volatility']=[derivative_marketDeptOrderBook["otherInfo"]["annualisedVolatility"]]
        data['Implied Volatility']=[derivative_marketDeptOrderBook["otherInfo"]["impliedVolatility"]]
        data['Settlement Price']=[derivative_marketDeptOrderBook["otherInfo"]["settlementPrice"]]
        data['Client Wise Position Limits']=[derivative_marketDeptOrderBook["otherInfo"]["clientWisePositionLimits"]]
        data['Market Wide Position Limits']=[derivative_marketDeptOrderBook["otherInfo"]["marketWidePositionLimits"]]

        data['carryOfCost price']=["PRICE"]
        data["bestBuy"]=[derivative_marketDeptOrderBook["carryOfCost"]["price"]["bestBuy"]]
        data["bestSell"]=[derivative_marketDeptOrderBook["carryOfCost"]["price"]["bestSell"]]
        data["lastPrice"]=[derivative_marketDeptOrderBook["carryOfCost"]["price"]["lastPrice"]]
        data['carryOfCost carry']=["CARRY"]
        data["bestBuy_Carryofcost"]=[derivative_marketDeptOrderBook["carryOfCost"]["carry"]["bestBuy"]]
        data["bestSell_Carryofcost"]=[derivative_marketDeptOrderBook["carryOfCost"]["carry"]["bestSell"]]
        data["lastPrice_Carryofcost"]=[derivative_marketDeptOrderBook["carryOfCost"]["carry"]["lastPrice"]]


        # print("data:",data)
        excel_file_path = f'excel/{stockName}.xlsx'
        current_directory = os.getcwd()
        file_path = os.path.join(current_directory, excel_file_path )

        # Load the existing data from the Excel file
        try:
            existing_df = pd.read_excel(file_path)
            print(existing_df)
            insert_index = len(existing_df)
            # Use loc to insert the new row at the specified index
            data_dict={}
            for key, value in data.items():
                data_dict[key]=value[0]
                existing_df.loc[insert_index] = data_dict
                existing_df.to_excel(excel_file_path, index=False)
        except FileNotFoundError:
            # If the file doesn't exist, create a new DataFrame
            print("Excel File file Path:", file_path)
            data_frame = pd.DataFrame.from_dict(data)

            # Save the DataFrame to Excel
            data_frame.to_excel(excel_file_path, index=False)

        

        print(f'{stock} successfully written to Excel file: {excel_file_path}')

except Exception as e:
    print(e)
    print("An Exception Occured")    
    traceback.print_exc()