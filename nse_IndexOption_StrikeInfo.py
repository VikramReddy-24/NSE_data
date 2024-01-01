import requests
import traceback
import pandas as pd
from common_indexOption_strikeInfo import find_nearest_number,date_time_format,get_weekday,main_dataFrame_style
from datetime import datetime
import pytz
import os
from openpyxl import load_workbook



nofLatestExpiryDates=1
nofStrikes_toSelect=[1]

if nofLatestExpiryDates != len(nofStrikes_toSelect):
    print("In valid no of expiry dates and No of Strikes Prices")
    exit()


derivative_urls={
    
    "NIFTY":"https://www.nseindia.com/api/quote-derivative?symbol=NIFTY&identifier=OPTIDXNIFTY04-01-2024CE20500.00",
    "FINNIFTY":"https://www.nseindia.com/api/quote-derivative?symbol=FINNIFTY",
    "MIDCPNIFTY":"https://www.nseindia.com/api/quote-derivative?symbol=MIDCPNIFTY",
    "BANKNIFTY":"https://www.nseindia.com/api/quote-derivative?symbol=BANKNIFTY&identifier=OPTIDXBANKNIFTY03-01-2024CE45600.00",
    
}

equity_stock_urls={
   "NIFTY":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%2050",
   "FINNIFTY":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20FINANCIAL%20SERVICES",
   "MIDCPNIFTY":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20MIDCAP%20SELECT",
   "BANKNIFTY":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20BANK",

}



try:

    f = open("cookie.txt", "r")
    cookie=f.read()

    header={
    'Accept':'application/json, text/javascript, */*; q=0.01',
    'Accept-Encoding':'gzip, deflate, br',
    'Accept-Language':'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7,te;q=0.6',
    'Cookie':f"{cookie}",
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'X-Requested-With':'XMLHttpRequest'
    }
    
    # https://www.nseindia.com/api/quote-derivative?symbol=NIFTY&identifier=OPTIDXNIFTY04-01-2024CE20500.00
    # https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY


    for key in derivative_urls:
        try:

            last_dir=os.getcwd().strip(os.getcwd().split("\\")[-1])
            folder_path=f'{last_dir}derivatives/{key.lower()}'

            derivative_url_response=requests.get(url=derivative_urls[key],headers=header,timeout=20)
            print(derivative_url_response.status_code)
            print("derivative_url_response Status===>",derivative_url_response)

            equity_stock_url_response=requests.get(url=equity_stock_urls[key],headers=header,timeout=20)
            print(equity_stock_url_response.status_code)
            print("equity_stock_indices Status===>",equity_stock_url_response)

            if  derivative_url_response.status_code != 200 and equity_stock_url_response.status_code !=200:
                print("check the validity of url or cookie")
                exit() 

            derivative_url_response=derivative_url_response.json()
            equity_stock_url_response=equity_stock_url_response.json()

        except requests.exceptions.Timeout:
            print('The request timed out, Please change the cookie')
            exit()

        derivative_opt_timestamp=derivative_url_response["opt_timestamp"] 
        derivative_tsList=date_time_format(derivative_opt_timestamp)
        derivative_weekday=get_weekday(derivative_opt_timestamp)
        derivative_stocks=derivative_url_response["stocks"] 
        option_expiry_dates=derivative_url_response["expiryDatesByInstrument"]["Index Options"] 
        underlyingValue=derivative_url_response["underlyingValue"]
        
        equity_stock_timestamp=equity_stock_url_response["timestamp"]
        equity_stock_data= [i for i in equity_stock_url_response["data"] if i["priority"]==1][0]

        selected_equity_stock_data={
            "eq_ts":equity_stock_timestamp,
            "open": equity_stock_data["open"],
            "high": equity_stock_data["dayHigh"],
            "low": equity_stock_data["dayLow"],
            "lastPrice": equity_stock_data["lastPrice"],
            "change":equity_stock_data["change"],
            "%change":equity_stock_data["pChange"]

        }

        derivative_df=pd.DataFrame()
        for stock in derivative_stocks:
            derivative_metadata=stock["metadata"]
            derivative_tradeInfo=stock['marketDeptOrderBook']["tradeInfo"]
            derivative_otherInfo=stock['marketDeptOrderBook']["otherInfo"]
            derivative_carry={
                "volumeFreezeQuantity":stock["volumeFreezeQuantity"],
                "totalBuyQuantity":stock["marketDeptOrderBook"]["totalBuyQuantity"],
                "totalSellQuantity":stock["marketDeptOrderBook"]["totalSellQuantity"],
                "priceBestBuy":stock["marketDeptOrderBook"]["carryOfCost"]["price"]["bestBuy"],
                "priceBestSell":stock["marketDeptOrderBook"]["carryOfCost"]["price"]["bestSell"],
                "pricelastPrice":stock["marketDeptOrderBook"]["carryOfCost"]["price"]["lastPrice"],
                "carryBestBuy":stock["marketDeptOrderBook"]["carryOfCost"]["carry"]["bestBuy"],
                "carryBestSell":stock["marketDeptOrderBook"]["carryOfCost"]["carry"]["bestSell"],
                "carrylastPrice":stock["marketDeptOrderBook"]["carryOfCost"]["carry"]["lastPrice"],
                "derivative_ts":derivative_opt_timestamp
            }
            
            derivative={**derivative_metadata,**derivative_tradeInfo,**derivative_otherInfo,**derivative_carry}
            num_rows, num_columns = derivative_df.shape

            if num_rows == 0 or num_columns==0:
                derivative_df=pd.DataFrame(derivative,index=[0])
            else:    
                derivative_df=pd.concat([derivative_df,pd.DataFrame(derivative,index=[num_rows])])
                

        print(derivative_df.head(3))
       
        for i in range(0,nofLatestExpiryDates,1):
           expire_date=option_expiry_dates[i] 
           dummy_derivative_df=derivative_df.loc[(derivative_df.expiryDate==expire_date),:]
           dummy_derivative_df.head(3)
           strike_list=list(set(dummy_derivative_df["strikePrice"]))
           strike_list.sort()
           print("Strike List====>",strike_list)
           neartStrike=find_nearest_number(underlyingValue,strike_list)
           print("underLaying value ======>",underlyingValue)
           print("Nearest strike ====>",neartStrike)
           minStrikePrice=strike_list[strike_list.index(neartStrike)-nofStrikes_toSelect[i]] if strike_list.index(neartStrike)-nofStrikes_toSelect[i]>0 else strike_list[0]
           maxStrikePrice=strike_list[strike_list.index(neartStrike)+nofStrikes_toSelect[i]] if strike_list.index(neartStrike)+nofStrikes_toSelect[i]<len(strike_list) else strike_list[-1]
           print("selected values are from",minStrikePrice,"===========>",maxStrikePrice)

           dummy_derivative_df=dummy_derivative_df[(dummy_derivative_df['strikePrice']>=minStrikePrice) & (dummy_derivative_df['strikePrice']<=maxStrikePrice)]

           dummy_derivative_df.head(5)
           selected_strikeList=strike_list[strike_list.index(minStrikePrice):strike_list.index(maxStrikePrice)+1]
           print("Selected Strike price List===>",selected_strikeList)
           
           excel_name=f'{expire_date}.xlsx'

           if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                print(f'Folder at {folder_path} created successfully.')
                # Create an empty DataFrame (you can skip this step if you want an entirely empty Excel file)  
           else:
                print(f'Folder at {folder_path} already exists.') 

           print("file Path===>",f'{folder_path}/{excel_name}')

              

           for s in  selected_strikeList:
                print(f"for strike===> {s} and {key}")
                CE_df=dummy_derivative_df[(dummy_derivative_df['strikePrice']==s) & (dummy_derivative_df['optionType']=="Call")]
                PE_df=dummy_derivative_df[(dummy_derivative_df['strikePrice']==s) & (dummy_derivative_df['optionType']=="Put")]
                CE_dic=CE_df.to_dict('index')
                CE_dic=CE_dic[list(CE_dic.keys())[0]] if len(list(CE_dic.keys()))==1 else {}
                PE_dic=PE_df.to_dict('index')
                PE_dic=PE_dic[list(PE_dic.keys())[0]] if len(list(PE_dic.keys()))==1 else {}
                # print("ce==>",CE_dic)
                # print("pe==>",PE_dic)
                CEPE_dic={
                   
                      "date":f'{derivative_tsList[0]}/{derivative_tsList[1]}-{derivative_weekday}',
                    #   "day":f'{derivative_weekday}',
                      "time":f'{derivative_tsList[-1]}',


                      "%chg":round(selected_equity_stock_data["%change"],2),
                      "chg":selected_equity_stock_data["change"],
                      "ltp":selected_equity_stock_data["lastPrice"],
                      "high":selected_equity_stock_data["high"],
                      "low":selected_equity_stock_data["low"],
                      "hl":(selected_equity_stock_data["high"]-selected_equity_stock_data["low"]),
                      "%hl":round((selected_equity_stock_data["high"]-selected_equity_stock_data["low"])*100 / selected_equity_stock_data["low"],2),
                      "open":selected_equity_stock_data["open"],
                      
                      '%chg_ce':round(CE_dic.get('pChange',0),2),
                      'chg_ce':CE_dic.get('change',0),
                      "open_ce":CE_dic.get('openPrice',0),
                      "high_ce": CE_dic.get('highPrice',0),
                      "low_ce":CE_dic.get('lowPrice',0),
                      "ltp_ce":CE_dic.get('closePrice',0),

                      "strikeP":s,

                      "ltp_pe":PE_dic.get('closePrice',0),
                      "high_pe":PE_dic.get('highPrice',0),
                      "low_pe":PE_dic.get('lowPrice',0),
                      "open_pe":PE_dic.get('openPrice',0),
                      '%chg_pe':round(PE_dic.get('pChange',0),2),
                      'chg_pe':PE_dic.get('change',0), 

                      'p_close_ce':CE_dic.get('prevClose',0),

                      'qty_ce':CE_dic.get('numberOfContractsTraded',0),
                      'tto_ce':CE_dic.get('totalTurnove',0), 
                      'tradedVolume_ce':CE_dic.get('tradedVolume',0),
                      'value_ce':CE_dic.get('value',0),
                      'vmap_ce':CE_dic.get('vmap',0),
                      'premiumTurnover_ce':CE_dic.get('premiumTurnover',0),
                      'oi_ce':CE_dic.get('openInterest',0),
                      'coi_ce':f"{CE_dic.get('changeinOpenInterest',0)}",
                      '%coi_ce':f"{round(CE_dic.get('pchangeinOpenInterest',0),0)}%",
                      'Lot_ce':CE_dic.get('marketLot',0),
                      'sp_ce':CE_dic.get('settlementPrice',0),
                      'dvol_ce':CE_dic.get('dailyvolatility',0),
                      'annualisedVolatility_ce':CE_dic.get('annualisedVolatility',0),
                      'IV_ce':CE_dic.get('impliedVolatility',0),
                      'CWPL_ce':CE_dic.get('clientWisePositionLimits',0),
                      'MWPL_ce':CE_dic.get('marketWidePositionLimits',0),
                      "volumeFreezeQuantity_ce":CE_dic.get("volumeFreezeQuantity_ce",0),
                      "totalBuyQuantity_ce":CE_dic.get("totalBuyQuantity",0),
                      "totalSellQuantity_ce":CE_dic.get("totalSellQuantity",0),
                      "priceBestBuy_ce":CE_dic.get("priceBestBuy",0),
                      "priceBestSell_ce":CE_dic.get("priceBestSell",0),
                      "pricelastPrice_ce":CE_dic.get("pricelastPrice",0),
                      "carryBestBuy_ce":CE_dic.get("carryBestBuy",0),
                      "carryBestSell_ce":CE_dic.get("carryBestSell",0),
                      "carrylastPrice_ce":CE_dic.get("carrylastPrice",0),

                      'p_close_pe':PE_dic.get('prevClose',0),
                      
                     'qty_pe':PE_dic.get('numberOfContractsTraded',0),
                      'tto_pe':PE_dic.get('totalTurnove',0), 
                      'tradedVolume_pe':PE_dic.get('tradedVolume',0),
                      'value_pe':PE_dic.get('value',0),
                      'vmap_pe':PE_dic.get('vmap',0),
                      'premiumTurnover_pe':PE_dic.get('premiumTurnover',0),
                      'oi_pe':PE_dic.get('openInterest',0),
                      'coi_pe':PE_dic.get('changeinOpenInterest',0),
                      '%coi_pe':f"{round(PE_dic.get('pchangeinOpenInterest',0),0)}%",
                      'Lot_pe':PE_dic.get('marketLot',0),
                      'sp_pe':PE_dic.get('settlementPrice',0),
                      'dvolt_pe':PE_dic.get('dailyvolatility',0),
                      'annualisedVolatility_pe':PE_dic.get('annualisedVolatility',0),
                      'IV_pe':PE_dic.get('impliedVolatility',0),
                      'CWPL_pe':PE_dic.get('clientWisePositionLimits',0),
                      'MWPL_pe':PE_dic.get('marketWidePositionLimits',0),
                      "volumeFreezeQuantity_pe":PE_dic.get("volumeFreezeQuantity",0),
                      "totalBuyQuantity_pe":PE_dic.get("totalBuyQuantity",0),
                      "totalSellQuantity_pe":PE_dic.get("totalSellQuantity",0),
                      "priceBestBuy_pe":PE_dic.get("priceBestBuy",0),
                      "priceBestSell_pe":PE_dic.get("priceBestSell",0),
                      "pricelastPrice_pe":PE_dic.get("pricelastPrice",0),
                      "carryBestBuy_pe":PE_dic.get("carryBestBuy",0),
                      "carryBestSell_pe":PE_dic.get("carryBestSell",0),
                      "carrylastPrice_pe":PE_dic.get("carrylastPrice",0),

                      "index_ts":f"{equity_stock_timestamp}",
                      "derivative_ts":f"{derivative_opt_timestamp}",
                      "sys_ts":f"{datetime.now(pytz.timezone('Asia/Kolkata'))}"

               }
               
                CEPE_df=pd.DataFrame(CEPE_dic,index=[0])
                # print("keys=====>",CEPE_dic.keys())
                sheet_name=f'{s}.xlsx'
                
                existing_workbook_status=0
                try:
                    workbook=load_workbook(f'{folder_path}/{excel_name}')  
                    existing_workbook_status=1
                except FileNotFoundError:
                    print(f"FileNotFoundError in the Specified path Hence creating a new workbook {excel_name}.xlxs ")
                    existing_workbook_status=0
                    with pd.ExcelWriter(f'{folder_path}/{excel_name}', engine='xlsxwriter') as writer:
                        CEPE_df_style=main_dataFrame_style(CEPE_df)
                        CEPE_df_style.to_excel(writer,sheet_name=sheet_name,index=False)
                
                if existing_workbook_status==1:
                    with pd.ExcelWriter(f'{folder_path}/{excel_name}', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
                        sheets=writer.sheets.keys()
                        if sheet_name in sheets:     
                              existing_CEPE_df=pd.read_excel(f'{folder_path}/{excel_name}',sheet_name=sheet_name)
                              CEPE_df=pd.concat([CEPE_df,existing_CEPE_df]).reset_index(drop=True)
                        
                        CEPE_df_style=main_dataFrame_style(CEPE_df)
                        CEPE_df_style.to_excel(writer,sheet_name=sheet_name,index=False)
                        

                print(f"option index ==> {sheet_name} data successfully dumped into respective excel sheets{excel_name}")       
        print(f"{key} completed {'$'*100}")

except Exception as e:
    print(e)
    print("An Exception Occured")    
    traceback.print_exc()  
