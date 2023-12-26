import sys
import requests
import pandas as pd
import json
import os
import traceback
from datetime import date
from datetime import datetime
import xlsxwriter
import pytz
import colorsys

from botocore.exceptions import ClientError
from openpyxl import load_workbook
# from common_functions import find_nearest_number, get_current_datetime_weekday,upload_excel_to_s3,read_s3_file_data

def get_current_datetime_weekday():
    # Specify the Indian time zone
    indian_time_zone = 'Asia/Kolkata'

    # Get the current date and time in the specified time zone
    current_time = datetime.now(pytz.timezone(indian_time_zone))

    # Format the date as "20th Dec 2023"
    formatted_date = current_time.strftime("%dth %b %Y")

    # Format the time in 12-hour format
    formatted_time = current_time.strftime("%I:%M %p")

    # Get the day of the week (0 = Monday, 1 = Tuesday, ..., 6 = Sunday)
    day_of_week = current_time.strftime("%A")

    return formatted_date, formatted_time, day_of_week, f'{formatted_date}-{formatted_time}'

def get_weekday(recived_date):
    timestamp_str = recived_date

    # Define the timezone
    timezone_str = "Asia/Kolkata"

    # Convert the timestamp string to a datetime object
    timestamp_datetime = datetime.strptime(timestamp_str, "%d-%b-%Y %H:%M:%S")

    # Set the timezone for the datetime object
    timestamp_datetime = pytz.timezone(timezone_str).localize(timestamp_datetime)

    # Get the day of the week (Monday is 0 and Sunday is 6)
    day_of_week = timestamp_datetime.strftime("%A")

    return day_of_week[0:3]

def sheet_name_format(recived_data):
    # Define the timestamp in string format
    timestamp_str =recived_data

    # Define the timezone
    timezone_str = "Asia/Kolkata"

    # Convert the timestamp string to a datetime object
    timestamp_datetime = datetime.strptime(timestamp_str, "%d-%b-%Y %H:%M:%S")

    # Set the timezone for the datetime object
    timestamp_datetime = pytz.timezone(timezone_str).localize(timestamp_datetime)

    # Extract date, month, year, and time
    date = timestamp_datetime.strftime("%d")
    month = timestamp_datetime.strftime("%m")
    year = timestamp_datetime.strftime("%Y")
    time = timestamp_datetime.strftime("%H%M:%S")
    print(time)

    return f'{date}{month}-{time[0:4]}'

def gradient_color(val):
    global selected

    print(val)

    if val<=selected:
        brightness =  val/ selected
    elif val>selected:
        brightness =  selected/val 
    brightness=int(str(brightness*100).split('.')[0])
    # print(brightness)
    # if brightness==0:
    #     brightness=10
    # elif brightness==100:
    #     brightness=90  

    # HSL values
    h = 120  # Hue
    s = 100  # Saturation
    l = brightness   # Lightness

    # Convert HSL to RGB
    r, g, b = colorsys.hls_to_rgb(h / 360.0, l / 100.0, s / 100.0)

    # Convert RGB to integer values (0-255)
    r_int = int(r * 255)
    g_int = int(g * 255)
    b_int = int(b * 255)
    

    # Decrease brightness from top to bottom
    color = f'RGB({r_int}, {g_int}, {b_int})'  # Using HSL color representation
    print(f"background-color: {color}")
    return f'background-color: {color}'

def find_nearest_number(target, numbers):
    nearest_number = min(numbers, key=lambda x: abs(x - target))
    return nearest_number

selected=0

try:

    current_date = date.today()
    print("Current Date:", current_date)

    cookie='_ga=GA1.1.1416523229.1701883786; _ga_PJSKY6CFJH=GS1.1.1701890785.2.1.1701890785.60.0.0; nseQuoteSymbols=[{"symbol":"WIPRO","identifier":null,"type":"equity"},{"symbol":"BANKNIFTY","identifier":"FUTIDXBANKNIFTY28-12-2023XX0.00","type":"equity"},{"symbol":"FINNIFTY","identifier":"FUTIDXFINNIFTY26-12-2023XX0.00","type":"equity"},{"symbol":"MIDCPNIFTY","identifier":"FUTIDXMIDCPNIFTY22-12-2023XX0.00","type":"equity"},{"symbol":"NIFTY","identifier":"FUTIDXNIFTY28-12-2023XX0.00","type":"equity"}]; defaultLang=en; _ga_QJZ4447QD3=GS1.1.1703498103.28.0.1703498103.0.0.0; nsit=XNZwzHe2BRoGTmdpF5Oqn1Vu; nseappid=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJhcGkubnNlIiwiYXVkIjoiYXBpLm5zZSIsImlhdCI6MTcwMzUxOTE5MiwiZXhwIjoxNzAzNTI2MzkyfQ.lCxTzupjKMPd0Eyxgq7gw2X5sruy9yfMKKMQlIJWM8Y; AKA_A2=A; bm_mi=518EF789C75E98EB343B972E398A6279~YAAQcDLUF5IxIXmMAQAAoAaooRZ74yEAAp+8bEpAjgFmpBzdnvY6ob7JJ2pz2bthgoSRUUxwnqXenbRFc/HpyWschOQEWF/ZnSWjs9En4o8nioZCdWBE3eMzuUNZs1zSvRBafsF9Dluxwutod8b7T/7ybU03B8sNrsha0TEJJIk+bHwRYuH5W6p76B+q8OIFGaYmzsYeF2vVrLpU9ZOu6bYxdHD7n6mKjCwp+k9L+sbVIA5x1bGpYK/Ooa/kYkchAusUlK7r+d50cnM47NdmEuDf/PbPyBEuTeSNPQUaFEM3fMGrDJCQqDBih3iAgGfJx7yBmecyXZZIFbZkdjs3kJDK4265FAjV8PT+tjPQ~1; RT="z=1&dm=nseindia.com&si=19ab3b89-93b4-417c-84c4-0b273bf7ad6f&ss=lqkwdhmo&sl=0&se=8c&tt=0&bcn=%2F%2F684d0d42.akstat.io%2F"; ak_bmsc=EF6159F952EE306F5228C4C1850FB1E7~000000000000000000000000000000~YAAQcDLUF6wxIXmMAQAAHgqooRabFl3hXG2CYfvE943VmxRz6Qk45Z9v6auTv3RCoAtpnyh0gkkQNch1xOjvdM1EJTfyzBhPpfBho7dL0Ubj6ByPCfSLpXx1vRxNOH750q7TcNOM9dwQ0LeuD6fbWeBZLT/vKoxdAJZf3MzOMme3vO72nDrLEI+OgnJln0SogGHV83V0mQqikTZ9go3UTVZhuLf4YjFWjPijgsEOeoRzqRN6ZTYIz7+41zER/fonwwypVII2von3LkUZQc5Pf6RN66TLbR2HWj19XsAJ0JiF5KaoT/Ry+GXKG7gR2Z4B61crPB7zNnxzGzL6I2iCgNqNy2HpIjUe3QyXPhWvpNiqPbHawk915malfAKf6gGm9uPQt0ILsY8tk5w7jrWhXGDSHHbbuy8ARG3ESSy/9fMwTUQoh4if6ctgf0AoO5Luk4fvsWf8kXIC4skWo/Od1bu8QG1JxXZpEo8EPR/FoYzwex2t0Qqm8poLBq8R2mpD5t79kkBVcmZeM43uXACoIURtFH8/nx1JgScBTHjHBRCJ15xyTB2TGl3lH7v/7ZLsel6httVCI+Pp3jY=; _ga_87M7PJ3R97=GS1.1.1703519192.61.1.1703519193.0.0.0; bm_sv=2A713EEB9664961624720D311B8B7EE0~YAAQcDLUF8UxIXmMAQAAOgyooRYLROAnAQ3AZxluULP0Wy8rhXCozdrss9DD7x2lIyZaGhbA6vfNKmNyqWrSpJENDxyfasSvKb7fgjjph07EPVS43jQwEjOhEF64i2wHx7AO7fxXj2EAA0WEdL8gJt38/cnKoaw96chVgIBuIx6MKrjRp6+HGq35h/+0wKkdp955KsdO0Q2S1eYntmGaBJ5XHMTma6caWXfrhN0cc8i3BOIpJiZHgVQkt3vNyMGgpTm5Vw==~1'
    #####################################################

    header={
    'Accept':'application/json, text/javascript, */*; q=0.01',
    'Accept-Encoding':'gzip, deflate, br',
    'Accept-Language':'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7,te;q=0.6',
    'Cookie':f"{cookie}",
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'X-Requested-With':'XMLHttpRequest'
    }

    urls={
            "OPTIDXNIFTY":'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY',
            "OPTIDXBANKNIFTY":"https://www.nseindia.com/api/option-chain-indices?symbol=BANKNIFTY",
            "OPTIDXFINNIFTY":"https://www.nseindia.com/api/option-chain-indices?symbol=FINNIFTY",
            "OPTIDXMIDCP":"https://www.nseindia.com/api/option-chain-indices?symbol=MIDCPNIFTY",
          }
    
    nifty_urls={
         "OPTIDXNIFTY":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%2050",
         "OPTIDXBANKNIFTY":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20BANK",
         "OPTIDXFINNIFTY":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20FINANCIAL%20SERVICES",
         "OPTIDXMIDCP":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20MIDCAP%20SELECT"
    }
   

    
    for key in urls:
        print("$"*100)
        print(f"Fetching {key} data")

        nifty_url=nifty_urls[key]
        nifty_response=requests.get(url=nifty_url,headers=header)
        print("Response===>",nifty_response)
        # today_date=nifty_response.json()["timestamp"].split(" ")[0]
        nifty_data=nifty_response.json()['data']
        timestamp=nifty_response.json()['timestamp']
        calculated_data={}
        calculated_data['Date']=f'{timestamp[0:6]}-{get_weekday(timestamp)}'
        calculated_data['ts']=str(timestamp).split(" ")[1][0:5]
    
        for stock in nifty_data:
            if stock["priority"]==1:
                del stock["priority"]
                del stock["identifier"]
                del stock["lastUpdateTime"]
                del stock["chart365dPath"]
                del stock["chartTodayPath"]
                del stock["chart30dPath"]
                del stock["date30dAgo"]  
                del stock["symbol"]
                del stock["ffmc"]
                
                stock['totalTradedVolume(Lacs)']=round(stock['totalTradedVolume']/100000,3)
                stock['totalTradedValue(Crs)']=round(stock['totalTradedValue']/10000000,3)
                stock['nearWKH']=round(stock['nearWKH'],2)
                stock['nearWKL']=round(stock['nearWKL'],2)
                del stock['totalTradedVolume']
                del stock['totalTradedValue']
                underlaying_value=stock['lastPrice']

                calculated_data['chg']=stock['change']
                calculated_data['%chg']=stock['pChange']
                calculated_data['high']=stock["dayHigh"]
                calculated_data['low']=stock["dayLow"]
                calculated_data['close']=stock["lastPrice"]
                calculated_data['open']=stock["open"]

                nifty_stock=stock
                break

        url=urls[key]
        response=requests.get(url=url,headers=header)
        print("Response===>",response)
        expiry_date=response.json()["records"]["expiryDates"][0]
        data=response.json()[ "records"]["data"]
        
        PECE_df=pd.DataFrame()
        for stock in data:
            PE_CE=[]
            if stock["expiryDate"]==expiry_date:
                PE=stock.get("PE",{})
                CE=stock.get("CE",{})
                strikePrice=stock.get("strikePrice"," ")
                dummy={}

                dummy=nifty_stock
                
                dummy['OI_ce']=int(CE.get("openInterest",0))
                dummy['CHNG IN OI_ce']=int(CE.get("changeinOpenInterest",0))
                dummy['VOLUME_ce']=int(CE.get("totalTradedVolume",0))
                dummy['IV_ce']=int(CE.get("impliedVolatility",0)) 
                dummy['LTP_ce']=int(CE.get("lastPrice",0))
                dummy['CHNG_ce']=int(CE.get("change",0))
                dummy['BID QTY_ce']=int(CE.get("bidQty",0))
                dummy['BID_ce']=int(CE.get("bidprice",0))
                dummy['ASK_ce']=int(CE.get("askPrice",0))
                dummy['ASK QTY_ce']=int(CE.get("askQty",0))
                dummy['TBQ_ce']=int(CE.get("totalBuyQuantity",0))
                dummy['TSQ_ce']=int(CE.get("totalSellQuantity",0))
        
                dummy['STRIKE_PRICE']=strikePrice
                dummy['Expiry_date']=expiry_date

                dummy['OI_pe']=int(PE.get("openInterest",0))
                dummy['CHNG IN OI_pe']=int(PE.get("changeinOpenInterest",0))
                dummy['VOLUME_pe']=int(PE.get("totalTradedVolume",0))
                dummy['IV_pe']=int(PE.get("impliedVolatility",0)) 
                dummy['LTP_pe']=int(PE.get("lastPrice",0))
                dummy['CHNG_pe']=int(PE.get("change",0))
                dummy['BID QTY_pe']=int(PE.get("bidQty",0))
                dummy['BID_pe']=int(PE.get("bidprice",0))
                dummy['ASK_pe']=int(PE.get("askPrice",0))
                dummy['ASK QTY_pe']=int(PE.get("askQty",0))
                dummy['TBQ_pe']=int(PE.get("totalBuyQuantity",0))
                dummy['TSQ_pe']=int(PE.get("totalSellQuantity",0))
               


                PE_CE.append(dummy)
                num_rows, num_columns = PECE_df.shape
                if num_rows == 0 or num_columns==0:
                   PECE_df=pd.DataFrame(PE_CE)
                else:
                   PECE_df = pd.concat([PECE_df,pd.DataFrame(PE_CE)],ignore_index=True)

        calculated_data["qty[c/p]"]= round(PECE_df['VOLUME_ce'].sum() / PECE_df['VOLUME_pe'].sum() , 2)
        calculated_data["oi[c/p]"]=round(PECE_df['OI_ce'].sum()/PECE_df['OI_pe'].sum() ,2)
        calculated_data["coi[c/p]"]=round(PECE_df['CHNG IN OI_ce'].sum()/PECE_df['CHNG IN OI_pe'].sum(),2)

        calculated_data["qty[p/c]"]= round(PECE_df['VOLUME_pe'].sum()  / PECE_df['VOLUME_ce'].sum() ,2) 
        calculated_data["oi[p/c]"]=round(PECE_df["OI_pe"].sum()/PECE_df["OI_ce"].sum() ,2)
        calculated_data["coi[p/c]"]=round(PECE_df['CHNG IN OI_pe'].sum()/PECE_df['CHNG IN OI_ce'].sum() ,2)

        calculated_data['oi_c']=PECE_df['OI_ce'].sum()
        calculated_data['qty_c']=PECE_df['VOLUME_ce'].sum()
        calculated_data['oi_cs']=PECE_df['OI_ce'].sum()
        calculated_data['qty_cs']=PECE_df['VOLUME_ce'].sum()
        calculated_data['coi_c']=PECE_df['CHNG IN OI_ce'].sum()
        calculated_data['iv_c']=PECE_df['IV_ce'].sum()
        calculated_data['ltp_c']=PECE_df['LTP_ce'].sum()
        calculated_data['chng_c']=PECE_df['CHNG_ce'].sum()

        calculated_data['oi_p']=PECE_df['OI_pe'].sum()
        calculated_data['coi_p']=PECE_df['CHNG IN OI_pe'].sum()
        calculated_data['qty_p']=PECE_df['VOLUME_pe'].sum()
        calculated_data['iv_p']=PECE_df['LTP_pe'].sum()
        calculated_data['ltp_p']=PECE_df['LTP_pe'].sum()
        calculated_data['chng_c']=PECE_df['CHNG_ce'].sum()
        calculated_data['sy_ts']=get_current_datetime_weekday()[3]
        

        

        print("calculated_data==>",calculated_data)

        
        # print(calculated_data_df.head())
        
        # Define a color code (e.g., light green)
        # color_code = '#5dd55d'

        
        # selected=find_nearest_number(underlaying_value,list(PECE_df['STRIKE_PRICE']))
        # Apply color coding to a specific column (e.g., 'Age')
        # PECE_df = PECE_df.style.applymap(lambda x:f'background-color: {color_code}', subset=['STRIKE_PRICE'])
        # PECE_df.style.background_gradient(cmap='viridis',subset=['STRIKE_PRICE'])

        PECE_df =PECE_df.style.set_properties(subset = ['STRIKE_PRICE'],
                        **{"background-color": "#047E8E",  
                           
                           "border" : "1px solid white"})  \
                        .set_properties(subset = ['lastPrice'],
                        **{"background-color": "#FFE177",  
                           
                           "border" : "1px solid white"}) \
                        .set_properties(subset = ['dayHigh'],
                        **{"background-color": "#A4C24F",  
                           
                           "border" : "1px solid white"}) \
                        .set_properties(subset = ['dayLow'],
                        **{"background-color": "#DA7A3D",  
                           
                           "border" : "1px solid white"})    
                        
          
        
        existing_excel_file_path=f'{os.getcwd()}/index_option_data'
        excel_name=f'{key}.xlsx'
        sheet_name=f'{sheet_name_format(timestamp)}.xlsx'
        calculated_sheet_name = f'cp.xlsx'
        print(sheet_name)

        # Create the folder if it doesn't exist
        if not os.path.exists(existing_excel_file_path):
            os.makedirs(existing_excel_file_path)
            print(f'Folder at {existing_excel_file_path} created successfully.')
            # Create an empty DataFrame (you can skip this step if you want an entirely empty Excel file)  

        else:
            print(f'Folder at {existing_excel_file_path} already exists.')    
    
        existing_calculated_df_status=0   
        try:
            # Open the Excel workbook
            workbook=load_workbook(f'{existing_excel_file_path}/{excel_name}')
            existing_calculated_df_status=1
            # worksheet_name = f'{key}.xlsx'
            # Convert the worksheet to a Pandas DataFrame
            # existing_calculated_df = pd.read_excel(f'{existing_excel_file_path}/{excel_name}', sheet_name=worksheet_name, header=None)
        except FileNotFoundError:
            print(f"FileNotFoundError in the Specified path Hence creating a new File {excel_name} and appending {sheet_name} and {key}.xlxs ")
            existing_calculated_df_status=0
            calculated_data_df=pd.DataFrame([calculated_data])
            with pd.ExcelWriter(f'{existing_excel_file_path}/{excel_name}', engine='xlsxwriter') as writer:
                PECE_df.to_excel(writer, sheet_name= sheet_name, index=False)
                calculated_data_df.to_excel(writer, sheet_name= calculated_sheet_name, index=False)
            
        if existing_calculated_df_status==1:
        #    worksheet = workbook[calculated_sheet_name]
           # Convert the worksheet to a Pandas DataFrame
           existing_calculated_df = pd.read_excel(f'{existing_excel_file_path}/{excel_name}', sheet_name=calculated_sheet_name)
        #    print(existing_calculated_df.shape)
           first_row=existing_calculated_df.iloc[0].to_dict()
           calculated_data['oi_cs']=calculated_data['oi_cs']-first_row['oi_cs']
           calculated_data['qty_cs']=calculated_data['qty_cs']-first_row['qty_cs']
           calculated_data_df=pd.DataFrame([calculated_data])
           calculated_data_df = pd.concat([calculated_data_df,existing_calculated_df])
        #    print(calculated_data_df.head(0))
           with pd.ExcelWriter(f'{existing_excel_file_path}/{excel_name}', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:
                PECE_df.to_excel(writer, sheet_name= sheet_name, index=False)
                calculated_data_df.to_excel(writer, sheet_name= calculated_sheet_name, index=False)
               
        print(f"option index ==> {key} data successfully dumped into excel")       
     
        
except Exception as e:
    print(e)
    print("An Exception Occured")    
    traceback.print_exc()  


