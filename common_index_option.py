from datetime import datetime
import xlsxwriter
import pytz
import traceback
import pandas as pd
import pywhatkit as kit



def find_nearest_number(target, numbers):
    nearest_number = min(numbers, key=lambda x: abs(x - target))
    return nearest_number

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

def main_dataFrame_style(df):
    style_df=df.style.set_properties(subset = ['STRIKE_PRICE'],
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
                           
                           "border" : "1px solid white"}) \
                          
                        
    return style_df

def cp_dataFrame_style(df):
    try:
        print(type(df))
        min_max_cols=["open","low","high","close"]
        style_df=df.style\
                         .highlight_min(subset=min_max_cols,color="#ff3333")\
                         .highlight_max(subset=min_max_cols,color="Lime")  
    except Exception as e:
        print(e)
        print("An Exception Occured")    
        traceback.print_exc()      
     
                        
    return style_df

def send_whatsapp_message():
    phone_number='+919963435360'
    message="Bro, Added the common file now you can pull the latest code"
    kit.sendwhatmsg_instantly(phone_no=phone_number,message=message,tab_close=True)