import requests
import pandas as pd
import os
import traceback
from datetime import date
import json

current_date = date.today()
print("Current Date:", current_date)

try:

    cookie='_ga=GA1.1.1416523229.1701883786; _ga_PJSKY6CFJH=GS1.1.1701890785.2.1.1701890785.60.0.0; _ga_QJZ4447QD3=GS1.1.1703146097.20.1.1703146191.0.0.0; defaultLang=en; nseQuoteSymbols=[{"symbol":"WIPRO","identifier":null,"type":"equity"}]; nsit=hhyc8H9Bit0gsj4xj1xecW_4; nseappid=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJhcGkubnNlIiwiYXVkIjoiYXBpLm5zZSIsImlhdCI6MTcwMzMxMTkwNywiZXhwIjoxNzAzMzE5MTA3fQ.6d_Fzjc1GydYFFaNY4LwgSa_Jty_y-1yyWCyXzpMICE; AKA_A2=A; _ga_87M7PJ3R97=GS1.1.1703311906.47.1.1703311908.0.0.0; ak_bmsc=0C89A633266232D73B2206C8AFA1D72E~000000000000000000000000000000~YAAQcDLUF2SX93iMAQAA/yRNlRb9Paowour6kNBak6EYkeEkXDi4LUN0EHxJSQ8+4qvdCE6ryuqRisGBL/vQtOQtI6rSHBOvhS4S05HvpwY0efBfVCDTdE78HvwFNr/SC3hCarCBTvpbSV03Epe6v8m7Kd56SBvFP15NV9OjZ8P4WpLD4NVyJT45S+h/rPqLPJG2v/nWqfVTShx8Hyhq7XhUB3f6zHX8abF2Bn6ctTaUBZXak2m9a9frw2S2qgE3vzPOqqVJb4+43yuBoErS2IoWEnr5p+TB+lRvvu+foKwvKqwnBdhY1PoVsODtKDesusa/ND/zrvYoMq2nwB3tbGsGKs0BRdlqfOnbHdZcepca4sxGCBxiz2DifYiQv50GWmMUNhVvBrENKddAZ5H8gthABfAdDFTS3/WpTt1EqYC37XMqofSeN5bPjsXW14KcUHHk2/2Up+c+RmR1yoWmhpj6dYPf8No2iDXSvc6GlNqgv0h0VEViL/wiDr8dRVaD5qX2; RT="z=1&dm=nseindia.com&si=19ab3b89-93b4-417c-84c4-0b273bf7ad6f&ss=lqgv0cl8&sl=0&se=8c&tt=0&bcn=%2F%2F684d0d49.akstat.io%2F"; bm_sv=FA4037D60FF1528ADBF44A7950C03E32~YAAQcDLUF5ma93iMAQAAFA9OlRZXEvgq3L4TrGm9rf/Bhdn9niaWUnTsgsPi0SE2iiyCqCGByGSV+GqCojRKxrevIRXycZKkYLaF/XoQb5jJ/AIMDc1WGtzGfFklWgH3BjnuXyTb/06DXJHdVOkQk0V4Apq27+v2lh8YY1r2M7UGJpYqmOjKBZXR9wP2jlZwn9lWBbAn+KjGFsvx6l56ZHwrfGV5lpiWgusgeS9tVQUZNm/vnv6p0xMgvmxIwRMu9kE=~1'
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
        "NIFTY NEXT 50":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20NEXT%2050",
        "NIFTY 50":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%2050",
        "NIFTY MIDCAP 50":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20MIDCAP%2050",
        "NIFTY MIDCAP 100":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20MIDCAP%20100",
        "NIFTY MIDCAP 150":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20MIDCAP%20150",
        "NIFTY SMALLCAP 50":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20SMALLCAP%2050",
        "NIFTY BANK":"https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20BANK"

    }

    for key in urls:
        url=urls[key]
        response=requests.get(url=url,headers=header)
        today_date=response.json()["timestamp"].split(" ")[0]
        data=response.json()['data']
        for stock in data:
            if stock["priority"]!=1:
               del stock["meta"]

            del stock["priority"]
            del stock["identifier"]
            del stock["lastUpdateTime"]
            del stock["chart365dPath"]
            del stock["chartTodayPath"]
            del stock["chart30dPath"]
            del stock["date30dAgo"]   

        folder_path = f'{os.getcwd()}/index_data/{key}'

        print("Current Working Directory:", folder_path)

        # Create the folder if it doesn't exist
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f'Folder at {folder_path} created successfully.')
        else:
            print(f'Folder at {folder_path} already exists.')    
        

        data=pd.DataFrame(data)  
        data.to_excel(f"{folder_path}/{today_date}.xlsx", index=False)      
        print(f"{today_date}Data of {key} Successfully dumped into Excel")
        print("$"*100)
except Exception as e:
    print(e)
    print("An Exception Occured")    
    traceback.print_exc()    