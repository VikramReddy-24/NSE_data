from datetime import datetime
import pytz



def date_time_format(recived_data):
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
    time = timestamp_datetime.strftime("%H%M") #:%S")

    print(time)

    return date,month,year,time 

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


def find_nearest_number(target, numbers):
    nearest_number = min(numbers, key=lambda x: abs(x - target))
    return nearest_number

ce_list=['%chg_ce','chg_ce','open_ce',
            'high_ce',
            'low_ce',
            'close_ce',
            'ltp_ce',
            'p_close_ce',
            'qty_ce',
            'tto_ce',
            'tradedVolume_ce',
            'value_ce',
            'vmap_ce',
            'premiumTurnover_ce',
            'oi_ce',
            'coi_ce',
            '%coi_ce',
            'Lot_ce',
            'sp_ce',
            'dvol_ce',
            'annualisedVolatility_ce',
            'IV_ce',
            'CWPL_ce',
            'MWPL_ce',
            'volumeFreezeQuantity_ce',
            'totalBuyQuantity_ce',
            'totalSellQuantity_ce',
            'priceBestBuy_ce',
            'priceBestSell_ce',
            'pricelastPrice_ce','carryBestBuy_ce','carryBestSell_ce','carrylastPrice_ce']

pe_list=[   'ltp_pe',
            'close_pe',
            'high_pe',
            'low_pe',
            'open_pe',
            '%chg_pe',
            'chg_pe',
            'p_close_pe',
            'qty_pe',
            'tto_pe',
            'tradedVolume_pe',
            'value_pe',
            'vmap_pe',
            'premiumTurnover_pe',
            'oi_pe',
            'coi_pe',
            '%coi_pe',
            'Lot_pe',
            'sp_pe',
            'dvolt_pe',
            'annualisedVolatility_pe',
            'IV_pe',
            'CWPL_pe',
            'MWPL_pe',
            'volumeFreezeQuantity_pe',
            'totalBuyQuantity_pe',
            'totalSellQuantity_pe',
            'priceBestBuy_pe',
            'priceBestSell_pe',
            'pricelastPrice_pe',
            'carryBestBuy_pe',
            'carryBestSell_pe',
            'carrylastPrice_pe']

index_list=['%chg',
    'chg',
    'ltp',
    'high',
    'low',
    'hl',
    '%hl',
    'open']

def main_dataFrame_style(df):
    style_df=df.style.set_properties(subset = ce_list,**{"background-color": "#ccffcc"}) \
                       .set_properties(subset = pe_list,**{"background-color": "#ffd6cc"}) \
                      .set_properties(subset=index_list,**{"background-color": "#ccffff"})
                           
    return style_df                       