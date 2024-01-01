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