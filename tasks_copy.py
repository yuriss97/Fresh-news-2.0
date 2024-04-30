from robocorp.tasks import task
from robocorp import browser
import time
from RPA.Excel.Files import Files
import datetime
from pytz import timezone
import re
import requests
import os
from dateutil.relativedelta import relativedelta

@task
def get_limit_date():
    # Current time in UTC
    cloud_number_of_months = 5

    now_utc = datetime.datetime.now(timezone('UTC'))

    # Convert to Asia/Qatar time zone
    now_qatar = now_utc.astimezone(timezone('Asia/Qatar'))

    if cloud_number_of_months > 1:
        #days_to_subtract = (cloud_number_of_months - 1) * 30
        #now_qatar = now_qatar - datetime.timedelta(days=days_to_subtract)
        now_qatar = now_qatar - relativedelta(months=cloud_number_of_months-1)
    
    return now_qatar.replace(day=1).replace(tzinfo=None)

