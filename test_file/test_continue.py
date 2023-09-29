import datetime
from datetime import timedelta
now = datetime.datetime.now().date()
print (now)
year = int(now.strftime('%Y'))
month = int(now.strftime('%m'))
print (year, month)
start_date_obj = datetime.datetime(year,month,1) - timedelta(days=1)
start_date = start_date_obj.strftime('%Y-%m-%d')
print (start_date)
end_date = now
print (end_date)
