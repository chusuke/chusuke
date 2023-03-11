from datetime import datetime, date, timedelta

today = datetime.today()
tomorrow = today + timedelta(days=1)
dft = today + timedelta(days=2)
yesterday = today - timedelta(days=1)
print("tomorrow -> " + datetime.strftime(tomorrow, '%Y-%m-%d'))
print("dft -> " + datetime.strftime(dft, '%Y-%m-%d'))
print(datetime.strftime(today, '%Y%m%d'))
print("yesterday -> " + datetime.strftime(yesterday, '%Y-%m-%d'))