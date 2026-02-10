import datetime
import time

#Get current date
now = datetime.datetime.now()
print(now)
today = datetime.datetime.today()
print(today)

#Custom Date
d = datetime.date(2025, 11, 21)
print("Custom Date:", d)

#Date Arithmetic
today = datetime.datetime.today()
tomorrow = today + datetime.timedelta(days=1)
print("Tomorrow:", tomorrow)

#Formatting & Parsing
formatted = today.strftime("%d/%m/%Y")
print("Formatted:", formatted)

parsed = datetime.datetime.strptime("01/01/2025", "%d/%m/%Y")
print("Parsed Date:", parsed)

#pasuing execution for a countdown effect
halloween2024 = datetime.datetime(2024, 11, 21,0,0,0)
while datetime.datetime.now() < halloween2024:
    time.sleep(1)

now = time.time()
print(time.time())

print(time.ctime(now))

