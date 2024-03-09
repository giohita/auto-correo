import csv
from time import sleep
import win32com.client as client

template =  "{}, please submit your time as soon as posible."
with open("people.csv", "r", newline="") as f:
    reader = csv.reader(f)
    distro = [row for row in reader]

chunks = [distro[:x+30] for x in range(0, len(distro), 30)]
outlook = client.Dispatch("Outlook.Application")
for chunk in chunks:
    for name, email in chunk:
        message = outlook.CreateItem(0)
        message = outlook.CreateItem(0)
        message.To = email
        message.Subject = "Your time entry is past due!"
        message.Body = template.format(name)
        message.Send()
    sleep(60)
