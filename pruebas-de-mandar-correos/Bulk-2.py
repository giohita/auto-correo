import csv
import time
import win32com.client as client
with open('people.csv', newline='') as f:
    reader = csv.reader(f)
    distro = [row for row in reader]

#Cuidado corriendo este codigo que tengo mucho contenido en el csv
template = "{}, please submit your timeas soon as posible!"
outlook = client.Dispatch("Outlook.Application")
for name, adress in distro:
    message = outlook.CreateItem(0)
    message.To = adress
    message.Subject = "Your time entry is past due!"
    message.Body = template.format(name)
    message.Display() 