import win32com.client as client

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = "giohandmelo@outlook.com"
message.Subject = "Happy Birthday"

