#Instale el pip pywin32

#Prueba de como mandar un correo en Outlook
import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()
message.To = "giohandmelo@outlook.com"
message.CC = "giohandmelo@outlook.com"
message.BCC = "giohandmelo@outlook.com"
message.Subject = "Happy Birthday"
message.Body = "Wish you a happy birthday!"
message.Save()
message.Send()