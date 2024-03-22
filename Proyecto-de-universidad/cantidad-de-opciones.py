import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)


message.Subject = "¿Qué sigue cuando termine las materias de mi maestría?"
message.Body = '''
'''

message.Display()