import os
import win32com.client as client
from PIL import ImageGrab

#Copy and Save Excel range as Image
workbook_path =os.getcwd() + '\\heatmap.xlsx'
excel = client.Dispatch('Excel.Application')
wb = excel.Workbooks.Open (workbook_path)
sheet = wb.Sheets.Item(1)
sheet = wb.Sheets('Sheet1')
excel.visible = 1
copyrange = sheet.Range('A1:M11')
copyrange.Select()
copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save('paste.png')
excel.Quit()

#Create Outlook email and insert Excel content
image_path = os.getcwd() + '\\paste.png'
html_body = """
    <div>
        Please review the following report and response with your feedback<br></br>
    </div>
    <div>
        <img src="cid:paste.png"></img>
    </div>
"""
outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "giohandmelo@outlook.com"
message.Subject = "Please review!"
image_attachment = message.Attachments.Add(image_path)
image_attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "paste.png")
message.HTMLBody = html_body
message.Display()