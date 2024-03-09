import win32com.client
import pathlib

cake_path = pathlib.Path("birthday-cake.jpg")
cert_path = pathlib.Path("certificate.jfif")
cake_absolute = str(cake_path.absolute())
cert_absolute = str(cert_path.absolute())

outlook = win32com.client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()

message.To = "giohandmelo@outlook.com"
message.Subject = "Happy Birthday"
message.Attachments.Add(cert_absolute)

# Adjuntar la imagen
image = message.Attachments.Add(cake_absolute)
image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "cake-img")

# Crear el cuerpo HTML con referencia a la imagen adjunta
html_body = """
    <div>
        <h1 style="font-family: 'Lucida Handwriting'; font-size: 56; font-weight: bold; color: #9eac9c;"> Happy Birthday!! </h1>
        <span style="font-family: 'Lucida Sans'; font-size: 28; color: #8d395c;"> Wishing you all the best on your birthday!! </span>
    </div><br>
    <div>
        <img src="cid:cake-img" width=50%>
    </div>
"""

# Establecer el cuerpo HTML del correo
message.HTMLBody = html_body