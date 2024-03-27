import pathlib
from tkinter import messagebox
import win32com.client as client
from tkinter import *
import customtkinter

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme('dark-blue')

root = customtkinter.CTk()
root.title('Tkinter.comm - Custom Tkinter Buttons')
root.geometry('600x350')

def correo():
    try:
        outlook = client.Dispatch("Outlook.Application")
        opciones_path = pathlib.Path("opciones.pdf")
        opciones_absolute = str(opciones_path.absolute())
        message = outlook.CreateItem(0) 
        message.SentOnBehalfOfName = 'prueba269@outlook.es'
        message.Attachments.Add(opciones_absolute)
        message.Subject = "¿Qué sigue cuando termine las materias de mi maestría?"
        message.Body = '''Saludos,\n
        Me complace informarle sobre las opciones de grado disponibles según cada programa de
        maestría en nuestra institución. Es fundamental que esté al tanto de estas alternativas, ya que
        su elección determinará el enfoque final de sus estudios y el tipo de proyecto que llevará a
        cabo para culminar su programa de maestría.\n
        Le insto a revisar detenidamente estas opciones y considerar cuál se ajusta mejor a sus
        intereses académicos y profesionales. Recuerde que su elección también debe estar en
        consonancia con los requisitos específicos del programa y las expectativas de su departamento
        académico.\n
        Si tiene alguna pregunta o necesita orientación adicional para tomar una decisión informada, no
        dude en ponerse en contacto con su asesor académico. Estamos aquí para ayudarlo/a en cada
        paso del camino hacia la culminación exitosa de su programa de maestría.\n
        Quedo a su disposición para cualquier consulta adicional que pueda surgir.
        '''
        message.Display()
    except Exception as e:
        messagebox.showerror("Error", f"Un error a ocurrido!: {e}")

my_button = customtkinter.CTkButton(root, text='Mandar correo', command=correo)
my_button.pack(pady=80)

root.mainloop()