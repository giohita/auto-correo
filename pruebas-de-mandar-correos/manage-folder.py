import win32com.client as client

# Crear una instancia de Outlook
outlook = client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Obtener la carpeta de borradores (Drafts) y la carpeta de la bandeja de entrada (Inbox)
drafts = namespace.GetDefaultFolder(16)
inbox = namespace.GetDefaultFolder(6)

# Acceder a la carpeta "Python" en la bandeja de entrada (asegúrate de que esta carpeta exista)
pyfolder = inbox.Folders["Python"]

# Alternativamente, puedes acceder a la carpeta de la primera posición
pyfolder = inbox.Folders[0]


# También puedes usar el método "Item" para acceder a la carpeta por índice
pyfolder = inbox.Folders.Item(1)

# Iterar sobre todas las carpetas en la bandeja de entrada
for folder in inbox.Folders:
    print(folder.Name)

# Obtener el número total de carpetas en la bandeja de entrada
num_folders = inbox.Folders.Count
print("Total folders in Inbox:", num_folders)

# Agregar una nueva carpeta llamada "YTVideos" en la bandeja de entrada
inbox.Folders.Add("YTVideos")

# Corregir un error tipográfico en el nombre de la carpeta "YTVideos"
inbox.Folders["YTVideos"].Description = "Messages related to my channel"
print("Inbox Description:", inbox.Description)

# Acceder a la ruta de la carpeta
print("Inbox Folder Path:", inbox.FolderPath)

# Acceder al nombre del padre de la carpeta "Python"
parent = pyfolder.Parent
print("Parent Folder Name:", parent.Name)

# Acceder al asunto del primer elemento en la carpeta de la bandeja de entrada
print("Subject of the first item:", inbox.Items[0].Subject)

# Crear una nueva carpeta llamada "MyNewFolder", moverla a la carpeta "Python" y luego eliminarla
newfolder = inbox.Folders.Add("MyNewFolder")
newfolder.MoveTo(pyfolder)
newfolder.Delete()

# Iterar sobre todas las carpetas en la biblioteca de nombres
for folder in namespace.Folders:
    print(folder.Name)

# Acceder a la cuenta de Gmail y la carpeta "Inbox" (ajusta según tu configuración)
gmail = namespace.Folders["giohandmelo@gmail.com"]
print("Gmail Account Name:", gmail.Name)

gmail_inbox = gmail.Folders["Inbox"]
print("Gmail Inbox Name:", gmail_inbox.Name)

# Alternativamente, puedes acceder a la carpeta de la primera posición
gmail_inbox = gmail.Folders[0]
print("Gmail Inbox Name:", gmail_inbox.Name)

# También puedes usar el método "Item" para acceder a la carpeta por índice
gmail_inbox = gmail.Folders.Item(1)
print("Gmail Inbox Name:", gmail_inbox.Name)
