import win32com.client as client

distro = []
for i in range(1567):
    distro.append("giohandmelo@outlook.com")

chunks = []
for x in range(0, len(distro), 500):
    chunks.append(distro[x:x+500])

outlook = client.Dispatch("Outlook.Application")
for recipients in chunks:
    print("Count of items:", len(recipients))

    message = outlook.CreateItem(0)
    message.To = ";".join(recipients)
    message.Subject = "Missing time alert!"
    message.Body = "Please submit your time as soon as possible!"
    message.Display()
