import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
drafts = namespace.GetDefaultFolder(16)


print(drafts.Items.count)