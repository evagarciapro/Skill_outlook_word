import win32com.client
from datetime import datetime, timedelta

def get_outlook_emails(days=7):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6)   # Boîte de réception
    sent = outlook.GetDefaultFolder(5)    # Éléments envoyés

    cutoff_date = datetime.now() - timedelta(days=days)

    emails = []

    def read_folder(folder, folder_name):
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        for item in items:
            try:
                received_time = getattr(item, "ReceivedTime", None)
                if received_time and received_time.replace(tzinfo=None) >= cutoff_date:
                    emails.append({
                        "folder": folder_name,
                        "subject": item.Subject,
                        "sender": getattr(item, "SenderName", ""),
                        "to": getattr(item, "To", ""),
                        "date": received_time.strftime("%Y-%m-%d %H:%M"),
                        "body": getattr(item, "Body", "")
                    })
            except Exception:
                continue

    read_folder(inbox, "reçus")
    read_folder(sent, "envoyés")

    return emails