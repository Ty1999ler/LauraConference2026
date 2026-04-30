import win32com.client
import pythoncom


def get_outlook_folder(folder_path: list):
    """
    Navigate from the Inbox down through each name in folder_path.
    Returns the final folder COM object.
    """
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # 6 = olFolderInbox
        folder = namespace.GetDefaultFolder(6)

        for name in folder_path[1:]:  # skip "Inbox" — already there
            try:
                folder = folder.Folders(name)
            except Exception:
                raise RuntimeError(
                    f"Could not find folder: '{name}' inside '{folder.Name}'. "
                    f"Check that FOLDER_PATH in config.py matches the real folder names."
                )
        return folder

    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Failed to connect to Outlook: {exc}") from exc


def get_folder_items(folder) -> list:
    """
    Return all MailItem objects from folder, sorted newest-first.
    Non-mail items (meeting requests, etc.) are skipped.
    """
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # True = descending

        mail_items = []
        for item in items:
            try:
                # olMailItem class constant = 43
                if item.Class == 43:
                    mail_items.append(item)
            except AttributeError:
                continue
            except Exception:
                continue

        return mail_items

    except Exception as exc:
        raise RuntimeError(f"Failed to read folder items: {exc}") from exc
