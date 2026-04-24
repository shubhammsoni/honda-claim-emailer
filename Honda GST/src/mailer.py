import win32com.client as win32
import os


def send_email(to_email, cc_list, subject, html_body, attachment_path=None):
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = to_email
        mail.CC = ";".join(cc_list) if cc_list else ""
        mail.Subject = subject
        mail.HTMLBody = html_body

        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)

        mail.Send()

        return True, None

    except Exception as e:
        return False, str(e)