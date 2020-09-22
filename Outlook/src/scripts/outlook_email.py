# This function will send an email with or without an attachment using Microsoft Outlook
# pip install pypiwin32
# pypiwin32 version 223

import win32com.client as win32


def send_outlook_email(text, subject, recipient, attachment_list, auto=True):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    for file in attachment_list:
        mail.Attachments.Add(file)

    if auto:
        mail.send
    else:
        mail.Display(True)


send_outlook_email("test text", "test subject", "testemail@gmail.com", [], auto=True)
