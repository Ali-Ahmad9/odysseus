import win32com.client as win32
from lib.constants import Constants

def send_email():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = Constants.EMAIL
    mail.Subject = 'Test'
    mail.HTMLBody = '<h2>This is a test</h2>'  # this field is optional
    mail.Send()


if __name__ == '__main__':
    send_email()
