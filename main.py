import win32com.client as win32
from lib.constants import Constants
from yahoo_fin import stock_info


TICKERS = [
    'AAPL',
    'AAL',
    'MCX'
]


def get_live_prices():
    return [(ticker, stock_info.get_live_price(ticker)) for ticker in TICKERS]


def send_email():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = Constants.EMAIL
    mail.Subject = 'Test'
    mail.Body = str(get_live_prices())
    mail.Send()


def main():
    send_email()


if __name__ == '__main__':
    main()
