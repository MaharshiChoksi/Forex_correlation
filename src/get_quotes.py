import datetime
import os
import MetaTrader5 as mt5
import pandas as pd
from dotenv import load_dotenv
import calculate_correlation
import argparse
load_dotenv(override=True)

parser = argparse.ArgumentParser()
parser.add_argument('-fn', '--filename',
                    type=str,
                    default=f"../data/correlation_quotes_{datetime.datetime.now().strftime('%B_%Y-%m-%d_%H-%M')}.xlsx",
                    help='pass the excel file location and file name where the quotes will be stored')
args = parser.parse_args()


class Quotes:
    def __init__(self, excel_file: str):
        self.symbols = ['AUDCAD', 'AUDJPY', 'AUDUSD', 'EURAUD', 'EURCAD',
                        'EURCHF', 'EURGBP', 'EURJPY', 'EURUSD', 'GBPCHF',
                        'GBPJPY', 'GBPUSD', 'USDCAD', 'USDCHF', 'USDJPY',
                        'NZDUSD',
                        'XAUUSD', 'XAGUSD', 'XTIUSD', 'XBRUSD',
                        'US30', 'US500', 'US2000']
        # Short-term: 15M, Med-term: 4Hr, Long-term: 1D
        self.timeframes = [mt5.TIMEFRAME_M15, mt5.TIMEFRAME_H4, mt5.TIMEFRAME_D1]
        self.datetime_today = datetime.date.today()
        self.excel_file = excel_file
        self.time_to_calculation = [datetime.datetime.strptime((self.datetime_today - datetime.timedelta(days=3 * 30)).strftime("%d.%m.%Y"), "%d.%m.%Y"),
                                    datetime.datetime.strptime((self.datetime_today - datetime.timedelta(days=365)).strftime("%d.%m.%Y"), "%d.%m.%Y"),
                                    datetime.datetime.strptime((self.datetime_today - datetime.timedelta(days=5 * 365)).strftime("%d.%m.%Y"), "%d.%m.%Y")]

    # @staticmethod
    # def get_filename(currency: str, timeframe: str, start_date_time: str, end_date_time: str) -> str:
    #     filename = f"{currency}_{timeframe}_{start_date_time}-00:00_{end_date_time}-00:00"
    #     return filename

    def start_mt5(self):
        if not mt5.initialize():
            print("MT5 terminal initialization failed!!")
            self.terminate_mt5_connection()
        try:
            mt5.login(server=os.environ['SERVER'], login=os.environ['LOGIN'], password=os.environ['PASSWORD'])
            print(f'ACCOUNT INFO: {mt5.account_info()}')
        except Exception as e:
            raise Exception("Error occurred while logging in to server", e)

    @staticmethod
    def quotes(tickers: list, tf: int, start_time: datetime.datetime) -> pd.DataFrame:  # returns quotes in form of pandas Dataframe
        quotes_dict = {}
        for ticker in tickers:
            rates = mt5.copy_rates_from(ticker, tf, start_time, 0)['close']
            print(f"Fetched rates for {ticker}.")
            quotes_dict[ticker] = rates
            del rates
        df = pd.DataFrame.from_dict(quotes_dict, orient='index')
        df = df.transpose()
        return df

    def pull_quotes_for_currencies(self):
        writer = pd.ExcelWriter(self.excel_file, engine='openpyxl')

        s_time = datetime.datetime.now()
        for i, tf in enumerate(self.timeframes):
            df = self.quotes(self.symbols, tf, self.time_to_calculation[i])
            sheet_name = {mt5.TIMEFRAME_M15: '15Minutes', mt5.TIMEFRAME_H4: '4Hour', mt5.TIMEFRAME_D1: '1Day'}.get(tf)  # gets name of the timeframe according to the iteration
            df.fillna(0, inplace=True)
            df.to_excel(writer, sheet_name, index=False)

        writer.save()
        print(f"Quotes saved successfully to {self.excel_file} &&\n Time Taken for process is: {datetime.datetime.now() - s_time}")
        return self

    @staticmethod
    def terminate_mt5_connection():
        mt5.shutdown()
        quit()


def main():
    try:
        quote = Quotes(args.filename)
        quote.start_mt5()
        quote.pull_quotes_for_currencies()
        calculate_correlation.corr_calculation(args.filename)
        # implement display_charts file to store charts to Excel file
    except Exception as e:
        print(e)


main()
