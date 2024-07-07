import MetaTrader5 as mt5
import datetime


class Quotes:
    def __init__(self):
        self.Currencies = ['AUDCAD', 'AUDJPY', 'AUDUSD', 'EURAUD', 'EURCAD',
                           'EURCHF', 'EURGBP', 'EURJPY', 'EURUSD', 'GBPCHF',
                           'GBPJPY', 'GBPUSD', 'USDCAD', 'USDCHF', 'USDJPY',
                           'NZDUSD']
        # Short-term: 15M, Med-term: 4Hr, Long-term: 1D
        self.timeframes = ['15M', '4H', '1D']
        self.datetime_today = datetime.date.today()
        self.time_to_calculation = [datetime.date.today() - datetime.timedelta(days=3 * 30),
                                    datetime.date.today() - datetime.timedelta(days=365),
                                    datetime.date.today() - datetime.timedelta(days=5*365)]
        self.start_mt5()

    def start_mt5(self):
        if not mt5.initialize(path="C:\\Program Files\\MetaTrader 5\\terminal64.exe",
                              server='ICMarketsSC-Demo', login=51749804, password='a5CiLOYY@KDw1e',
                              headless=True):
            self.terminate_mt5_connection()
            return "MT5 initialization failed!!"
        print(f'TERMINAL INFO: {mt5.terminal_info()}')
        print(f'VERSION: {mt5.version()}')

    @staticmethod
    def terminate_mt5_connection():
        mt5.shutdown()
        quit()


Quotes()
