import MetaTrader5 as mt5
from datetime import datetime
import pandas as pd
from pandas.plotting import register_matplotlib_converters

register_matplotlib_converters()

# connect to MetaTrader 5
if not mt5.initialize():
    print("initialize() failed")
    mt5.shutdown()

# request connection status and parameters
print(mt5.terminal_info())
# get data on MetaTrader 5 version
print(mt5.version())

