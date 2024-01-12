#
# Python Module for
# Mean-Volatility Portfolio Analysis
#
import numpy as np
import pandas as pd
import xlwings as xw
from pylab import plt

plt.style.use('seaborn')

fn = 'http://hilpisch.com/tr_eikon_eod_data.csv'
raw = pd.read_csv(fn, index_col=0,
        parse_dates=True).dropna()

rets = np.log(raw / raw.shift(1))


def plot_data():
    wb = xw.Book.caller()
    sht = wb.sheets.active
    fig, ax = plt.subplots()
    (raw / raw.iloc[0]).plot(ax=ax)
    sht.pictures.add(fig, name='Financial Data',
            update=True)

def portfolio_return(port):
    data = rets[port['Symbols']]
    return np.dot(data.mean() * 252, port['Weights'])

def portfolio_volatility(port):
    data = rets[port['Symbols']]
    var = np.dot(port['Weights'],
            np.dot(data.cov() * 252, port['Weights']))
    return var ** 0.5

def calculate_statistics():
    wb = xw.Book.caller()
    sht = wb.sheets.active
    port = sht.range('B6').expand().options(
            pd.DataFrame, index=False).value
    # sht.range('B16').value = port
    pr = portfolio_return(port)
    sht.range('B17').value = pr
    pv = portfolio_volatility(port)
    sht.range('C17').value = pv

