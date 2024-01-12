#
# Modern Portfolio Theory
# with Python and Excel
#
# The Python Quants GmbH
#
import numpy as np
import pandas as pd
import xlwings as xw
from pylab import plt
from scipy.optimize import minimize

plt.style.use('seaborn')

# reading & preparing the data
raw = pd.read_csv('http://hilpisch.com/tr_eikon_eod_data.csv',
                 index_col=0, parse_dates=True).dropna()
rets = np.log(raw / raw.shift(1))

# core portfolio statistics
def portfolio_return(weights, symbols):
    return np.dot(weights, rets[symbols].mean() * 252)

def portfolio_volatility(weights, symbols):
    return np.dot(weights, np.dot(rets[symbols].cov() * 252,
                                  weights)) ** 0.5


# advanced portfolio calculations & simulations
def derive_min_risk_port(symbols):
    nos = len(symbols)
    bnds = [(0, 1) for _ in range(nos)]
    cons = {'type': 'eq',
            'fun': lambda weights: weights.sum() - 1}
    opt = minimize(portfolio_volatility, nos * [1 / nos],
               bounds=bnds, constraints=cons,
               args=(symbols))
    return opt['x']

def simulate_port_stats(symbols, runs=250):
    rpa = np.random.random((runs, len(symbols)))
    rpa = (rpa.T / rpa.sum(axis=1)).T
    pstats = [(portfolio_volatility(w, symbols),
           portfolio_return(w, symbols)) for w in rpa]
    return np.array(pstats)

# basic UDFs
@xw.func
@xw.arg('params', pd.DataFrame, index=False)
def calculate_port_return(params):
    return portfolio_return(params['weights'], params['symbols'])

@xw.func
@xw.arg('params', pd.DataFrame, index=False)
def calculate_port_volatility(params):
    return portfolio_volatility(params['weights'], params['symbols'])

# advanced UDFs
@xw.func
@xw.ret(transpose=True, expand='table')
def calculate_min_risk_weights(symbols):
    return derive_min_risk_port(symbols)

# script to be called
def generate_simulation_plot():
    wb = xw.Book.caller()
    sht = wb.sheets.active
    runs = int(sht.range('B16').value)
    symbols = sht.range('B4:B7').value
    ps = simulate_port_stats(symbols, runs)
    fig = plt.figure()
    plt.plot(ps[:, 0], ps[:, 1], 'ro')
    plt.xlabel('portfolio volatility')
    plt.ylabel('portfolio return')
    plt.title(' | '.join(symbols))
    sht.pictures.add(fig, name='mpt', update=True)
    
    
    
    
    



