#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time         : 2020/11/9 11:29
# @Author       : Chizhuo Zhou (Github: Tridro)
# @Organization : Everbright Futures
# @E-mail       : tridro@beneorigin.com
# @File         : futures_calendar_deviation_monitor.py
# @Software     : PyCharm
# All CopyRight Reserved

from os import system
from sys import stdout
from datetime import datetime
from threading import Thread
from tqsdk import TqApi, TqAuth
from numpy import poly1d, polyfit, isnan
from matplotlib.pyplot import rcParams, xticks, autoscale, figure, show
from matplotlib.animation import FuncAnimation
from matplotlib import style

system("mode con cols=71 lines=30")
stdout.write(f"+{'-' * 68}+\n"
             f"|{'期货跨月偏离度监测':^59}|\n"
             f"+{'-' * 68}+\n"
             f"|{' ':68}|\n"
             f"|{'数据行情由天勤量化提供支持，如不同意天勤量化使用条款，请终止使用':^36}|\n"
             f"|{' ':68}|\n"
             f"+{'-' * 68}+\n\n")
tq_account = input('天勤量化账号: ')
tq_password = input('天勤量化密码: ')
try:
    api = TqApi(auth=TqAuth(tq_account, tq_password))
except Exception as err:
    print(err, '\n')
    input(f"{'连接失败, 按任意键程序将退出, 请检查后重试'}")
    raise SystemExit()
quotes = []
labels = []
product = input("\n请输入品种代码(例rb/hc): ").lower()
try:
    available_contracts = sorted(api.query_quotes(ins_class="FUTURE", product_id=product, expired=False))
    m_contract = api.query_cont_quotes(product_id=product)[0]
except IndexError:
    available_contracts = sorted(api.query_quotes(ins_class="FUTURE", product_id=product.upper(), expired=False))
    m_contract = api.query_cont_quotes(product_id=product.upper())[0]
m_contract_index = available_contracts.index(m_contract)


def batch_get_quote(contracts_list):
    for code_full in contracts_list:
        quotes.append(api.get_quote(code_full))
        labels.append(code_full.split(".")[1])
    if product == 'rb' or product == 'hc':
        labels.insert(0, '现货')


batch_get_quote(available_contracts)


def fetch_benchmark(method='current', quote_idx=0):
    if method == 'current':
        return quotes[quote_idx].pre_settlement if isnan(quotes[quote_idx].last_price) else quotes[quote_idx].last_price
    elif method == 'last':
        return quotes[quote_idx].pre_settlement if isnan(quotes[quote_idx].pre_close) else quotes[quote_idx].pre_close


spot_price = float(input(f"{product}当日现货价格:") or 'nan') if product in ['rb', 'hc'] \
    else fetch_benchmark(method='current', quote_idx=0)
last_spot_price = float(input(f"{product}昨日现货价格:") or 'nan') if product in ['rb', 'hc'] \
    else fetch_benchmark(method='last', quote_idx=0)

xs = range(len(labels))
bias_list = [0.0 for i in xs]
last_bias_list = [0.0 for i in xs]
open_interest_change_list = [0 for i in xs]
trend_line_list = [0.0 for i in xs]
last_trend_line_list = [0.0 for i in xs]
fig = figure(figsize=(11, 6))
ax = fig.add_axes([0.1, 0.15, 0.8, 0.8])
ln1, = ax.plot(xs, bias_list, label='实时偏离', color='red', marker='o', linestyle='solid', linewidth=2, markersize=4)
ln2, = ax.plot(xs, last_bias_list, label='昨收偏离', color='blue', marker='o', linestyle='solid', linewidth=2, markersize=4)
tln1, = ax.plot(xs, trend_line_list, label='实时趋势线', color='red', linestyle='dashed', linewidth=1)
tln2, = ax.plot(xs, trend_line_list, label='昨日趋势线', color='blue', linestyle='dashed', linewidth=1)
ax2 = ax.twinx()
ba = ax2.bar(xs, open_interest_change_list, width=0.35, label='持仓增减', align='center')
rect_text_list = []
ax.grid()


def init():
    rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
    rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
    style.use('fivethirtyeight')
    ax.set_xlabel('合约')
    ax.set_ylabel('偏离度')
    ax2.set_ylabel('持仓')
    lns, lns_label = ax.get_legend_handles_labels()
    bas, bas_label = ax2.get_legend_handles_labels()
    ax.legend(lns + bas, lns_label + bas_label,
              loc='lower center',
              ncol=5,
              bbox_to_anchor=(0.1, -0.13, 0.8, 0.2),
              fontsize='small',
              borderaxespad=0.,
              mode='expand')
    for rect in ba:
        height = rect.get_height()
        rect_text = ax2.annotate(text='',
                                 xy=(rect.get_x() + rect.get_width() / 2, height),
                                 xytext=(0, 3),
                                 textcoords="offset points",
                                 ha='center',
                                 va='bottom',
                                 fontsize='small')
        rect_text_list.append(rect_text)
    xticks(xs, labels)
    autoscale(enable=True, axis='both', tight=True)
    return ln1, ln2,


def trend_line(x, y, n):
    if product == 'rb' or product == 'hc':
        f = poly1d(polyfit(x[1:], y[1:], n))
        return [f(i) for i in x]
    else:
        f = poly1d(polyfit(x, y, n))  # n=1为一次函数，返回函数参数
        return [f(i) for i in x]


def daily_diff_average(method='current'):
    days_interval = datetime.fromtimestamp(quotes[-1].expire_datetime) - datetime.today()
    if method == 'current':
        if product == 'rb' or product == 'hc':
            diff = spot_price - fetch_benchmark(method, quote_idx=-1)
        else:
            diff = fetch_benchmark(method, quote_idx=0) - fetch_benchmark(method, quote_idx=-1)
        return diff / days_interval.days
    elif method == 'last':
        if product == 'rb' or product == 'hc':
            diff = last_spot_price - fetch_benchmark(method, quote_idx=-1)
        else:
            diff = fetch_benchmark(method, quote_idx=0) - fetch_benchmark(method, quote_idx=-1)
        return diff / days_interval.days


def price_bias(quote, method='current'):
    days_interval = datetime.fromtimestamp(quote.expire_datetime) - datetime.today()
    if method == 'current':
        theoretical_diff = daily_diff_average(method) * days_interval.days
        if product == 'rb' or product == 'hc':
            theoretical_price = spot_price - theoretical_diff
        else:
            theoretical_price = fetch_benchmark(method, quote_idx=0) - theoretical_diff
        return quote.last_price - theoretical_price
    elif method == 'last':
        theoretical_diff = daily_diff_average(method) * days_interval.days
        if product == 'rb' or product == 'hc':
            theoretical_price = last_spot_price - theoretical_diff
        else:
            theoretical_price = fetch_benchmark(method, quote_idx=0) - theoretical_diff
        return quote.pre_close - theoretical_price


def data_process():
    global bias_list, last_bias_list, trend_line_list, last_trend_line_list, open_interest_change_list
    while True:
        api.wait_update()
        if api.is_changing(quotes[m_contract_index], 'datetime'):
            if len(set(last_bias_list)) == 1 and 0.0 in set(last_bias_list):
                del last_bias_list[:]
                for i in range(len(quotes)):
                    last_bias = price_bias(quotes[i], method='last')
                    last_bias_list.append(last_bias if not isnan(last_bias) else 0.0)
                if product == 'rb' or product == 'hc':
                    last_bias_list.insert(0, 0.0)
                last_trend_line_list = trend_line(xs, last_bias_list, 1)
            del bias_list[:]
            del open_interest_change_list[:]
            for i in range(len(quotes)):
                bias = price_bias(quotes[i], method='current')
                bias_list.append(bias if not isnan(bias) else 0.0)
                open_interest_change_list.append(quotes[i].open_interest - quotes[i].pre_open_interest)
            if product == 'rb' or product == 'hc':
                bias_list.insert(0, 0.0)
                open_interest_change_list.insert(0, 0)
            trend_line_list = trend_line(xs, bias_list, 1)


def fetch_data():
    t = Thread(target=data_process)
    t.start()
    while True:
        yield bias_list, last_bias_list, trend_line_list, last_trend_line_list, open_interest_change_list


def update_fig(data):
    ys1, ys2, ys3, ys4, ys5 = data
    for rect, height, rect_text in zip(ba.patches, ys5, rect_text_list):
        rect.set_height(height)
        rect_text.xy = (rect.get_x() + rect.get_width() / 2, height)
        rect_text.set_text('{}'.format(height))
        rect_text.update_positions(rect_text)
    left_axis_lower = min(min(ys1), min(ys2), min(ys3), min(ys4)) * 1.1
    left_axis_upper = max(max(ys1), max(ys2), max(ys3), max(ys4)) * 1.1
    if left_axis_lower < left_axis_upper:
        ax.set_ylim(left_axis_lower, left_axis_upper)
    right_axis_lower = min(ys5) * 1.2
    right_axis_upper = max(ys5) * 1.2
    if right_axis_lower < right_axis_upper:
        ax2.set_ylim(right_axis_lower, right_axis_upper)
    fig.canvas.draw()
    ln1.set_ydata(ys1)
    ln2.set_ydata(ys2)
    tln1.set_ydata(ys3)
    tln2.set_ydata(ys4)
    return ln1, ln2, tln1, tln2,


try:
    stdout.write('\n运行中...')
    ani = FuncAnimation(fig, update_fig, fetch_data, init_func=init, interval=500, blit=True)
    show(block=True)
except RuntimeError:
    raise SystemExit()
