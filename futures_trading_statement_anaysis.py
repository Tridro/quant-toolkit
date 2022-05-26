#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time     : 2019/10/16 11:04
# @Author   : 周驰卓
# @Company  : 光大期货
# @Site     :
# @File     : 期货交易结算单数据分析.py
# @Software : PyCharm

import getopt
import os.path
import re
import sys
from datetime import datetime
from itertools import islice

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_interval
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers
from openpyxl.chart import Reference, LineChart, PieChart, BarChart, RadarChart
from openpyxl.chart.axis import NumericAxis, TextAxis
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout

# ---------------------------------------------------- 基础数据 开始 ----------------------------------------------------
BASE_DIR = os.path.dirname(os.path.realpath(sys.argv[0]))

CONTRACT_CODE = {'if': '沪深300股指', 'ih': '上证50股指', 'ic': '中证500股指', 'tf': '五债', 't': '十债', 'ts': '二债',
                 'cu': '铜', 'al': '铝', 'zn': '锌', 'pb': '铅', 'ni': '镍', 'sn': '锡', 'au': '黄金', 'ag': '白银',
                 'j': '焦炭', 'jm': '焦煤', 'zc': '动力煤ZC',
                 'rb': '螺纹钢', 'i': '铁矿石', 'hc': '热轧卷板', 'sf': '硅铁', 'sm': '锰硅', 'fg': '玻璃', 'ss': '不锈钢',
                 'wr': '线材', 'ru': '天然橡胶', 'sp': '漂针浆', 'bb': '细木工板', 'fb': '纤维板', 'nr': '20号胶',
                 'fu': '燃料油', 'bu': '石油沥青', 'l': '线型低密度聚乙烯', 'pp': '聚丙烯', 'v': '聚氯乙烯', 'ta': '精对苯二甲酸',
                 'ma': '甲醇MA', 'eg': '乙二醇', 'eb': '苯乙烯', 'ur': '尿素', 'sa': '纯碱', 'pg': '液化石油气', 'lu': '低硫燃料油',
                 'm': '豆粕', 'y': '豆油', 'oi': '菜籽油', 'a': '黄大豆1号', 'b': '黄大豆2号', 'p': '棕榈油', 'c': '黄玉米',
                 'rm': '菜籽粕', 'cs': '玉米淀粉',
                 'cf': '一号棉花', 'cy': '棉纱', 'sr': '白砂糖', 'wh': '强筋小麦', 'ri': '旱籼稻', 'rr': '粳米',
                 'rs': '油菜籽', 'jr': '粳稻谷', 'lr': '晚籼稻', 'pm': '普通小麦', 'sc': '原油',
                 'ap': '鲜苹果', 'jd': '鲜鸡蛋', 'cj': '干制红枣', 'pf': '短纤', 'bc': '国际铜', 'lh': '生猪', 'pk': '花生'}

TRADING_UNIT = {'if': 300, 'ih': 300, 'ic': 200, 'tf': 10000, 't': 10000, 'ts': 20000, 'cu': 5, 'al': 5, 'zn': 5,
                'sc': 1000, 'pb': 5, 'ni': 1, 'sn': 1, 'au': 1000, 'ag': 15, 'j': 100, 'jm': 60, 'zc': 100, 'rb': 10,
                'i': 100, 'hc': 10, 'sf': 5, 'sm': 5, 'wr': 10, 'fu': 10, 'bu': 10, 'ru': 10, 'l': 5, 'pp': 5, 'v': 5,
                'ta': 5, 'ma': 10, 'sp': 10, 'm': 10, 'y': 10, 'oi': 10, 'a': 10, 'b': 10, 'p': 10, 'c': 10, 'rm': 10,
                'cs': 10, 'jd': 10, 'bb': 500, 'fb': 500, 'cf': 5, 'cy': 5, 'sr': 10, 'wh': 20, 'ri': 20, 'jr': 20,
                'lr': 20, 'fg': 20, 'ss': 5, 'nr': 10, 'eg': 10, 'eb': 5, 'ur': 20, 'rr': 10, 'rs': 10, 'ap': 10,
                'cj': 5, 'pm': 50, 'sa': 20, 'pg': 20, 'lu': 10, 'pf': 5, 'bc': 5, 'lh': 16, 'pk': 5}


# ---------------------------------------------------- 基础数据 结束 ----------------------------------------------------


# ---------------------------------------------------- 数据读取 开始 ----------------------------------------------------
def read_statement_files(dir_input=''):
    print(f"+{'-' * 23}  使用准备  {'-' * 23}+\n" +
          f"|{' ':58}|\n" +
          f"|{'请把所有交易结算单txt文件放入当前或指定路径的文件夹中!':^33}|\n" +
          f"|{' ':58}|\n" +
          f"+{'-' * 58}+\n")
    if os.path.isdir(dir_input):
        if os.path.exists(dir_input):
            path = dir_input
            folder_name = os.path.basename(path)
        else:
            input(f"\n{datetime.now()} | 错误 | {dir_input} 路径错误，请检查路径并重试，按任意键退出！\n")
            raise SystemExit()
    else:
        folder = input("输入文件夹名/文件夹路径：") if dir_input == '' else dir_input
        path = os.path.join(BASE_DIR, folder) if not os.path.isdir(folder) else folder
        folder_name = os.path.basename(path)
        if not os.path.exists(path):
            input(f"\n{datetime.now()} | 错误 | 找不到 {folder_name} 文件夹，请检查并重试，按任意键退出！\n")
            raise SystemExit()
    files_list = os.listdir(path)
    statement_data_list = []
    for file in files_list:
        file_dir = os.path.join(path, file)
        if os.path.isfile(file_dir):
            with open(file_dir) as f:
                statement_data_list.append(f.readlines())
    print(f'\n{datetime.now()} | 信息 | 已读取 {folder_name} 文件夹')
    return statement_data_list


# ---------------------------------------------------- 数据读取 结束 ----------------------------------------------------


# ---------------------------------------------------- 数据提取 开始 ----------------------------------------------------
def data_extract(source, client_id=''):
    if client_id == '':
        client_id = re.search(r'[^客户号 ClientID：][0-9]+', source[0][8]).group()
    print(f'{datetime.now()} | 信息 | 开始客户号 {client_id} 的结算数据提取')
    account = pd.DataFrame(columns=['日期', '期初结存', '出入金', '平仓盈亏', '持仓盯市盈亏', '手续费', '期末结存',
                                    '客户权益', '保证金占用', '风险度'])
    transaction_record = pd.DataFrame(columns=['成交日期', '交易所', '品种', '合约', '买/卖', '投/保', '成交价', '手数',
                                               '成交额', '开/平', '手续费', '平仓盈亏', '权利金收支', '成交序号'])
    position_closed = pd.DataFrame(columns=['平仓日期', '交易所', '品种', '合约', '开仓日期', '买/卖', '手数', '开仓价',
                                            '昨结算', '成交价', '平仓盈亏', '权利金收支', '交易盈亏', '盈亏率'])
    sep = re.compile(r'[|\s|]+')

    for i in range(len(source)):
        if client_id != re.search(r'[^客户号 ClientID：][0-9]+', source[i][8]).group():
            input(f"{datetime.now()} | 错误 | "
                  f"{re.search(r'[^日期 Date：][0-9][0-9][0-9][0-9][0-9][0-9][0-9]', source[i][10]).group()} "
                  f"结算单客户号不正确, 请检查后重试, 按任意键退出!\n")
            raise SystemExit()
        for j in range(len(source[i])):
            if re.match(r'\s*资金状况', source[i][j]):
                date = re.search(r'[^日期 Date：][0-9][0-9][0-9][0-9][0-9][0-9][0-9]', source[i][10]).group()
                balance_bf = float(source[i][j + 4][17:45].strip())
                deposit_withdraw = float(source[i][j + 6][25:47].strip())
                realized_pl = float(source[i][j + 8][18:46].strip())
                mtm_pl = float(source[i][j + 10][15:44].strip())
                commission = float(source[i][j + 14][17:46].strip())
                balance_cf = float(source[i][j + 6][65:-1].strip())
                client_equity = float(source[i][j + 10][65:-1].strip())
                margin_occupied = float(source[i][j + 14][70:-1].strip())
                data = {'日期': pd.to_datetime(date), '期初结存': balance_bf, '出入金': deposit_withdraw,
                        '平仓盈亏': realized_pl, '持仓盯市盈亏': mtm_pl, '手续费': commission, '期末结存': balance_cf,
                        '客户权益': client_equity, '保证金占用': margin_occupied, '风险度': margin_occupied / client_equity}
                df = pd.DataFrame([data])
                account = pd.concat([account, df], ignore_index=True)
        for j in range(len(source[i])):
            if re.match(r'\s*成交记录 Transaction Record', source[i][j]):
                for ldx in range(j + 10, len(source[i])):
                    if re.match(r'^-', source[i][ldx]):
                        break
                    if re.match(r'^\n', source[i][ldx]):
                        continue
                    row = sep.split(source[i][ldx])
                    data = {'成交日期': [pd.to_datetime(row[1])], '交易所': [row[2]], '品种': [row[3]], '合约': [row[4]],
                            '买/卖': [row[5]], '投/保': [row[6]], '成交价': [float(row[7])], '手数': [int(row[8])],
                            '成交额': [float(row[9])], '开/平': [row[10]], '手续费': [float(row[11])],
                            '平仓盈亏': [float(row[12])], '权利金收支': [float(row[13])], '成交序号': [int(row[14])]}
                    df = pd.DataFrame(data)
                    transaction_record = pd.concat([transaction_record, df], ignore_index=True)
        for j in range(len(source[i])):
            if re.match(r'\s*平仓明细 Position Closed', source[i][j]):
                for ldx in range(j + 10, len(source[i])):
                    if re.match(r'^-', source[i][ldx]):
                        break
                    if re.match(r'^\n', source[i][ldx]):
                        continue
                    row = sep.split(source[i][ldx])
                    price_margin = float(row[10]) - float(row[8]) if row[6] == '卖' else float(row[8]) - float(row[10])
                    data = {'平仓日期': [pd.to_datetime(row[1])], '交易所': [row[2]], '品种': [row[3]], '合约': [row[4]],
                            '开仓日期': [pd.to_datetime(row[5])], '买/卖': [row[6]], '手数': [int(row[7])],
                            '开仓价': [float(row[8])], '昨结算': [float(row[9])], '成交价': [float(row[10])],
                            '平仓盈亏': [float(row[11])], '权利金收支': [float(row[12])],
                            '交易盈亏': [price_margin * int(row[7]) *
                                     TRADING_UNIT[re.sub(r'[^A-Za-z]', '', row[4]).lower()]],
                            '盈亏率': [price_margin / float(row[8])],
                            }
                    df = pd.DataFrame(data)
                    position_closed = pd.concat([position_closed, df], ignore_index=True)
                    position_closed['持仓天数'] = position_closed['平仓日期'] - position_closed['开仓日期']
                    position_closed['持仓天数'].apply(lambda x: x.days)
    print(f'{datetime.now()} | 信息 | 已提取所有结算单数据')

    return client_id, account, transaction_record, position_closed


# ---------------------------------------------------- 数据提取 结束 ----------------------------------------------------


# ---------------------------------------------------- 数据统计 开始 ----------------------------------------------------
def net_worth_calc(account):
    print(f'{datetime.now()} | 信息 | 开始净值化处理')
    net_worth = pd.DataFrame(columns=['日期', '总权益', '净值', '份额', '份额变动'])
    net_worth['日期'] = account['日期']
    df = pd.merge(account, net_worth, on=['日期'])
    df.loc[0, '净值'] = 1.0
    df['总权益'] = df['客户权益']
    if df['期初结存'].iloc[0] != 0:
        df.loc[0, '份额'] = (df.iloc[0]['期初结存'] + df.iloc[0]['出入金']) / df.iloc[0]['净值']
    else:
        df.loc[0, '份额'] = df.iloc[0]['出入金'] / df.iloc[0]['净值']
    df.loc[0, '份额变动'] = df.iloc[0]['份额']
    for index in range(1, len(df.index)):
        if df.iloc[index]['出入金'] != 0:
            df.loc[index, '净值'] = (df.iloc[index]['总权益'] - df.iloc[index]['出入金']) / df.iloc[index - 1]['份额']
            df.loc[index, '份额变动'] = df.iloc[index]['出入金'] / df.iloc[index]['净值']
            df.loc[index, '份额'] = df.iloc[index - 1]['份额'] + df.iloc[index]['份额变动']
        else:
            df.loc[index, '净值'] = df.iloc[index]['总权益'] / df.iloc[index - 1]['份额']
            df.loc[index, '份额'] = df.iloc[index - 1]['份额']
    df['份额变动'].fillna(0, inplace=True)
    net_worth = pd.DataFrame(df, columns=net_worth.columns)
    print(f'{datetime.now()} | 信息 | 已完成净值化处理')
    return net_worth


def data_statistic(transaction_record, position_closed):
    print(f'{datetime.now()} | 信息 | 开始数据统计')
    statistic_by_contracts = pd.DataFrame(columns=['品种', '合约', '平仓盈亏', '净利润', '交易次数', '交易手数', '盈利次数',
                                                   '盈利手数', '交易成功率', '交易盈亏率', '均次盈亏', '均手盈亏', '最大盈利',
                                                   '最大亏损', '成交额'])
    statistic_by_contracts['合约'] = position_closed.groupby('合约').groups
    statistic_by_contracts = statistic_by_contracts.set_index('合约')
    for index in statistic_by_contracts.index:
        statistic_by_contracts.loc[index]['品种'] = CONTRACT_CODE[re.sub(r'[^A-Za-z]', '', index).lower()]
    position_closed_group_by_contracts = position_closed.groupby('合约')
    statistic_by_contracts['平仓盈亏'] = position_closed_group_by_contracts['交易盈亏'].sum().astype('float64')
    statistic_by_contracts['净利润'] = position_closed_group_by_contracts['交易盈亏'].sum().astype('float64') - \
                                    transaction_record.groupby('合约')['手续费'].sum().astype('float64')
    statistic_by_contracts['交易次数'] = position_closed_group_by_contracts['品种'].count().astype('int64')
    statistic_by_contracts['交易手数'] = position_closed_group_by_contracts['手数'].sum().astype('int64')
    statistic_by_contracts['盈利次数'] = position_closed_group_by_contracts.apply(lambda x: sum(x['交易盈亏'] > 0))
    statistic_by_contracts['盈利手数'] = position_closed_group_by_contracts.apply(lambda x: sum(x[x['交易盈亏'] > 0]['手数']))
    statistic_by_contracts['交易成功率'] = round(statistic_by_contracts['盈利次数'] / statistic_by_contracts['交易次数'], 4)
    statistic_by_contracts['交易盈亏率'] = round(statistic_by_contracts['盈利手数'] / statistic_by_contracts['交易手数'], 4)
    statistic_by_contracts['均次盈亏'] = round(statistic_by_contracts['平仓盈亏'] / statistic_by_contracts['交易次数'], 2)
    statistic_by_contracts['均手盈亏'] = round(statistic_by_contracts['平仓盈亏'] / statistic_by_contracts['交易手数'], 2)
    statistic_by_contracts['最大盈利'] = position_closed_group_by_contracts.apply(lambda x: max(x['交易盈亏']))
    statistic_by_contracts['最大亏损'] = position_closed_group_by_contracts.apply(lambda x: min(x['交易盈亏']))
    statistic_by_contracts['成交额'] = transaction_record.groupby('合约')['成交额'].sum().astype('float64')
    statistic_by_contracts = statistic_by_contracts.reset_index()

    statistic_by_categories = pd.DataFrame(columns=['品种', '平仓盈亏', '净利润', '交易次数', '交易手数', '盈利次数', '盈利手数',
                                                    '交易成功率', '交易盈亏率', '均次盈亏', '均手盈亏', '最大盈利', '最大亏损', '成交额'])
    contracts_analysis_group_by_categories = statistic_by_contracts.groupby('品种')
    statistic_by_categories['品种'] = contracts_analysis_group_by_categories.groups
    statistic_by_categories = statistic_by_categories.set_index('品种')
    statistic_by_categories['平仓盈亏'] = contracts_analysis_group_by_categories['平仓盈亏'].sum()
    statistic_by_categories['净利润'] = contracts_analysis_group_by_categories['净利润'].sum()
    statistic_by_categories['交易次数'] = contracts_analysis_group_by_categories['交易次数'].sum()
    statistic_by_categories['交易手数'] = contracts_analysis_group_by_categories['交易手数'].sum()
    statistic_by_categories['盈利次数'] = contracts_analysis_group_by_categories['盈利次数'].sum()
    statistic_by_categories['盈利手数'] = contracts_analysis_group_by_categories['盈利手数'].sum()
    statistic_by_categories['交易成功率'] = round(statistic_by_categories['盈利次数'] / statistic_by_categories['交易次数'], 4)
    statistic_by_categories['交易盈亏率'] = round(statistic_by_categories['盈利手数'] / statistic_by_categories['交易手数'], 4)
    statistic_by_categories['均次盈亏'] = round(statistic_by_categories['平仓盈亏'] / statistic_by_categories['交易次数'], 2)
    statistic_by_categories['均手盈亏'] = round(statistic_by_categories['平仓盈亏'] / statistic_by_categories['交易手数'], 2)
    statistic_by_categories['最大盈利'] = contracts_analysis_group_by_categories.apply(lambda x: max(x['最大盈利']))
    statistic_by_categories['最大亏损'] = contracts_analysis_group_by_categories.apply(lambda x: min(x['最大亏损']))
    statistic_by_categories['成交额'] = contracts_analysis_group_by_categories['成交额'].sum()
    statistic_by_categories = statistic_by_categories.reset_index()

    statistic_by_trade_direction = pd.DataFrame(columns=['统计指标', '总盈利', '总亏损', '总盈利/总亏损', '手续费', '净利润',
                                                         '盈利比率', '盈利手数', '亏损手数', '持平手数', '平均盈利', '平均亏损',
                                                         '平均盈利/平均亏损', '平均手续费', '平均净利润', '最大盈利', '最大亏损',
                                                         '最大盈利/总盈利', '最大亏损/总亏损', '净利润/最大亏损'])
    position_closed_group_by_trade_direction = position_closed.groupby('买/卖')
    statistic_by_trade_direction['统计指标'] = position_closed_group_by_trade_direction.groups
    statistic_by_trade_direction = statistic_by_trade_direction.set_index('统计指标')
    statistic_by_trade_direction['总盈利'] = position_closed_group_by_trade_direction.apply(
        lambda x: sum(x[x['交易盈亏'] > 0]['交易盈亏']))
    statistic_by_trade_direction['总亏损'] = position_closed_group_by_trade_direction.apply(
        lambda x: sum(x[x['交易盈亏'] < 0]['交易盈亏']))
    statistic_by_trade_direction['总盈利/总亏损'] = round(
        abs(statistic_by_trade_direction['总盈利'] / statistic_by_trade_direction['总亏损']), 4)
    statistic_by_trade_direction['手续费'] = transaction_record.groupby('买/卖')['手续费'].sum().astype('float64')
    statistic_by_trade_direction['净利润'] = statistic_by_trade_direction['总盈利'] + statistic_by_trade_direction['总亏损'] - \
                                          transaction_record.groupby('买/卖')['手续费'].sum().astype('float64')
    statistic_by_trade_direction['盈利手数'] = position_closed_group_by_trade_direction.apply(
        lambda x: sum(x[x['交易盈亏'] > 0]['手数']))
    statistic_by_trade_direction['亏损手数'] = position_closed_group_by_trade_direction.apply(
        lambda x: sum(x[x['交易盈亏'] < 0]['手数']))
    statistic_by_trade_direction['持平手数'] = position_closed_group_by_trade_direction.apply(
        lambda x: sum(x[x['交易盈亏'] == 0]['手数']))
    statistic_by_trade_direction['盈利比率'] = round(statistic_by_trade_direction['盈利手数'] / (
            statistic_by_trade_direction['盈利手数'] + statistic_by_trade_direction['亏损手数'] +
            statistic_by_trade_direction['持平手数']), 4)
    statistic_by_trade_direction['平均盈利'] = round(
        statistic_by_trade_direction['总盈利'] / statistic_by_trade_direction['盈利手数'], 2)
    statistic_by_trade_direction['平均亏损'] = round(
        statistic_by_trade_direction['总亏损'] / statistic_by_trade_direction['亏损手数'], 2)
    statistic_by_trade_direction['平均盈利/平均亏损'] = round(
        abs(statistic_by_trade_direction['平均盈利'] / statistic_by_trade_direction['平均亏损']), 4)
    statistic_by_trade_direction['平均手续费'] = round(statistic_by_trade_direction['手续费'] / (
            statistic_by_trade_direction['盈利手数'] + statistic_by_trade_direction['亏损手数'] +
            statistic_by_trade_direction['持平手数']), 2)
    statistic_by_trade_direction['平均净利润'] = statistic_by_trade_direction['平均盈利'] + statistic_by_trade_direction[
        '平均亏损'] - statistic_by_trade_direction['平均手续费']
    statistic_by_trade_direction['最大盈利'] = position_closed_group_by_trade_direction['交易盈亏'].max()
    statistic_by_trade_direction['最大亏损'] = position_closed_group_by_trade_direction['交易盈亏'].min()
    statistic_by_trade_direction['最大盈利/总盈利'] = round(
        statistic_by_trade_direction['最大盈利'] / statistic_by_trade_direction['总盈利'], 4)
    statistic_by_trade_direction['最大亏损/总亏损'] = round(
        statistic_by_trade_direction['最大亏损'] / statistic_by_trade_direction['总亏损'], 4)
    statistic_by_trade_direction['净利润/最大亏损'] = round(
        abs(statistic_by_trade_direction['净利润'] / statistic_by_trade_direction['最大亏损']), 4)
    statistic_by_trade_direction = statistic_by_trade_direction.reset_index()
    print(f'{datetime.now()} | 信息 | 已完成数据统计')
    return statistic_by_contracts, statistic_by_categories, statistic_by_trade_direction


# ---------------------------------------------------- 数据统计 结束 ----------------------------------------------------


# --------------------------------------------------- 数据格式化 开始 ---------------------------------------------------
def excel_data_format(excel_file):
    def cell_format_by_columns(worksheet):
        """
        单元格根据列表头格式化
        :param worksheet:
        :return:
        """
        for column_dataset in worksheet.columns:
            header = column_dataset[0].value
            for row_index in range(worksheet.min_row, worksheet.max_row):
                if '日期' in header:
                    column_dataset[row_index].number_format = numbers.FORMAT_DATE_YYYYMMDD2
                elif '风险度' in header or '率' in header or '/' in header:
                    column_dataset[row_index].number_format = numbers.FORMAT_PERCENTAGE_00
                elif '利润' in header or '盈亏' in header or '结存' in header or '权益' in header or '保证金' in header \
                        or '出入金' in header or '成交额' in header:
                    column_dataset[row_index].number_format = numbers.BUILTIN_FORMATS[39]

    def cell_format_by_rows(worksheet):
        """
        单元格根据行表头格式化
        :param worksheet:
        :return:
        """
        for row_dataset in worksheet.rows:
            header = row_dataset[0].value
            for column_index in range(worksheet.min_column, worksheet.max_column):
                if '日期' in header:
                    row_dataset[column_index].number_format = numbers.FORMAT_DATE_YYYYMMDD2
                elif '风险度' in header or '率' in header or '/' in header:
                    row_dataset[column_index].number_format = numbers.FORMAT_PERCENTAGE_00
                elif '利润' in header or '盈亏' in header or '结存' in header or '权益' in header or '保证金' in header \
                        or '出入金' in header or '成交额' in header:
                    row_dataset[column_index].number_format = numbers.BUILTIN_FORMATS[39]

    def dimension_format(worksheet, dimension='columns'):
        """
        行列格式化
        :param worksheet:
        :param dimension:
        :return:
        """
        if dimension == 'columns':
            for column_index in get_column_interval(worksheet.min_column, worksheet.max_column):
                worksheet.column_dimensions[column_index].auto_size = True
                worksheet.column_dimensions[column_index].best_fit = True
        elif dimension == 'rows':
            for row_index in range(worksheet.min_row, worksheet.max_row):
                pass

    def data_transposition(workbook, sheet_name, index=False):
        """
        数据转置
        :param workbook:
        :param sheet_name:
        :param index:
        :return:
        """
        data = workbook[sheet_name].values
        if index is True:
            cols = next(data)[1:]
            data = list(data)
            index = [r[0] for r in data]
            data = (islice(r, 1, None) for r in data)
        else:
            cols = next(data)[0:]
            data = list(data)
            index = None
            data = (islice(r, 0, None) for r in data)
        df = pd.DataFrame(data, index=index, columns=cols).T
        df.reset_index(level=0, inplace=True)
        idx = workbook.sheetnames.index(sheet_name)
        workbook.remove(workbook.worksheets[idx])
        workbook.create_sheet(sheet_name, idx)
        for r in dataframe_to_rows(df, index=index, header=None):
            workbook[sheet_name].append(r)
        for cell in workbook[sheet_name]['A']:
            cell.style = 'Pandas'

    wb = load_workbook(excel_file)
    print(f'{datetime.now()} | 信息 | 开始Excel数据格式化')
    for ws_name in wb.sheetnames:
        cell_format_by_columns(wb[ws_name])
        dimension_format(wb[ws_name])
        if ws_name == '交易分析(按买卖)':
            data_transposition(wb, '交易分析(按买卖)')
            cell_format_by_rows(wb[ws_name])
            dimension_format(wb[ws_name])
    wb.save(excel_file)
    wb.close()
    print(f'{datetime.now()} | 信息 | 已生成Excel数据表')


# --------------------------------------------------- 数据格式化 结束 ---------------------------------------------------


# ---------------------------------------------------- 生成图表 开始 ----------------------------------------------------
def excel_create_chart(excel_file):
    """
    生成图表
    :param excel_file:
    :return:
    """
    wb = load_workbook(excel_file)
    print(f'{datetime.now()} | 信息 | 开始Excel图表渲染')
    # 进行净值走势图渲染
    net_worth_sheet = wb['账户净值']
    net_worth_chart_sheet = wb.create_chartsheet(title='净值走势图')
    chart1 = LineChart()
    dates1 = Reference(net_worth_sheet, min_col=1, min_row=2, max_row=len(net_worth_sheet['A']))
    data1 = Reference(net_worth_sheet, min_col=3, min_row=1, max_row=len(net_worth_sheet['C']))
    chart1.add_data(data1, titles_from_data=True)
    chart1.y_axis = NumericAxis(title='净值', majorTickMark='out')
    chart1.x_axis = TextAxis(majorTickMark='out', tickLblSkip=10, tickMarkSkip=10, noMultiLvlLbl=True,
                             numFmt='yyyy-mm-dd')
    chart1.legend = None
    chart1.title = str()
    chart1.style = 1
    chart1.set_categories(dates1)
    net_worth_chart_sheet.add_chart(chart1)
    # 进行权益走势图渲染
    account_sheet = wb['账户统计']
    account_chart_sheet = wb.create_chartsheet(title='权益走势图')
    dates2 = Reference(account_sheet, min_col=1, min_row=2, max_row=len(account_sheet['A']))
    chart2 = LineChart(varyColors=True)
    data2 = Reference(account_sheet, min_col=8, min_row=1, max_row=len(account_sheet['H']))
    chart2.add_data(data2, titles_from_data=True)
    chart2.y_axis = NumericAxis(title='权益', majorTickMark='out')
    chart2.legend = None
    chart2.set_categories(dates2)
    chart3 = BarChart()
    data3 = Reference(account_sheet, min_col=10, min_row=1, max_row=len(account_sheet['J']))
    chart3.add_data(data3, titles_from_data=True)
    chart3.y_axis = NumericAxis(axId=200, title='风险度', majorGridlines=None, majorTickMark='out', crosses='max')
    chart3.x_axis = TextAxis(majorTickMark='out', tickLblSkip=10, tickMarkSkip=10, noMultiLvlLbl=True,
                             numFmt='yyyy-mm-dd')
    chart3.legend = None
    chart3.set_categories(dates2)
    chart2 += chart3
    account_chart_sheet.add_chart(chart2)
    # 进行交易分布图的渲染
    categories_analysis_sheet = wb['交易分析(按品种)']
    trading_frequency_analysis_sheet = wb.create_chartsheet(title='交易分布图')
    labels = Reference(categories_analysis_sheet, min_col=1, min_row=2, max_row=len(categories_analysis_sheet['A']))
    chart4 = PieChart(varyColors=True)
    chart4.style = 34
    data4 = Reference(categories_analysis_sheet, min_col=4, min_row=2, max_row=len(categories_analysis_sheet['D']))
    chart4.add_data(data4)
    chart4.set_categories(labels)
    chart4.legend = None
    # chart4.series[0].data_points = [DataPoint(idx=i, explosion=8)
    #                                 for i in range(len(categories_analysis_sheet['D']) - 1)]
    chart4.series[0].dLbls = DataLabelList(dLblPos='bestFit', showPercent=True, showCatName=True, showVal=True,
                                           showLeaderLines=True)
    chart4.layout = Layout(manualLayout=ManualLayout(x=0, y=0, h=0.75, w=0.75, xMode='factor', yMode='factor'))
    trading_frequency_analysis_sheet.add_chart(chart4)
    # 进行品种盈亏图的渲染
    categories_win_and_loss_chart_sheet = wb.create_chartsheet(title='品种盈亏图')
    chart5 = BarChart(barDir='col')
    chart5.style = 18
    data5 = Reference(categories_analysis_sheet, min_col=2, min_row=2, max_row=len(categories_analysis_sheet['B']))
    chart5.add_data(data5)
    chart5.set_categories(labels)
    chart5.legend = None
    chart5.series[0].dLbls = DataLabelList(showVal=True)
    chart5.y_axis = NumericAxis(title='平仓盈亏', majorTickMark='out', minorTickMark='out')
    categories_win_and_loss_chart_sheet.add_chart(chart5)
    # 进行交易盈亏图的渲染
    trading_win_and_loss_chart_sheet = wb.create_chartsheet(title='交易盈亏图')
    chart6 = RadarChart()
    chart6.style = 24
    data6 = Reference(categories_analysis_sheet, min_col=8, max_col=9, min_row=1,
                      max_row=categories_analysis_sheet.max_row)
    chart6.add_data(data6, titles_from_data=True)
    chart6.set_categories(labels)
    trading_win_and_loss_chart_sheet.add_chart(chart6)
    # 图表保存
    wb.save(excel_file)
    wb.close()
    # 输出信息
    print(f'{datetime.now()} | 信息 | 已生成Excel图表')


# ---------------------------------------------------- 生成图表 结束 ----------------------------------------------------


# -------------------------------------------------- 生成excel文件 开始 -------------------------------------------------
def output_excel(net_worth, account, transaction_record, position_closed, contracts_analysis, categories_analysis,
                 trade_direction_analysis, client_id=''):
    try:
        with pd.ExcelWriter(os.path.join(BASE_DIR, client_id + '交易统计.xlsx')) as writer:
            net_worth.to_excel(writer, sheet_name='账户净值', encoding='ansi', index=False)
            account.to_excel(writer, sheet_name='账户统计', encoding='ansi', index=False)
            transaction_record.to_excel(writer, sheet_name='交易记录', encoding='ansi', index=False)
            position_closed.to_excel(writer, sheet_name='平仓明细', encoding='ansi', index=False)
            contracts_analysis.to_excel(writer, sheet_name='交易分析(按合约)', encoding='ansi', index=False)
            categories_analysis.to_excel(writer, sheet_name='交易分析(按品种)', encoding='ansi', index=False)
            trade_direction_analysis.to_excel(writer, sheet_name='交易分析(按买卖)', encoding='ansi', index=False)
        excel_data_format(os.path.join(BASE_DIR, client_id + '交易统计.xlsx'))
        excel_create_chart(os.path.join(BASE_DIR, client_id + '交易统计.xlsx'))
        input(f'{datetime.now()} | 信息 | 任务结束, 感谢您的使用, 按任意键退出!\n')
        # raise SystemExit()
    except PermissionError:
        input(f'{datetime.now()} | 错误 | 分析结果写入Excel被拒绝, 请检查文件是否已打开, 按任意键退出!\n')
        # raise SystemExit()


# -------------------------------------------------- 生成excel文件 结束 -------------------------------------------------


# ---------------------------------------------------- 终端命令 开始 ----------------------------------------------------
def main(argv):
    client_id = ''
    files_folder = ''
    try:
        opts, args = getopt.getopt(argv, "hd:i:", ["dir=", "id="])
    except getopt.GetoptError:
        print('参数选项:\n-d/--dir <settlement statement files\' folder  结算单文件夹路径>\n-i/--id <client id  客户号>')
        sys.exit(2)
    if len(opts) != 0:
        for opt, arg in opts:
            if opt == '-h':
                print('参数选项:\n-d/--dir <settlement statement files\' folder  结算单文件夹路径>\n-i/--id <client id  客户号>')
                sys.exit()
            elif opt in ("-d", "--dir"):
                files_folder = arg
            elif opt in ("-i", "--id"):
                client_id = arg
    statement_list = read_statement_files(files_folder)
    client_id, account, transaction_record, position_closed = data_extract(statement_list, client_id=client_id)
    net_worth = net_worth_calc(account)
    contracts_analysis, categories_analysis, trade_direction_analysis = data_statistic(transaction_record,
                                                                                       position_closed)
    output_excel(net_worth, account, transaction_record, position_closed, contracts_analysis, categories_analysis,
                 trade_direction_analysis, client_id=client_id)


if __name__ == '__main__':
    main(sys.argv[1:])
# ---------------------------------------------------- 终端命令 结束 ----------------------------------------------------
