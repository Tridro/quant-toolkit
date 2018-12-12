import requests
import sys
import time
import csv


class Futures_data_sina():
    '''
    An API script request data of China's futures market from sina.com
    Update data: 30/11/2018

    request(): request a specific contract data of China's futures markets from sina.com
    show(): form a dataframe like list and display properly..
    output_csv(): output data, create a csv file in working directory with symbol, time period and data acquisition time.
    '''
    def __init__(self):
        # self.timeperiod = ['5m', '15m', '30m', '60m', '1d']
        self.commodity_future_code = ['au', 'ag', 'cu', 'al', 'zn', 'pb', 'ni', 'sn', 'rb', 'i', 'hc', 'wr', 'sf', 'sm',
                                      'fg', 'sp', 'fb', 'bb', 'jm', 'j', 'zc', 'fu', 'sc', 'ru', 'l', 'ta', 'v', 'eg', 'ma',
                                      'pp', 'bu', 'c', 'a', 'cs', 'wh', 'ri', 'jr', 'lr', 'b', 'm', 'y', 'rs', 'rm',
                                      'oi', 'p', 'cf', 'sr', 'cy', 'jd', 'ap']
        self.index_future_code = ['if', 'ih', 'ic', 'ts', 'tf', 't']

    def request(self, future_code, future_timeperiod):
        global data
        try:
            self.future_timeperiod = future_timeperiod
            self.future_code = future_code
            self.request_time = time.localtime()

            for code_prefix in self.commodity_future_code:
                if future_code.find(code_prefix) != -1:
                    if future_timeperiod == '1d':
                        url_str = (
                                'http://stock2.finance.sina.com.cn/futures/api/json.php/IndexService'
                                '.getInnerFuturesDailyKLine?symbol=' + future_code)
                        data = requests.get(url_str)
                    elif future_timeperiod == ('5m' or '15m' or '30m' or '60m'):
                        url_str = (
                                'http://stock2.finance.sina.com.cn/futures/api/json.php/IndexService'
                                '.getInnerFuturesMiniKLine' + future_timeperiod + '?symbol=' + future_code)
                        data = requests.get(url_str)

            for code_prefix in self.index_future_code:
                if future_code.find(code_prefix) != -1:
                    if future_timeperiod == '1d':
                        url_str = (
                                'http://stock2.finance.sina.com.cn/futures/api/json.php/CffexFuturesService'
                                '.getCffexFuturesDailyKLine?symbol=' + future_code)
                        data = requests.get(url_str)
                    elif future_timeperiod == ('5m' or '15m' or '30m' or '60m'):
                        url_str = (
                                'http://stock2.finance.sina.com.cn/futures/api/json.php/CffexFuturesService'
                                '.getCffexFuturesMiniKLine?' + future_timeperiod + '?symbol=' + future_code)
                        data = requests.get(url_str)

            data_json = data.json()
            self.data_lists = list(data_json)

            assert self.data_lists

        except (AttributeError, SyntaxError) as err:
            print('Input Error, Data Time period Support: 5m, 15m, 30m, 60m, 1d')
        else:
            print('Data Acquired ' + str(data))
            return self.data_lists

    def show(self):
        print('date, open, high, low, close, vol, code')
        for data_set in self.data_lists:
            for data in data_set:
                print(data + ',', end='')
            print(self.future_code)

    def output_csv(self):
        t = time.strftime('%Y%m%d%H%M%S', self.request_time)
        file_name = str(self.future_code + '-' + self.future_timeperiod + '-' + t)
        data_lists = self.data_lists
        with open('%s.csv' % file_name, 'w', encoding='utf8', newline='') as f:
            csv_writer = csv.writer(f)
            csv_writer.writerow(['date', 'open', 'high', 'low', 'close', 'vol', 'code'])
            for data_set in data_lists:
                data_set = data_set.append(self.future_code)
            csv_writer.writerows(data_lists)


if __name__ == '__main__':
    # Terminal directly invoking
    future_code = sys.argv[1]
    future_timeperiod = sys.argv[2]
    temp = Futures_data_sina()
    temp.request(future_code, future_timeperiod)
    temp.output_csv()
