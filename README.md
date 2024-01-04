# Tridro量化工具集
Tridro的量化工具集，方便程序化交易及交易分析工作
<p align="left">
    <img src ="https://img.shields.io/badge/platform-windows|linux|-green.svg" />
    <img src ="https://img.shields.io/badge/python-3.7+-blue.svg" />
    <img src ="https://img.shields.io/badge/license-Apache2.0-orange" />
</p>

## 期货行情数据源API（新浪）（不进行更新）
  
## 期货结算单数据分析
最近更新: 2023/12/28
1. Fix bugs about skew and kurtosis function. 修改了峰度和偏度的函数计算问题
### 使用方式
* 安装python，并安装所依赖库
``` {.sourceCode .bash}
$ pip install numpy pandas openpyxl
```
* 直接运行或命令行运行，会在脚本同路径目录生成excel分析结果。
``` {.sourceCode .bash}
$ python futures_trading_statement_analysis.py -d/--dir <完整结算单文件夹路径/当前路径文件夹名> -i/--id <客户号>
```
## 期货跨月偏离度监测
最近更新: 2022/6/6
1. 首次上传脚本
### 使用方式
* 安装python，并安装所依赖库
``` {.sourceCode .bash}
$ pip install numpy tqsdk matplotlib
```
* 直接运行, 输入天勤量化账号和密码后，指定期货品种，会生成一个matplotlib窗口，动态监控当前品种跨月结构，每隔0.5秒刷新。
## 自动计算保证金需求
最近更新: 2023/6/29
1. 首次上传脚本
### 使用方式
* 安装python，并安装所依赖库
``` {.sourceCode .bash}
$ pip install pandas tqsdk openpyxl
```
* 直接运行, 输入天勤量化账号和密码后，会读取本地的保证金比率表<margin_ratio.xlsx>, 根据时间计算当日或者判断下一个交易日的合约保证金需求，并生成保证金需求数据表。
