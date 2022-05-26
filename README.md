# Tridro量化工具集
Tridro的量化工具集，方便程序化交易及交易分析工作
<p align="left">
    <img src ="https://img.shields.io/badge/platform-windows|linux|-green.svg" />
    <img src ="https://img.shields.io/badge/python-3.7+-blue.svg" />
    <img src ="https://img.shields.io/badge/license-Apache2.0-orange" />
</p>

## 期货行情数据源API（新浪）（不进行更新）
  
## 期货结算单数据分析
最近更新: 2022/5/26
1. 修正了之前的净值计算方式，出入金根据交易日盘后净值计算份额
### 使用方式
* 安装python，并安装所依赖库
``` {.sourceCode .bash}
$ pip install pandas openpyxl
```
* 直接运行或命令行运行，会在脚本同路径目录生成excel分析结果。
``` {.sourceCode .bash}
$ python futures_trading_statement_analysis.py -d/--dir <完整结算单文件夹路径/当前路径文件夹名> -i/--id <客户号>
```
