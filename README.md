# Tridro量化工具集
Tridro的量化工具集，方便程序化交易及交易分析工作
<p align="left">
    <img src ="https://img.shields.io/badge/platform-windows|linux|-green.svg" />
    <img src ="https://img.shields.io/badge/python-3.7+-blue.svg" />
    <img src ="https://img.shields.io/badge/license-Apache2.0-orange" />
</p>

## 期货行情数据源API（新浪）（不进行更新）
  
## 期货结算单数据分析
最近更新: 2022/10/11
1. 更新正则规则以识别2022年9月23日以后的最新结算单格式
### 使用方式
* 安装python，并安装所依赖库
``` {.sourceCode .bash}
$ pip install pandas openpyxl
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
