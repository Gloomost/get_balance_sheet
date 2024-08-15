# get_balance_sheet
### 在利用大模型对上市公司短期偿债能力进行分析时，第一步就是在公司年报中读取相应指标（货币资金、流动负债合计等），完成现金比率、流动比率、速动比率的计算。但在实际完成过程中，对这些数据的提取会出现多种问题：
### 1.流动负债合计与其他指标放在一起查询时会导致所有指标均查询出错
### 2.很多公司的年报中有合并资产负债表和母公司资产负债表两张表，且这两张表属于跨页表格，大模型很难分清这两张表，查询时会将两张表中的数据混淆
### 3.资产负债表中的两列并不全是今年在前，去年在后，同时在大部分表头都是xx年12月31日时，也有部分表会出现xx年1月1日的表头，这导致大模型很难分清要查询的年份
### 为了解决以上问题，在尝试多种方案后，通过使用python中的pdfplumber库对pdf中的表格进行提取并对表格进行处理，最终完成了大部分年报中资产负债表的提取，截至目前，本项目已测试54个年报，均可正确完整地提取出资产负债表。
## 使用方法如下：
### 项目中有两个函数可供选择，分别是batch函数和one函数。batch函数中输入变量为start和last，目的是可以批量提取表格，两者是通过os遍历出的文件list的下标，通过修改这两个变量可以控制要提取资产负债表的数量及范围；test函数中可以单独输入文件路径，提取单个文件的表格，方便对单个年报进行资产负债表的提取，注意，输入路径中，文件名称前的分隔符为两个反斜杠\\
## 注意：
### 该项目只对年报中资产负债表的提取有效，尚未实现对年报中其他表格或者其他pdf表格中跨页表格的一般化提取
    
