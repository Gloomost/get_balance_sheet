import pdfplumber
import pandas as pd
import re, os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.font_manager as fm


# 显示所有列
pd.set_option('display.max_columns', None)
# 显示所有行
pd.set_option('display.max_rows', None)


def to_pdf(df, name):
    '''
    字体大小，已经过测试，字号为6偏大，5~5.5大小适中
    还未对其他年报进行测试，可能存在项目名称过长导致字体溢出表格或页面的问题
    '''
    fontsize = 5.5
    # df.replace('0', '')
    # 设置 matplotlib 的字体为 SimHei
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

    # 创建保存目录
    if not os.path.exists('资产负债表'):
        os.mkdir('资产负债表')

    with PdfPages('资产负债表/' + name) as pdf:
        # 创建一个表格
        fig, ax = plt.subplots()
        ax.axis('off')  # 关闭表格的坐标轴

        # 创建表格
        table = ax.table(cellText=df.values, colLabels=df.columns,
                         loc='center', cellLoc='center')

        # 设置字体大小
        table.auto_set_font_size(False)  # 关闭自动字体调整
        table.set_fontsize(fontsize)  # 设置表格字体大小
        table.scale(1, 1.5)  # 通过缩放表格调整布局

        # 遍历表格的所有单元格并设置字体大小
        for key, cell in table.get_celld().items():
            cell.set_fontsize(fontsize)

        pdf.savefig(fig, bbox_inches='tight')  # 将表格保存到PDF中
        plt.close()

    print(f"资产负债表文件已生成：{name}.pdf")


def swap_year(df):
    try:
        # 检查 "1月1日" 是否在 df.iloc[0, -1] 或 df.iloc[0, -2] 中
        if "1月1日" in str(df.iloc[0, -1]):
            df.iloc[0, -1] = (str(int(df.iloc[0, -1].split("年")[0]) - 1)
                              + "年12月31日")
        if "1月1日" in str(df.iloc[0, -2]):
            df.iloc[0, -2] = (str(int(df.iloc[0, -2].split("年")[0]) - 1)
                              + "年12月31日")

        # print(df.iloc[0, -1], df.iloc[0, -2])
        # 检查是否需要交换年份
        year1 = int(df.iloc[0, -2].split("年")[0])
        # print(year1)
        year2 = int(df.iloc[0, -1].split("年")[0])
        # print(year2)

        if year1 < year2:
            df.iloc[0, -2], df.iloc[0, -1] = df.iloc[0, -1], df.iloc[0, -2]
        new_header = df.iloc[0]  # 将第一行设置为表头
        df = df[1:]  # 从第二行开始作为数据行
        df.columns = new_header  # 设置表头
        return df
    except Exception as e:
        # 将错误信息和文件名保存到错误列表
        return str(e)


def judge_line(line):
    # print(line)
    before_word = 0
    newLine = []
    for word in line:
        if word is None: word = ''
        if len(word) == 0:
            before_word += 1
        else:
            newLine = line[1 + 2 * before_word:]
            project = word.replace("\n", "")
            break
    newLine = ['' if x is None else x for x in newLine]
    # print('newLine')
    # print(newLine)
    k = [i for i in range(len(newLine)) if len(newLine[i]) > 0]
    if len(k) == 2:
        return [project, newLine[k[0]], newLine[k[1]]]
    elif len(k) == 0:
        return [project, 0, 0]
    elif len(k) == 1:
        if k[0] + 1 <= len(newLine)/2:
            return [project, newLine[k[0]], 0]
        elif k[0] + 1 > len(newLine)/2:
            return [project, 0, newLine[k[0]]]


def judge_line_with_fuzhu(line):
    # print(line)
    before_word = 0
    newLine = []
    for word in line:
        if word is None: word = ''
        if len(word) == 0:
            before_word += 1
        else:
            newLine = line[2 + 2 * before_word:]
            project = word.replace("\n", "")
            break
    newLine = ['' if x is None else x for x in newLine]
    # print('newLine')
    # print(newLine)
    k = [i for i in range(len(newLine)) if len(newLine[i]) > 0]
    if len(k) == 2:
        return [project, newLine[k[0]], newLine[k[1]]]
    elif len(k) == 0:
        return [project, 0, 0]
    elif len(k) == 1:
        if k[0] + 1 <= len(newLine)/2:
            return [project, newLine[k[0]], 0]
        elif k[0] + 1 > len(newLine)/2:
            return [project, 0, newLine[k[0]]]


# 将pdf表格数据抽取到文件中
def extract_tables(input_file_path, errors):
    print("========================================资产负债表抽取开始========================================")
    # 读取pdf文件，保存为pdf实例
    pdf = pdfplumber.open(input_file_path)
    # 存储所有页面内的表格文本
    text_all_table = []
    # 访问每一页
    print("1===========开始抽取每页表格文本===========1")
    for page_num, page in enumerate(pdf.pages):
        # table对象，可以访问其row属性的bbox对象获取坐标
        # table_objects = page.find_tables()
        text_table_current_page = page.extract_tables()
        if text_table_current_page:
            # 获取有表格的页面数
            text_all_table.append(text_table_current_page)
    print("1===========抽取每页表格文本结束===========1")


    start_pattern = "20.*年.*月.*日"
    end_pattern = "负债和所有者权益"
    mark = 0
    fuzhu = 0
    # 保存excel
    print("2===========开始保存表格===========2")
    for table in text_all_table:
        if mark == 1:
            # print(table[0])
            for i in range(len(table[0])):
                # print(table[0][i])
                # print(pd.DataFrame(judgeThisLine(table[0][i])))
                new_table = (pd.concat
                             ([new_table, pd.DataFrame([judge_line(table[0][i])])],
                              axis=0))
            # print(table[0][-1])
            for x in table[0][-1]:
                if x is None: continue
                if end_pattern in x:
                    # print(new_table)
                    df = swap_year(new_table)
                    # print(df)
                    if isinstance(df, str):
                        errors.append((input_file_path.split('\\')[-1], df))
                    else:
                        to_pdf(df, input_file_path.split('\\')[-1])
                    return
        if fuzhu == 1:
            # print(table[0])
            for i in range(len(table[0])):
                # print(table[0][i])
                # print(pd.DataFrame(judgeThisLine(table[0][i])))
                new_table = pd.concat(
                    [new_table, pd.DataFrame([judge_line_with_fuzhu(table[0][i])])],
                    axis=0)
            # print(table[0][-1])
            for x in table[0][-1]:
                if x is None: continue
                if end_pattern in x:
                    # print(new_table)
                    df = swap_year(new_table)
                    # print(df)
                    if isinstance(df, str):
                        errors.append((input_file_path.split('\\')[-1], df))
                    else:
                        to_pdf(df, input_file_path.split('\\')[-1])
                    return


        if table and len(table[0][0]) > 2:
            # 确保该表表头一定是第一列为 项目，倒数第一列和倒数第二列模糊搜索均为 20 年 月 日（按顺序）
            table[0][0] = [item for item in table[0][0] if item != '']
            table[0][0] = [item for item in table[0][0] if item != None]
            for i in range(len(table[0][0])):
                table[0][0][i] = table[0][0][i].replace("\n", "")
            # print(table[0][0])
            if (table[0][0] and table[0][0][0] == '项目'
                and (isinstance(table[0][0][-1], str)
                and isinstance(table[0][0][-2], str)
                and re.search(start_pattern, table[0][0][-1])
                and re.search(start_pattern, table[0][0][-2]))):
                # print(table[0][0])
                new_table = pd.DataFrame()
                # print(table[0])
                if '附注' not in table[0][0]:
                    mark = 1
                    for i in range(len(table[0])):
                        # print(table[0][i])
                        # print(pd.DataFrame(judgeThisLine(table[0][i])))
                        new_table = pd.concat(
                            [new_table, pd.DataFrame([judge_line(table[0][i])])],
                            axis=0)
                else:
                    fuzhu = 1
                    for i in range(len(table[0])):
                        # print(table[0][i])
                        # print(pd.DataFrame(judgeThisLine(table[0][i])))
                        new_table = pd.concat(
                            [new_table, pd.DataFrame([judge_line_with_fuzhu(table[0][i])])],
                            axis=0)
                    # print(table_df)
    print("2===========保存表格结束===========2")


def batch(first, last):
    # 抽取表格
    # first = 0
    # last = 10
    errors = []
    for filepath, dirnames, filenames in os.walk('../../sd_reports'):
        k = first
        for filename in filenames[first:last + 1]:
            print('k:', k)
            k += 1
            print(filename)
            input_file = os.path.join(filepath, filename)
            extract_tables(input_file, errors)
    error_log_path = f'error_log_pdf_{first}-{last}.txt'
    with open(error_log_path, 'w', encoding='utf-8') as f:
        for error in errors:
            f.write(f"文件: {error[0]} - 错误: {error[1]}\n")


def one_test(input_file):
    errors = []
    extract_tables(input_file, errors)
    if len(errors) != 0:
        print(f"{errors[0]}文件未能成功提取资产负债表的错误是:\n{errors[1]}")


if __name__ == '__main__':
    '''
    batch函数中输入变量为start和last，目的是可以分开提取表格，两者是通过os遍历出的文件list的下标
    test函数中可以单独输入文件路径，提取单个文件的表格，注意，输入路径中，文件名称前的分隔符为两个反斜杠\\
    '''
    # batch(0, 10)
    one_test("../../sd_reports\\_ST亚星潍坊亚星化学股份有限公司2021年年度报告（修订版）.PDF")
