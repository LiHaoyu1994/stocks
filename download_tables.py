"""
@encoding:utf-8
@author:Tommy
@time:2020/10/9　21:40
@note:
@备注:
"""
import re, urllib
import xlwt
from bs4 import BeautifulSoup
from time import sleep
import os
import pandas as pd
import numpy as np

# 新建汇总表
workbook = xlwt.Workbook(encoding='gb18030')


# 从txt文件获取待查询"{} {}".format(股票代码,股票名称)的格式
def get_isbn_from_txt(file_name: str) -> dict:
    result = []
    # 先把文件中的全部信息录入data_total中.
    fp = open(file_name, 'r', encoding='gbk')
    words = fp.readline()
    while len(words) > 0:
        if len(words.strip()) > 0:
            # 有的首行有\ufeff,需要清除
            result.append(words.replace("\ufeff", "").strip())
        words = fp.readline()
    return result


# 清空文件夹及下面所有文件
def del_file(path_data):
    for i in os.listdir(path_data):  # os.listdir(path_data)#返回一个列表，里面是当前目录下面的所有东西的相对路径
        file_data = path_data + "\\" + i  # 当前文件夹的下面的所有东西的绝对路径
        if os.path.isfile(file_data) == True:  # os.path.isfile判断是否为文件,如果是文件,就删除.如果是文件夹.递归给del_file.
            os.remove(file_data)
        else:
            del_file(file_data)


# 清空文件夹,并重新下载文件
def get_tables(stock_infos: list) -> None:
    # del_file("./利润表")
    # del_file("./资产负债表")
    # del_file("./现金流量表")
    for index, count in enumerate(stock_infos):
        stock_num, stock_name = count.split(" ")[0], count.split(" ")[1]
        print("股票代码:{} 股票名称:{} 进度:{}/{}".format(stock_num, stock_name, index + 1, len(stock_infos)))
        url1 = 'http://quotes.money.163.com/service/lrb_' + str(stock_num) + '.html'
        'http://quotes.money.163.com/service/lrb_01830.html'
        while True:
            try:
                print("      利润表下载中,请稍候...")
                content = urllib.request.urlopen(url1, timeout=100).read()
                with open('./利润表/' + stock_name + "_" + stock_num + '利润表.csv', 'wb') as f:
                    f.write(content)
                sleep(0.5)
                break
            except Exception as e:
                print(e)
                if str(e) == 'HTTP Error 404: Not Found':
                    break
                else:
                    continue
        url2 = 'http://quotes.money.163.com/service/zcfzb_' + str(stock_num) + '.html'
        while True:
            try:
                print("      资产负债表下载中,请稍候...")
                content = urllib.request.urlopen(url2, timeout=100).read()
                with open('./资产负债表/' + stock_name + "_" + stock_num + '资产负债表.csv', 'wb') as f:
                    f.write(content)
                sleep(0.5)
                break
            except Exception as e:
                print(e)
                if str(e) == 'HTTP Error 404: Not Found':
                    break
                else:
                    continue
        url3 = 'http://quotes.money.163.com/service/xjllb_' + str(stock_num) + '.html'
        while True:
            try:
                print("      现金流量表下载中,请稍候...")
                content = urllib.request.urlopen(url3, timeout=100).read()
                with open('./现金流量表/' + stock_name + "_" + stock_num + '现金流量表.csv', 'wb') as f:
                    f.write(content)
                sleep(0.5)
                break
            except Exception as e:
                print(e)
                if str(e) == 'HTTP Error 404: Not Found':
                    break
                else:
                    continue


# 用23步分析指定代码的年表
def analyz_table_by_year_in_23_steps(stock_info: str) -> None:
    result = pd.DataFrame()
    stock_num, stock_name = stock_info.split(" ")[0], stock_info.split(" ")[1]
    df_zcfzb = pd.read_csv("./资产负债表/{}_{}资产负债表.csv".format(stock_name, stock_num), encoding='gbk')
    df_lrb = pd.read_csv("./利润表/{}_{}利润表.csv".format(stock_name, stock_num), encoding='gbk')
    df_xjllb = pd.read_csv("./现金流量表/{}_{}现金流量表.csv".format(stock_name, stock_num), encoding='gbk')

    print("股票代码:{} 股票名称:{} 23步分析报表中.".format(stock_num, stock_name))
    # 保留首列的基础上,删除非年报列.
    for column in df_zcfzb.columns:
        if "报告日期" not in column and "12-31" not in column:
            del df_zcfzb[column]
        else:
            result[column] = np.nan
    for column in df_lrb.columns:
        if "报告日期" not in column and "12-31" not in column:
            del df_lrb[column]
    for column in df_xjllb.columns:
        if "报告日期" not in column and "12-31" not in column:
            del df_xjllb[column]

    # 删除掉,result与三大表格只保留最近五年
    while True:
        if len(result.columns) == 6:
            break
        del df_zcfzb[df_zcfzb.columns[6]]
        del df_lrb[df_lrb.columns[6]]
        del df_xjllb[df_xjllb.columns[6]]
        del result[result.columns[6]]
    del result["报告日期"]
    result.insert(0, "款项名称", [])
    df_xjllb.rename(columns={" 报告日期": "报告日期"}, inplace=True)

    # 步骤3.合并资产负债表变动超过±30%,且占资产总计比例超过3%.提出并特殊查看
    KEYWORD_LINES = ["应收票据", "应收账款", "其他应收款", "预付款项", "存货", "在建工程", "长期待摊费",
                     "短期借款", "应付票据", "应付账款", "其他应付款"]
    for index, name in enumerate(df_zcfzb["报告日期"]):
        keep, add_line = False, True
        for keyword in KEYWORD_LINES:
            if keyword in name:
                result = result.append([{"款项名称": "步骤3:资产负债表-{}变动幅度(搜索异常)".format(name.replace("(万元)", ""))}],
                                       ignore_index=True)
                add_line = False
        if add_line:
            result = result.append([{"款项名称": "步骤3:资产负债表-{}变动幅度(搜索异常)".format(name.replace("(万元)", ""))}],
                                   ignore_index=True)
        for index2, name2 in enumerate(result.columns):
            if index2 == 0:
                continue
            elif index2 == len(result.columns) - 1:
                break
            else:
                # 计算变动幅度
                result.iloc[-1, index2] = round(find_accurate_data(df_zcfzb, result.columns[index2], name) / \
                                                (find_accurate_data(df_zcfzb, result.columns[index2 + 1],
                                                                    name) + 0.01) - 1, 4)
                # 计算占据资产比例
                rate_before = find_accurate_data(df_zcfzb, result.columns[index2 + 1], name) / \
                              find_accurate_data(df_zcfzb, result.columns[index2 + 1], "资产总计")

                rate_after = find_accurate_data(df_zcfzb, result.columns[index2], name) / \
                             find_accurate_data(df_zcfzb, result.columns[index2], "资产总计")
                if not -0.3 <= result.iloc[-1, index2] <= 0.3 and max(rate_before, rate_after) >= 0.03:
                    result.iloc[-1, index2] = float_to_percent(
                        result.iloc[-1, index2]) + f' ({float_to_percent(rate_after)})' + "(警示)"
                    keep = True
                else:
                    result.iloc[-1, index2] = float_to_percent(
                        result.iloc[-1, index2]) + f' ({float_to_percent(rate_after)})'
        if not keep:
            result = result[:-1]

    # 步骤4.检查步骤3中的问题
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤4:(搜索步骤3中异常科目,查明原因)"}], ignore_index=True)

    # 步骤5.看总资产,判断公司实力,需人工核对
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤5:总资产(判断公司实力) 人工对照同行业公司"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 导出资产总计值
            result.iloc[-1, index2] = find_accurate_data(df_zcfzb, result.columns[index2], "资产总计")
    #       步骤5.看总资产变动幅度,看扩张能力,以10%为界
    result = result.append([{"款项名称": "步骤5:资产变动幅度(判断公司扩张速度) 10%以上优秀"}], ignore_index=True)
    for index, name in enumerate(result.columns):
        if index == 0:
            continue
        elif index == len(result.columns) - 1:
            break
        else:
            result.iloc[-1, index] = float_to_percent(round(result.iloc[-2, index] / result.iloc[-2, index + 1] - 1, 4))
            if result.iloc[-2, index] / result.iloc[-2, index + 1] - 1 < 0.1:  # 10%以内警示
                result.iloc[-1, index] += "(警示)"
            elif result.iloc[-2, index] / result.iloc[-2, index + 1] - 1 > 0.1:  # 10%以上优秀
                result.iloc[-1, index] += "(优秀)"

    # 步骤6.看资产负债率,以60%为界
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤6:资产负债率 50%以下优秀 70%以上淘汰(判断债务比例)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算资产负债率
            result.iloc[-1, index2] = round(find_accurate_data(df_zcfzb, result.columns[index2], "负债合计") / \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "资产总计"), 4)
            if result.iloc[-1, index2] >= 0.7:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(淘汰)"
            elif result.iloc[-1, index2] < 0.5:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"

    # 步骤7.看有息负债和货币资金,排除偿债风险.有息负债/货币资金>1淘汰.
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤7:有息负债/(货币资金+交易性金融资产) 40%以下优秀 100%以上淘汰(步骤6 40%以上则需要 判断偿债危机)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算有息负债/货币资金
            result.iloc[-1, index2] = round((find_accurate_data(df_zcfzb, result.columns[index2], "短期借款") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "应付利息") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "一年内到期的非流动负债") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "长期借款") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "应付债券")) /
                                            (find_accurate_data(df_zcfzb, result.columns[index2], "货币资金") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "交易性金融资产")), 4)
            if result.iloc[-1, index2] >= 1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(淘汰)"
            elif result.iloc[-1, index2] < 0.4:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"

    # 步骤8.看"应付预收"和"应收预付",判断公司地位
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤8:应付预收-应收预付 0(判断公司地位) 人工对照同行业公司"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算预付预收-应收应付
            result.iloc[-1, index2] = round(find_accurate_data(df_zcfzb, result.columns[index2], "应付票据") + \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "应付账款") + \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "预收账款") - \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "应收票据") - \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "应收账款") - \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "预付款项"), 2)
            if result.iloc[-1, index2] <= 0:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(警示)"
            else:
                result.iloc[-1, index2] = result.iloc[-1, index2]

    result = result.append([{"款项名称": "步骤8:应付预收/应收预付 0.8以下淘汰 1.2以上优秀(判断公司地位)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算应付预收-应收预付
            result.iloc[-1, index2] = round((find_accurate_data(df_zcfzb, result.columns[index2], "应付票据") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "应付账款") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "预收账款")) / \
                                            (find_accurate_data(df_zcfzb, result.columns[index2], "应收票据") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "应收账款") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "预付款项")), 4)
            if result.iloc[-1, index2] <= 0.8:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(淘汰)"
            elif result.iloc[-1, index2] > 1.2:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"

    # 步骤9.看固定资产,判断公司轻重
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤9:(固定资产+在建工程+工程物资)/资产总计 40%以下优秀 40%以上需注意(判断公司轻重)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 (固定资产+在建工程+工程物资)/总资产
            result.iloc[-1, index2] = round((find_accurate_data(df_zcfzb, result.columns[index2], "固定资产") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "在建工程") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "工程物资")) / \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "资产总计"), 4)
            if result.iloc[-1, index2] >= 0.4:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"

    # 步骤10.看投资类资产,判断公司专注程度
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤10:(可供出售金融资产+持有至到期投资+投资性房地产)/资产总计 10%以下优秀(判断公司专注度)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 (固定资产+在建工程+工程物资)/总资产
            result.iloc[-1, index2] = round((find_accurate_data(df_zcfzb, result.columns[index2], "可供出售金融资产") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "持有至到期投资") + \
                                             find_accurate_data(df_zcfzb, result.columns[index2], "投资性房地产")) / \
                                            find_accurate_data(df_zcfzb, result.columns[index2], "资产总计"), 4)
            if result.iloc[-1, index2] >= 0.1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"

    # 步骤11.看利润表与现金流量表 标记异常科目
    result = result.append([{}], ignore_index=True)
    for index, name in enumerate(df_lrb["报告日期"]):
        keep = False
        result = result.append([{"款项名称": "步骤11:利润表-{}变动幅度(搜索异常)".format(name.replace("(万元)", ""))}],
                               ignore_index=True)
        for index2, name2 in enumerate(result.columns):
            if index2 == 0:
                continue
            elif index2 == len(result.columns) - 1:
                break
            else:
                # 计算变动幅度
                result.iloc[-1, index2] = round(find_accurate_data(df_lrb, result.columns[index2], name) / \
                                                (find_accurate_data(df_lrb, result.columns[index2 + 1], name) + 0.01)
                                                - 1, 4)
                # 计算占据营业总收入比例
                rate = max(find_accurate_data(df_lrb, result.columns[index2], name) /
                           find_accurate_data(df_lrb, result.columns[index2], "营业总收入"),
                           find_accurate_data(df_lrb, result.columns[index2 + 1], name) /
                           find_accurate_data(df_lrb, result.columns[index2 + 1], "营业总收入")
                           )
                if not -0.3 <= result.iloc[-1, index2] <= 0.3 and rate >= 0.03:
                    result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
                    keep = True
                else:
                    result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2])
        if not keep:
            result = result[:-1]
    result = result.append([{}], ignore_index=True)
    for index, name in enumerate(df_xjllb["报告日期"]):
        keep = False
        result = result.append([{"款项名称": "步骤11:现金流量表-{}变动幅度(搜索异常)".format(name.replace("(万元)", ""))}],
                               ignore_index=True)
        for index2, name2 in enumerate(result.columns):
            if index2 == 0:
                continue
            elif index2 == len(result.columns) - 1:
                break
            else:
                # 计算变动幅度
                result.iloc[-1, index2] = round(find_accurate_data(df_xjllb, result.columns[index2], name) / \
                                                (find_accurate_data(df_xjllb, result.columns[index2 + 1],
                                                                    name) + 0.01) - 1, 4)
                # 计算占据 三大流量流入 比例
                rate_after = find_accurate_data(df_xjllb, result.columns[index2], name) / ( \
                            find_accurate_data(df_xjllb, result.columns[index2], " 经营活动现金流入小计") + \
                            find_accurate_data(df_xjllb, result.columns[index2], " 投资活动现金流入小计") + \
                            find_accurate_data(df_xjllb, result.columns[index2], " 筹资活动现金流入小计"))
                rate_before = find_accurate_data(df_xjllb, result.columns[index2 + 1], name) / ( \
                            find_accurate_data(df_xjllb, result.columns[index2 + 1], " 经营活动现金流入小计") + \
                            find_accurate_data(df_xjllb, result.columns[index2 + 1], " 投资活动现金流入小计") + \
                            find_accurate_data(df_xjllb, result.columns[index2 + 1], " 筹资活动现金流入小计"))

                if not -0.3 <= result.iloc[-1, index2] <= 0.3 and max(rate_before, rate_after) >= 0.03:
                    result.iloc[-1, index2] = float_to_percent(
                        result.iloc[-1, index2]) + f' ({float_to_percent(rate_after)})(警示)'
                    keep = True
                else:
                    result.iloc[-1, index2] = float_to_percent(
                        result.iloc[-1, index2]) + f' ({float_to_percent(rate_after)})'
        if not keep:
            result = result[:-1]

    # 步骤12.检查步骤11中的问题
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤12:(搜索步骤11中异常科目,查明原因)"}], ignore_index=True)

    # 步骤13.看营业收入,判断公司行业地位及成长能力
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤13:销售商品、提供劳务收到的现金/营业收入 110%以上优秀 100%以下注意(判断公司行业地位,产品竞争力)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 销售商品、提供劳务收到的现金/营业收入
            result.iloc[-1, index2] = round(find_accurate_data(df_xjllb, result.columns[index2], " 销售商品、提供劳务收到的现金") / \
                                            find_accurate_data(df_lrb, result.columns[index2], "营业收入"), 4)
            if result.iloc[-1, index2] < 1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
            elif result.iloc[-1, index2] > 1.1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2])
    result = result.append([{"款项名称": "步骤13:营业收入增长率 10%以上优秀(判断公司成长性)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        elif index2 == len(result.columns) - 1:
            break
        else:
            # 计算 营业收入增长率
            result.iloc[-1, index2] = round(find_accurate_data(df_lrb, result.columns[index2], "营业收入") / \
                                            (find_accurate_data(df_lrb, result.columns[index2 + 1], "营业收入") + 0.01) - 1,
                                            4)
            if result.iloc[-1, index2] < 0.1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
            elif result.iloc[-1, index2] > 0.1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2])

    # 步骤14.看毛利率,判断公司产品竞争力
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤14:毛利率=(营业收入-营业成本)/营业收入 40%以上优秀(判断公司行业地位,产品竞争力)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 毛利率=(营业收入-营业成本)/营业收入
            result.iloc[-1, index2] = round((find_accurate_data(df_lrb, result.columns[index2], "营业收入") - \
                                             find_accurate_data(df_lrb, result.columns[index2], "营业成本")) / \
                                            find_accurate_data(df_lrb, result.columns[index2], "营业收入"), 4)
            if result.iloc[-1, index2] < 0.4:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
            elif result.iloc[-1, index2] > 0.4:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2])

    # 步骤15.看费用率,判断公司成本管控能力
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤15.费用率=(销售费用+管理费用)/营业收入 40%以下优秀 && 60%以上淘汰(判断公司成本管控能力)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 费用率=(销售费用+管理费用)/营业收入
            result.iloc[-1, index2] = round((find_accurate_data(df_lrb, result.columns[index2], "销售费用") + \
                                             find_accurate_data(df_lrb, result.columns[index2], "管理费用")) / \
                                            find_accurate_data(df_lrb, result.columns[index2], "营业收入"), 4)
            if result.iloc[-1, index2] > 0.6:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(淘汰)"
            elif result.iloc[-1, index2] > 0.4:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"

    # 步骤16.看主营利润,判断公司盈利能力/利润质量
    #       先判断主营利润<0
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤16.主营利润=营业收入-营业成本-营业税金及附加-销售费用-管理费用-财务费用 0以下淘汰(判断公司盈利能力) 人工对照同行业公司"}],
                           ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 主营利润=营业收入-营业成本-营业税金及附加-销售费用-管理费用-财务费用
            result.iloc[-1, index2] = round(find_accurate_data(df_lrb, result.columns[index2], "营业收入") - \
                                            find_accurate_data(df_lrb, result.columns[index2], "营业成本") - \
                                            find_accurate_data(df_lrb, result.columns[index2], "营业税金及附加") - \
                                            find_accurate_data(df_lrb, result.columns[index2], "销售费用") - \
                                            find_accurate_data(df_lrb, result.columns[index2], "管理费用") - \
                                            find_accurate_data(df_lrb, result.columns[index2], "财务费用"), 4)
            if result.iloc[-1, index2] < 0:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(淘汰)"
            else:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(优秀)"
    #       判断主营利润率是否小于15%
    result = result.append([{"款项名称": "步骤16.主营利润率=主营利润/营业收入 15%以下淘汰(判断公司盈利能力)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 主营利润/营业收入
            result.iloc[-1, index2] = round(
                float(result.iloc[-2, index2].replace('(淘汰)', '').replace('(优秀)', '')) /
                find_accurate_data(df_lrb, result.columns[index2], "营业收入"), 4)
            if result.iloc[-1, index2] < 0.15:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(淘汰)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
    #       判断主营利润/利润总额是否小于80%
    result = result.append([{"款项名称": "步骤16.主营利润/利润总额 80%以下淘汰 100%以上优秀(判断公司盈利能力)"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 主营利润/利润总额
            result.iloc[-1, index2] = round(
                float(result.iloc[-3, index2].replace('(淘汰)', '').replace('(优秀)', '')) /
                find_accurate_data(df_lrb, result.columns[index2], "利润总额"), 4)
            if result.iloc[-1, index2] < 0.8:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(淘汰)"
            elif result.iloc[-1, index2] > 1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"

    # 步骤17.看净利润,判断公司的经营成果及含金量
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤17.净利润 0以下淘汰(判断公司的经营成果及含金量) 人工对照同行业公司"}],
                           ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 主营利润=营业收入-营业成本-营业税金及附加-销售费用-管理费用-财务费用
            result.iloc[-1, index2] = find_accurate_data(df_lrb, result.columns[index2], "净利润")
            if result.iloc[-1, index2] < 0:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(淘汰)"
            else:
                result.iloc[-1, index2] = result.iloc[-1, index2]
    result = result.append([{"款项名称": "步骤17.最近五年净利润现金比率=经营活动产生的现金流量净额/净利润(判断公司的经营成果及含金量) 100%以下淘汰 120%以上优秀"}],
                           ignore_index=True)
    a, b = 0, 0
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            a += find_accurate_data(df_xjllb, result.columns[index2], " 经营活动产生的现金流量净额")
            b += find_accurate_data(df_xjllb, result.columns[index2], " 净利润")
    # 计算 五年综合净利润现金比率
    result.iloc[-1, 1] = a / b
    if result.iloc[-1, 1] < 1:
        result.iloc[-1, 1] = float_to_percent(result.iloc[-1, 1]) + "(淘汰)"
    elif result.iloc[-1, 1] > 1.2:
        result.iloc[-1, 1] = float_to_percent(result.iloc[-1, 1]) + "(优秀)"
    else:
        result.iloc[-1, 1] = float_to_percent(result.iloc[-1, 1]) + "(警示)"

    # 步骤18.看归母净利润.判断公司自有资本的获利能力
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤18.ROE(净资产收益率)=归母净利润/归母股东权益 20%以上优秀 15%以下淘汰(注:这不是加权值,如因此被淘汰,可查询正规结果)"}],
                           ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 归母净利润/归母股东权益
            result.iloc[-1, index2] = round( \
                find_accurate_data(df_lrb, result.columns[index2], "归属于母公司所有者的净利润") /
                find_accurate_data(df_zcfzb, result.columns[index2], "归属于母公司股东权益合计"), 4)
            if result.iloc[-1, index2] < 0.15:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(淘汰)"
            elif result.iloc[-1, index2] > 0.2:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"

    # 步骤19.看经营活动产生的现金流量净额,判断公司造血能力
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤19.经营活动产生的现金流量净额-固定资产折旧-无形资产摊销-借款利息-现金股利 持续0以下淘汰(判断公司造血能力)"}],
                           ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 经营活动产生的现金流量净额-固定资产折旧-无形资产摊销-借款利息-现金股利
            result.iloc[-1, index2] = find_accurate_data(df_xjllb, result.columns[index2], " 经营活动产生的现金流量净额") - \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 固定资产折旧、油气资产折耗、生产性物资折旧") - \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 无形资产摊销") - \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 偿还债务支付的现金") - \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 分配股利、利润或偿付利息所支付的现金")
            if result.iloc[-1, index2] < 0:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(淘汰)"
            else:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(优秀)"

    result = result.append([{"款项名称": "步骤19.经营活动产生的现金流量净额-固定资产折旧-无形资产摊销-借款利息 持续0以下淘汰(判断公司造血能力)"}],
                           ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 经营活动产生的现金流量净额-固定资产折旧-无形资产摊销-借款利息-现金股利
            result.iloc[-1, index2] = find_accurate_data(df_xjllb, result.columns[index2], " 经营活动产生的现金流量净额") - \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 固定资产折旧、油气资产折耗、生产性物资折旧") - \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 无形资产摊销") - \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 偿还债务支付的现金")
            if result.iloc[-1, index2] < 0:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(淘汰)"
            else:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(优秀)"

    # 步骤20.看购买固定资产、无形资产和其他长期资产支付的现金,判断公司未来成长能力
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤20.购建固定资产、无形资产和其他长期资产支付的现金 (判断公司未来成长能力) 和其他公司做对比"}],
                           ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 购建固定资产、无形资产和其他长期资产支付的现金
            result.iloc[-1, index2] = find_accurate_data(df_xjllb, result.columns[index2], " 购建固定资产、无形资产和其他长期资产所支付的现金")
            if result.iloc[-1, index2] < 0:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(警示)"
            else:
                result.iloc[-1, index2] = str(result.iloc[-1, index2])
    result = result.append(
        [{"款项名称": "步骤20.购建固定资产、无形资产和其他长期资产支付的现金/经营活动产生的现金流量净额 (判断公司未来成长能力) 10%-60%优秀 连续两年>100%或<10%淘汰"}],
        ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 购建固定资产、无形资产和其他长期资产支付的现金
            result.iloc[-1, index2] = find_accurate_data(df_xjllb, result.columns[index2],
                                                         " 购建固定资产、无形资产和其他长期资产所支付的现金") / \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 经营活动产生的现金流量净额")
            if result.iloc[-1, index2] < 0.1 or result.iloc[-1, index2] > 1:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(警示)"
            else:
                result.iloc[-1, index2] = float_to_percent(result.iloc[-1, index2]) + "(优秀)"

    # 步骤21.看"分配给普通股股东及限制性股票持有者股利支付的现金",判断公司慷慨程度
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤21.分配给普通股股东及限制性股票持有者股利支付的现金 分红率低于30%淘汰 人工校验"}],
                           ignore_index=True)

    # 步骤22.看"三大活动现金流量净额的组合类型"
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤22.三大活动现金流量净额的组合类型 连续两年不为'正负负'/'正正负'淘汰."}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            # 计算 "三大活动现金流量净额的组合类型"
            result.iloc[-1, index2] = "正" if find_accurate_data(df_xjllb, result.columns[index2],
                                                                " 经营活动产生的现金流量净额") > 0 else "负"
            result.iloc[-1, index2] += "正" if find_accurate_data(df_xjllb, result.columns[index2],
                                                                 " 投资活动产生的现金流量净额") > 0 else "负"
            result.iloc[-1, index2] += "正" if find_accurate_data(df_xjllb, result.columns[index2],
                                                                 " 筹资活动产生的现金流量净额") > 0 else "负"

            if result.iloc[-1, index2] not in ["正负负", "正正负"]:
                result.iloc[-1, index2] = result.iloc[-1, index2] + "(警示)"
            else:
                result.iloc[-1, index2] = result.iloc[-1, index2] + "(优秀)"

    # 步骤23.看"现金及现金等价物的净增加额",判断公司稳定性
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "步骤23.现金及现金等价物的净增加额+现金分红<0淘汰(判断公司稳定性)."}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            result.iloc[-1, index2] = find_accurate_data(df_xjllb, result.columns[index2], " 现金及现金等价物的净增加额") + \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 分配股利、利润或偿付利息所支付的现金")

            if result.iloc[-1, index2] < 0:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(警示)"
            else:
                result.iloc[-1, index2] = str(result.iloc[-1, index2]) + "(优秀)"

    # 估值法1. 贷款法估计其最低市值.即按照7%贷款,货币资金优先偿还,利润需完美覆盖利息;
    result = result.append([{}], ignore_index=True)
    result = result.append([{}], ignore_index=True)
    result = result.append([{}], ignore_index=True)
    result = result.append([{}], ignore_index=True)
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "估值法1.贷款收购法(按7%贷款利率),计入交易性金融资产,要求净利润可覆盖贷款利息"}], ignore_index=True)

    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            result.iloc[-1, index2] = round(find_accurate_data(df_lrb, result.columns[index2], "净利润") / 0.07 - (
                    find_accurate_data(df_zcfzb, result.columns[index2], "短期借款") +
                    find_accurate_data(df_zcfzb, result.columns[index2], "应付利息") +
                    find_accurate_data(df_zcfzb, result.columns[index2], "一年内到期的非流动负债") +
                    find_accurate_data(df_zcfzb, result.columns[index2], "长期借款") +
                    find_accurate_data(df_zcfzb, result.columns[index2], "应付债券") -
                    find_accurate_data(df_zcfzb, result.columns[index2], "货币资金") -
                    find_accurate_data(df_zcfzb, result.columns[index2], "交易性金融资产")), 2)

    # 估值法2. 股息率估计其最低市值.即按照4%的存款利率计算,由当年股息反算其对应估值
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "估值法2.股利贴现法(按3.5%国债利率),要求股息可覆盖国债机会成本"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0:
            continue
        else:
            result.iloc[-1, index2] = round(
                find_accurate_data(df_xjllb, result.columns[index2], " 分配股利、利润或偿付利息所支付的现金") / 0.035, 2)

    # 估值法3. DCF估值法. 按照 股东权益 + 现金分红 过去五年增长率,反推未来十年增长情况
    #   自由现金流量 = 净利润 + 折旧 + 摊销 - 资本支出 - 资本运营增加 https://xueqiu.com/3207768732/135340312
    result = result.append([{}], ignore_index=True)
    result = result.append([{"款项名称": "估值法3-1.DCF估值法 自由现金流绝对值"}], ignore_index=True)
    for index2, name2 in enumerate(result.columns):
        if index2 == 0 or index2 == 5:
            continue
        else:
            result.iloc[-1, index2] = find_accurate_data(df_xjllb, result.columns[index2], " 净利润") + \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 固定资产折旧、油气资产折耗、生产性物资折旧") + \
                                      find_accurate_data(df_xjllb, result.columns[index2], " 无形资产摊销") - \
                                      find_accurate_data(df_xjllb, result.columns[index2],
                                                         " 购建固定资产、无形资产和其他长期资产所支付的现金") - \
                                      find_accurate_data(df_zcfzb, result.columns[index2], "流动资产合计") + \
                                      find_accurate_data(df_zcfzb, result.columns[index2], "流动负债合计") + \
                                      find_accurate_data(df_zcfzb, result.columns[index2 + 1], "流动资产合计") - \
                                      find_accurate_data(df_zcfzb, result.columns[index2 + 1], "流动负债合计")

    if result.iloc[-1][4] * result.iloc[-1][1] >= 0:  # 增长率为负值无法开根则跳过
        result = result.append([{"款项名称": "估值法3-2.DCF估值法 现金流增长效率"}], ignore_index=True)
        result.iloc[-1, 1] = (result.iloc[-2, 1] / result.iloc[-2, 4]) ** (1 / 3) - 1

        # result.iloc[-1, 1] = result.iloc[-2, 1:4].mean() / result.iloc[-2, 2:5].mean() - 1
        result = result.append([{"款项名称": "估值法3.DCF估值法 十年内现金流估值"}], ignore_index=True)
        sum, now, rate = 0, result.iloc[-3, 1], result.iloc[-2, 1]
        for i in range(10):
            now_rate = rate / 11 * (10 - i)
            now *= (1 + now_rate)
            sum += now
            # print(f'本年增长率:{float_to_percent(now_rate)} \t 本年收益:{now}')

        result.iloc[-1, 1], result.iloc[-1, 4], result.iloc[-1, 5] = round(sum, 2), '增长率', float_to_percent(
            result.iloc[-2, 1])
        result = result.drop(result.index[len(result) - 2])  # 删除增长率
    else:
        result = result.append([{"款项名称": "估值法3.DCF估值法 十年内现金流估值"}], ignore_index=True)
        result.iloc[-1, 1], result.iloc[-1, 4], result.iloc[-1, 5] = '无法计算', '增长率', '无法计算'
    # result = result.drop(result.index[len(result) - 2])  # 删除自由现金流数据

    # result.to_csv(".\\23式报告\\{}_{}_23式报告.csv".format(stock_name, stock_num), encoding='gbk', index=False)
    write_dataframe_to_sheet(pd.DataFrame(result), f"{stock_name}", ".\\23式报告\\汇总表.xls")


# 把dataframe文件写入对应的xlsx中的sheet
def write_dataframe_to_sheet(df: pd.DataFrame, sheet_name: str, output_file_name: str) -> None:
    sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)  # 新建一个sheet
    sheet.col(0).width, sheet.col(1).width, sheet.col(2).width, sheet.col(3).width, sheet.col(4).width, sheet.col(
        5).width = 30000, 5500, 5500, 5500, 5500, 5500  # 设置列宽
    # 设计三个颜色, 黄 红 白 绿
    font, alignment = xlwt.Font(), xlwt.Alignment()
    font.name, alignment.horz, alignment.vert = 'Consolas', 0x02, 0x01

    pattern1, pattern2, pattern3, pattern4 = xlwt.Pattern(), xlwt.Pattern(), xlwt.Pattern(), xlwt.Pattern()
    style1, style2, style3, style4 = xlwt.XFStyle(), xlwt.XFStyle(), xlwt.XFStyle(), xlwt.XFStyle()

    pattern1.pattern, pattern1.pattern_fore_colour = xlwt.Pattern.SOLID_PATTERN, 5
    style1.pattern, style1.font, style1.alignment = pattern1, font, alignment
    pattern2.pattern, pattern2.pattern_fore_colour = xlwt.Pattern.SOLID_PATTERN, 2
    style2.pattern, style2.font, style2.alignment = pattern2, font, alignment
    pattern3.pattern, pattern3.pattern_fore_colour = xlwt.Pattern.SOLID_PATTERN, 1
    style3.pattern, style3.font, style3.alignment = pattern3, font, alignment
    pattern4.pattern, pattern4.pattern_fore_colour = xlwt.Pattern.SOLID_PATTERN, 17
    style4.pattern, style4.font, style4.alignment = pattern4, font, alignment

    STYLE = [style1, style2, style3, style4]

    # 写标题
    for i in range(len(df.columns)):
        sheet.write(0, i, df.columns[i], style3)
    # 写内容
    for i in range(len(df)):  # 行
        for j in range(len(df.iloc[i])):  # 列
            if str(df.iloc[i, j]) == 'nan':  # 是否为nan,如是跳过
                continue
            elif '(警示)' in str(df.iloc[i, j]):  # 警示,染黄色
                sheet.write(i + 1, j, str(df.iloc[i, j]).replace('(警示)', ''), STYLE[0])
            elif '(淘汰)' in str(df.iloc[i, j]):  # 淘汰,染红色
                sheet.write(i + 1, j, str(df.iloc[i, j]).replace('(淘汰)', ''), STYLE[1])
            elif '(优秀)' in str(df.iloc[i, j]):  # 优秀,染绿色
                sheet.write(i + 1, j, str(df.iloc[i, j]).replace('(优秀)', ''), STYLE[3])
            else:  # 正常,染白色
                sheet.write(i + 1, j, df.iloc[i, j], STYLE[2])
    workbook.save(output_file_name)


def find_accurate_data(dataframe: pd.DataFrame, date: str, payment_name: str) -> float:
    index, column = -1, -1

    for index, name in enumerate(dataframe["报告日期"]):
        if payment_name in name and len(name) - len(payment_name) <= 4:
            break
    for column, name in enumerate(dataframe.columns):
        if date in name:
            break

    if payment_name not in dataframe["报告日期"][index] or date not in dataframe.columns[column]:
        print("date:{}  payment_name:{} 的参数使用失误,查无此数据.请核对后再输入.".format(date, payment_name))

    result = dataframe.iloc[index, column]
    if "--" in result:
        return 0
    else:
        return float(result)


# 一个Series元素 每个float元素除以其之后的float元素 最后一位填充np.nan
def Series_devide_self(Series: pd.Series, compare, threshold) -> tuple:
    result, show = [], False
    for i in range(len(Series)):
        if i == 0:
            continue
        if i == len(Series) - 1:
            break
        num = float(Series.iloc[i].replace("--", '0')) / (float(Series.iloc[i + 1].replace("--", '0')) + 10 ** -4)
        # 防止两个0引发的比例变动
        if compare == ">" and num > threshold and \
                float(Series.iloc[i].replace("--", '0')) != 0 or float(Series.iloc[i + 1].replace("--", '0')) != 0:
            show = True
        elif compare == "<" and num < threshold and \
                float(Series.iloc[i].replace("--", '0')) != 0 or float(Series.iloc[i + 1].replace("--", '0')) != 0:
            show = True
        result.append(num)
    return pd.Series(result), show


# 在指定的dataframe表格中,插入一个指定名字的Series为列
def dataframe_add_row(dataframe: pd.DataFrame, row_name: str, series: pd.Series) -> pd.DataFrame:
    result = dataframe
    result = result.reindex(index=list(result.index) + [row_name])
    # 按顺序插入Series数值内容.由于有报告日期,跳过
    for i in range(len(series)):
        result.iloc[-1, i] = series.iloc[i]
    return result


def float_to_percent(num: float) -> str:
    return "%.2f%%" % (num * 100)


if __name__ == '__main__':
    # 1.下载报表环节
    input_file_name = ".\股票代码测试.txt"
    stock_nums = get_isbn_from_txt(input_file_name)
    get_tables(stock_nums)
    # 2.按顺序用23步分析法分析指定序列号的报表
    for stock_info in stock_nums:
        analyz_table_by_year_in_23_steps(stock_info)
