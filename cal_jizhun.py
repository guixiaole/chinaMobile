from filecmp import cmp
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import openpyxl
import re

tie_ta_ji_zhun = pd.read_excel('C://Users//gxl//Desktop//中国移动//铁塔计费价格.xlsx', sheet_name=[0, 1, 2, 3])
# order_dao_chu = pd.read_excel('C://Users//gxl//Desktop//中国移动//订单清查_导出表.xlsx', converters={'站址编码': str}).values
order_dao_chu = pd.read_excel('C://Users//gxl//Desktop//中国移动//2021423//订单清查_20210423085741.xlsx').values
wu_li_qingcha = pd.read_excel('C://Users//gxl//Desktop//中国移动//2021423//物理清查_20210423085813.xlsx',
                              converters={'站址编码': str}).values


def jizhun():
    """
    计算基准价格
    其中基准价格为（铁塔基准价+机房基准价+配套基准价）+维护费*1.1
    其中分别在表一表二
    :return:
    """
    # 转成一个二维数组
    res_tieta_jizhun = {}

    # tie_ta_ji_zhun = pd.read_excel('C://Users//gxl//Desktop//中国移动//铁塔计费价格.xlsx', sheet_name=[0, 1, 2, 3])

    tie_ta_map = {}
    temp_tieta = tie_ta_ji_zhun[0].values
    for i in range(len(temp_tieta)):
        temp_money = Decimal(str(temp_tieta[i][len(temp_tieta[0]) - 1])).quantize(Decimal('0.00'),
                                                                                  rounding=ROUND_HALF_UP)
        print(temp_money)
        if i < len(temp_tieta) - 4:

            tie_ta_map[temp_tieta[i][0][:-1] + str(temp_tieta[i][1])] = float(temp_money)
        else:
            if temp_tieta[i][0] == '楼面抱杆':
                tie_ta_map['楼面'] = float(temp_money)
            else:
                tie_ta_map[temp_tieta[i][0][:-1]] = float(temp_money)

    # order_qingcha_len = len(order_qingcha)
    wu_li_qingcha_len = len(wu_li_qingcha)
    order_dao_chu_len = len(order_dao_chu)
    flag = 0
    for i in range(order_dao_chu_len):
        temp = order_dao_chu[i][2]
        if temp == temp:
            print(temp)
            for j in range(wu_li_qingcha_len):
                if temp == wu_li_qingcha[j][3]:
                    if wu_li_qingcha[j][18] == wu_li_qingcha[j][18]:
                        if wu_li_qingcha[j][18] not in tie_ta_map.keys() and wu_li_qingcha[j][19] == wu_li_qingcha[j][
                            19]:
                            money = tie_ta_map[wu_li_qingcha[j][18] + str(wu_li_qingcha[j][19])]
                            res_tieta_jizhun[str(temp)] = money
                        else:
                            if wu_li_qingcha[j][18] in tie_ta_map.keys():
                                money = tie_ta_map[wu_li_qingcha[j][18]]
                                res_tieta_jizhun[str(temp)] = money
                        break
    # 在物理清单总表中获取价格。
    return res_tieta_jizhun


def get_jifang(temp_int):
    jifang_map = {}
    res_jifang = {}
    temp_jifang = tie_ta_ji_zhun[temp_int].values
    for i in range(len(temp_jifang)):
        temp_money = Decimal(str(temp_jifang[i][len(temp_jifang[0]) - 1])).quantize(Decimal('0.00'),
                                                                                    rounding=ROUND_HALF_UP)
        jifang_map[temp_jifang[i][2]] = float(temp_money)
    order = order_dao_chu
    for i in range(len(order)):
        if order[i][17] == order[i][17] and order[i][16] == order[i][16]:
            temp = order[i][17].split('（')[0]
            print(temp)
            if (str(order[i][16])+temp) in jifang_map.keys():
                res_jifang[str(order[i][2])] = jifang_map[str(order[i][16]) + temp]
    return res_jifang


def get_weihu():
    temp_weihu = tie_ta_ji_zhun[3].values
    quxian = []
    taxing = []
    res = {}
    for i in range(1, len(temp_weihu)):
        pingyuan_money = Decimal(str(temp_weihu[i][len(temp_weihu[0]) - 2])).quantize(Decimal('0.00'),
                                                                                      rounding=ROUND_HALF_UP)
        shanqu_money = Decimal(str(temp_weihu[i][len(temp_weihu[0]) - 1])).quantize(Decimal('0.00'),
                                                                                    rounding=ROUND_HALF_UP)
        if temp_weihu[i][0] == temp_weihu[i][0]:
            # 进行保存
            quxian = temp_weihu[i][0].split('、')
        else:
            if temp_weihu[i][1] == temp_weihu[i][1]:
                taxing = temp_weihu[i][1].split('、')
        for h in range(len(quxian)):
            for j in range(len(taxing)):
                res[str(quxian[h]) + str(taxing[j]) + str(temp_weihu[i][2]) + '平原'] = float(pingyuan_money)
                res[str(quxian[h]) + str(taxing[j]) + str(temp_weihu[i][2]) + '山区'] = float(shanqu_money)
    return res


if __name__ == '__main__':
    """
    首先计算出现实的价格
    """
    # jifang = get_jifang(1)
    peitao = get_jifang(2)
    # print(jifang)
    # print(len(jifang))
    print(peitao)
    print(len(peitao))