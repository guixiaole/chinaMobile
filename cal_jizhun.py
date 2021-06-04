from filecmp import cmp
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import openpyxl
import re

from pandas import DataFrame

share_price = [1, 1, 0.7, 0.6, 0.6, 0.6, 0.6]
#
tie_ta_ji_zhun = pd.read_excel('C://Users//gxl//Desktop//中国移动//铁塔计费价格.xlsx', sheet_name=[0, 1, 2, 3])
# order_dao_chu = pd.read_excel('C://Users//gxl//Desktop//中国移动//订单清查_导出表.xlsx', converters={'站址编码': str}).values
order_dao_chu = pd.read_excel('C://Users//gxl//Desktop//中国移动//202164//订单清查截止6月3日.xlsx')

wu_li_qingcha = pd.read_excel('C://Users//gxl//Desktop//中国移动//202164//物理清查截止6月3日.xlsx',
                              converters={'站址编码': str}).values
chanpinguagao = pd.read_excel('C://Users//gxl//Desktop//中国移动//202164//塔类产品服务费结算详单-移动 .xlsx')



def jizhun():
    """
    铁塔挂高价格计算
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
    for i in range(len(temp_tieta) - 3):
        temp_money = Decimal(str(temp_tieta[i][len(temp_tieta[0]) - 1])).quantize(Decimal('0.00'),
                                                                                  rounding=ROUND_HALF_UP)
        if i < len(temp_tieta) - 3:
            tie_ta_map[temp_tieta[i][2]] = float(temp_money)

    tie_ta_map['楼面塔'] = 381.04
    tie_ta_map['普通楼面塔'] = 381.04
    tie_ta_map['楼面抱杆'] = 113.20
    tie_ta_map['无铁塔'] = 0.00
    return tie_ta_map


def get_jifang(temp_int):
    jifang_map = {}
    res_jifang = {}
    temp_jifang = tie_ta_ji_zhun[temp_int].values
    for i in range(len(temp_jifang)):
        temp_money = Decimal(str(temp_jifang[i][len(temp_jifang[0]) - 1])).quantize(Decimal('0.00'),
                                                                                    rounding=ROUND_HALF_UP)
        jifang_map[temp_jifang[i][2]] = float(temp_money)
    return jifang_map
    # order = order_dao_chu
    # for i in range(len(order)):
    #     if order[i][17] == order[i][17] and order[i][16] == order[i][16]:
    #         temp = order[i][17].split('（')[0]
    #         print(temp)
    #         if (str(order[i][16])+temp) in jifang_map.keys():
    #             res_jifang[str(order[i][2])] = jifang_map[str(order[i][16]) + temp]
    # return res_jifang


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


def getorder_guagao():
    """
    查找物理账单，然后进行匹配
    :return:
    """
    tie_ta_map = jizhun()
    res_tieta_jizhun = {}
    # order_qingcha_len = len(order_qingcha)
    wu_li_qingcha_len = len(wu_li_qingcha)
    order_dao_chu_len = len(order_dao_chu.values)
    flag = 0
    for i in range(1,wu_li_qingcha_len):
        if wu_li_qingcha[i][3] == wu_li_qingcha[i][3] and wu_li_qingcha[i][21] == wu_li_qingcha[i][21]:
            temp = wu_li_qingcha[i][21].split('、')[0]
            if temp in tie_ta_map.keys():
                res_tieta_jizhun[str(int(wu_li_qingcha[i][3]))] = tie_ta_map[temp]

    return res_tieta_jizhun


def getchanpinqingdan():
    tie_ta_map = jizhun()
    res_chanpin_guagao = {}
    chanpin = chanpinguagao.values
    chanpin_len = len(chanpin)
    for i in range(chanpin_len):
        if chanpin[i][7] == chanpin[i][7] and chanpin[i][14] == chanpin[i][14] and chanpin[i][11] == chanpin[i][11]:
            print(type(chanpin[i][14]))
            if chanpin[i][14] == '-' or chanpin[i][14] == '0.0':
                temp = chanpin[i][11]
            else:
                temp = str(chanpin[i][11]) + str(chanpin[i][14])
            if temp in tie_ta_map.keys():
                res_chanpin_guagao[str(int(chanpin[i][7]))] = tie_ta_map[temp]
            else:
                res_chanpin_guagao[str(int(chanpin[i][7]))] = 0.0
    return res_chanpin_guagao


if __name__ == '__main__':
    """
    首先计算出现实的价格
    """
    # 机房与配套查找的是这个
    jifang = get_jifang(1)
    peitao = get_jifang(2)
    # guagao = jizhun()
    #  订单的挂高查找的是这个
    # DataFrame(data).to_excel('example.xlsx', sheet_name='Sheet1', index=False, header=True)
    chanpinguagao = getchanpinqingdan()
    #  现场挂高是查找这个
    wuli_guagao = getorder_guagao()
    order = order_dao_chu.values
    order_len = len(order)
    for i in range(1, order_len):
        if i>=34:
            print(i)
        if order[i][2] == order[i][2] and order[i][6] == order[i][6] and order[i][8] == order[i][8] and order[i][10] == \
                order[i][10] and order[i][17] == order[i][17] and order[i][16] == order[i][16]:
            if order[i][9] == order[i][9] and order[i][11] == order[i][11] and order[i][12] == order[i][12] and \
                    order[i][18] == order[i][18] and order[i][19] == order[i][19] and order[i][20] == order[i][20]:
                # 订单中的在机房和配套中的索引

                order_index = str(order[i][8]) + str(order[i][10])
                # 现场中的在机房和配套中的索引
                temp = str(order[i][17]).split('（')[0]
                xianchang_index = str(order[i][16]) + temp
                #  站址编码中的索引

                zhanzhibianma = str(int(order[i][2])).split('.')[0]
                # 按照一个一个的来
                # 移动核算塔类基准价格
                if zhanzhibianma in wuli_guagao.keys() and xianchang_index in jifang.keys() and \
                        zhanzhibianma in chanpinguagao.keys() and order_index in jifang.keys() and \
                        xianchang_index in peitao.keys() and order_index in peitao.keys() and order[i][19] == order[i][
                    19] \
                        and order[i][20] == order[i][20] and order[i][12] == order[i][12] and order[i][11] == order[i][
                    11]:
                    print(i)
                    try:
                        mobile_guagao = wuli_guagao[zhanzhibianma]
                        if i >= 368:
                            print(i)
                            print(order[i][19])
                        order_dao_chu['移动核算塔类基准价格'][i] = mobile_guagao
                        # 铁塔账单塔类基准价格
                        order_guagao = chanpinguagao[zhanzhibianma]
                        order_dao_chu['铁塔账单塔类基准价格'][i] = order_guagao
                        # 铁塔账单机房基准价格

                        order_jifang = jifang[order_index]
                        order_dao_chu['铁塔账单机房基准价格'][i] = order_jifang
                        # 移动核算机房基准价格
                        mobile_jifang = jifang[xianchang_index]
                        order_dao_chu['移动核算机房基准价格'][i] = mobile_jifang
                        # 移动核算配套基准价格
                        mobile_peitao = peitao[xianchang_index]
                        order_dao_chu['移动核算配套基准价格'][i] = mobile_peitao
                        # 铁塔账单配套基准价格
                        order_peitao = peitao[order_index]
                        order_dao_chu['铁塔账单配套基准价格'][i] = order_peitao
                        # 移动按共享核算后塔类价格
                        share_mobile_guagao = mobile_guagao * share_price[len(order[i][18].split('+'))]
                        order_dao_chu['移动按共享核算后塔类价格'][i] = share_mobile_guagao
                        # 铁塔账单塔类共享核算后塔类价格
                        share_order_guagao = order_guagao * share_price[int(order[i][9])]
                        order_dao_chu['铁塔账单塔类共享核算后塔类价格'][i] = share_order_guagao
                        # 共享后机房移动价格
                        share_mobile_jifang = mobile_jifang * share_price[len(order[i][19].split('+'))]
                        order_dao_chu['共享后机房移动价格'][i] = share_mobile_jifang
                        # 铁塔账单共享后基准价
                        share_order_jifang = order_jifang * share_price[int(order[i][11])]
                        order_dao_chu['铁塔账单共享后基准价'][i] = share_order_jifang
                        # 共享后配套移动价格
                        share_mobile_peitao = mobile_peitao * share_price[len(order[i][20].split('+'))]
                        order_dao_chu['共享后配套移动价格'][i] = share_mobile_peitao
                        # 铁塔账单配套共享后基准价
                        share_order_peitao = order_peitao * share_price[int(order[i][12])]
                        order_dao_chu['铁塔账单配套共享后基准价'][i] = share_order_peitao
                        # 总体差异金额
                        mobile_all = (
                                mobile_peitao + mobile_jifang + mobile_guagao + share_mobile_guagao + share_mobile_jifang + share_mobile_peitao)
                        order_all = (
                                order_jifang + order_peitao + order_guagao + share_order_guagao + share_order_jifang + share_order_peitao)
                        all_minus = mobile_all - order_all
                        order_dao_chu['总体差异金额'][i] = all_minus
                    except RuntimeError:
                        print("====")
    print(order_dao_chu)
    DataFrame(order_dao_chu).to_excel('C://Users//gxl//Desktop//中国移动//202164//订单清查_导出.xlsx', index=False, header=True)
