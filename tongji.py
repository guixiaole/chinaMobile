from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
from pandas import DataFrame

#
# order_dao_chu = pd.read_excel('C://Users//gxl//Desktop//中国移动//2021428//订单清查_导出.xlsx')
# wu_li_qingcha = pd.read_excel('C://Users//gxl//Desktop//中国移动//2021428//物理清查_20210428091935.xlsx',
#                               converters={'站址编码': str}).values
# chanpinguagao = pd.read_excel('C://Users//gxl//Desktop//中国移动//塔类产品服务费结算详单.xlsx')
# tie_ta_ji_zhun = pd.read_excel('C://Users//gxl//Desktop//中国移动//铁塔计费价格.xlsx', sheet_name=[0, 1, 2, 3])


tie_ta_ji_zhun = pd.read_excel('C://Users//gxl//Desktop//中国移动//铁塔计费价格.xlsx', sheet_name=[0, 1, 2, 3])
# order_dao_chu = pd.read_excel('C://Users//gxl//Desktop//中国移动//订单清查_导出表.xlsx', converters={'站址编码': str}).values
order_dao_chu = pd.read_excel('C://Users//gxl//Desktop//中国移动//202164//订单清查截止6月3日.xlsx')

wu_li_qingcha = pd.read_excel('C://Users//gxl//Desktop//中国移动//202164//物理清查截止6月3日.xlsx',
                              converters={'站址编码': str}).values
chanpinguagao = pd.read_excel('C://Users//gxl//Desktop//中国移动//202164//塔类产品服务费结算详单-移动 .xlsx')




def get_jifang(temp_int):
    jifang_map = {}
    res_jifang = {}
    temp_jifang = tie_ta_ji_zhun[temp_int].values
    for i in range(len(temp_jifang)):
        temp_money = Decimal(str(temp_jifang[i][len(temp_jifang[0]) - 1])).quantize(Decimal('0.00'),
                                                                                    rounding=ROUND_HALF_UP)
        jifang_map[temp_jifang[i][2]] = float(temp_money)
    return jifang_map


def getorder_guagao():
    """
    查找物理账单，然后进行匹配
    :return:
    """

    res_tieta_jizhun = {}
    # order_qingcha_len = len(order_qingcha)
    wu_li_qingcha_len = len(wu_li_qingcha)
    order_dao_chu_len = len(order_dao_chu.values)
    flag = 0
    for i in range(wu_li_qingcha_len):
        if wu_li_qingcha[i][3] == wu_li_qingcha[i][3] and wu_li_qingcha[i][21] == wu_li_qingcha[i][21]:
            temp = wu_li_qingcha[i][21].split('、')[0]

            res_tieta_jizhun[str(wu_li_qingcha[i][3])] = temp

    return res_tieta_jizhun


def caljifang_buyizhi():
    qu_map = {
        '芙蓉区': 0,
        '开福区': 1,
        '天心区': 2,
        '雨花区': 3,
        '岳麓区': 4,
        '长沙县': 5,
        '望城区': 6,
        '浏阳市': 7,
        '宁乡县': 8
    }
    jifang = [0 for _ in range(9)]
    order = order_dao_chu.values
    order_len = len(order)
    for i in range(30, order_len):
        if order[i][13] == order[i][13]:
            temp = qu_map[order[i][3]]
            order_index = str(order[i][10]).split('、')[0]
            order_index = order_index.split('（')[0]
            # 现场中的在机房和配套中的索引
            temp1 = str(order[i][17]).split('、')[0]
            temp1 = temp1.split('（')[0]
            if order_index != temp1:
                jifang[temp] += 1
                print(order_index)
                print(temp1)
    print(jifang)


def cal_calshare_jifang():
    qu_map = {
        '芙蓉区': 0,
        '开福区': 1,
        '天心区': 2,
        '雨花区': 3,
        '岳麓区': 4,
        '长沙县': 5,
        '望城区': 6,
        '浏阳市': 7,
        '宁乡县': 8
    }
    share_jifang = [0 for _ in range(9)]
    order = order_dao_chu.values
    order_len = len(order)
    for i in range(30, order_len):
        if order[i][13] == order[i][13] and order[i][11] == order[i][11] and order[i][19] == order[i][19]:
            temp = qu_map[order[i][3]]
            xianchang_share_jifang = len(order[i][19].split('+'))
            order_share_jifang = int(order[i][11])
            if xianchang_share_jifang != order_share_jifang:
                share_jifang[temp] += 1
                print(order[i][19], end='')
                print("=", end='')
                print(xianchang_share_jifang)
                print(order_share_jifang)
                print("==============")
    print(share_jifang)


def getchanpinqingdan():
    res_chanpin_guagao = {}
    chanpin = chanpinguagao.values
    chanpin_len = len(chanpin)
    for i in range(chanpin_len):
        if chanpin[i][7] == chanpin[i][7] and chanpin[i][27] == chanpin[i][27] and chanpin[i][11] == chanpin[i][11]:
            if chanpin[i][27] != '-':
                temp = str(chanpin[i][27])

                res_chanpin_guagao[str(chanpin[i][7])] = temp

    return res_chanpin_guagao


def caltongji():
    qu_map = {
        '芙蓉区': 0,
        '开福区': 1,
        '天心区': 2,
        '雨花区': 3,
        '岳麓区': 4,
        '长沙县': 5,
        '望城区': 6,
        '浏阳市': 7,
        '宁乡县': 8
    }
    taxing = {
        '普通地面塔': 6,
        '景观塔': 5,
        '简易塔': 4,
        '普通楼面塔': 3,
        '楼面抱杆': 2,
        '无铁塔': 1
    }
    # 清理订单数，            0
    # 塔型不一致，            1
    # 塔型比铁塔造价低，       2
    # 机房类型不一致          3
    # 塔型比铁塔造价低        4
    # 挂高不一致             5
    # 比铁塔挂高低           6
    # 塔类共享不一致           7
    # 比铁塔造价低             8
    # 机房共享不一致           9
    # 比铁塔造价低            10
    # 配套共享不一致           11
    # 比铁塔造价低             12
    tongji = [[0 for _ in range(13)] for __ in range(9)]

    order = order_dao_chu.values
    order_len = len(order)
    chanpinguagao = getchanpinqingdan()
    order_guagao = getorder_guagao()
    jifang = get_jifang(1)

    for i in range(30, order_len):
        if order[i][13] == order[i][13] and order[i][8] == order[i][8] and order[i][3] == order[i][3] \
                and order[i][16] == order[i][16] and order[i][2] == order[i][2] and order[i][17] == order[i][17] \
                and order[i][18] == order[i][18] and order[i][9] == order[i][9] \
                and order[i][11] == order[i][11] and order[i][19] == order[i][19] and order[i][8] != '美化树' and order[i][
            16] != '美化树':
            # 统计的为清理订单数
            temp = qu_map[order[i][3]]
            tongji[temp][0] += 1
            #  第8列为订单的挂高
            #  第16列为现场的挂高
            if order[i][8] != order[i][16] and order[i][8] != '美化树' and order[i][16] != '美化树':
                tongji[temp][1] += 1
                order8 = order[i][8]
                if order[i][8] == '抱杆':
                    order8 = '楼面抱杆'
                order16 = order[i][16]
                if order[i][16] == '抱杆':
                    order16 = '楼面抱杆'
                if taxing[order8] < taxing[order16]:
                    tongji[temp][2] += 1
            else:
                # 假设塔型一样的话查找挂高。
                if order[i][8] != '普通楼面塔' and order[i][8] != '楼面抱杆' and order[i][8] != '无铁塔' and order[i][8] != '美化树':
                    if order[i][2] in chanpinguagao.keys() and order[i][2] in order_guagao.keys():
                        chanpin_gg = chanpinguagao[order[i][2]]
                        order_gg = order_guagao[order[i][2]]
                        flag_gg = 0
                        for j in range(len(order_gg)):
                            if order_gg[j] == '塔':
                                flag_gg = j
                                break
                        if flag_gg != 0 and flag_gg != len(order_gg) - 1:
                            temp_gg = order_gg[flag_gg + 1:]
                            if temp_gg != chanpin_gg:
                                tongji[temp][5] += 1
                                print(i)
                                print(chanpin_gg)
                                if len(temp_gg) > 2 and len(chanpin_gg) > 2:
                                    if temp_gg[0] == 'H':
                                        order_gg_int = int(temp_gg[2:])
                                    else:
                                        order_gg_int = int(temp_gg[:2])
                                    if chanpin_gg[0] == 'H':
                                        chanpin_gg_int = int(chanpin_gg[2:4])
                                    else:
                                        if chanpin_gg == '0.0':
                                            chanpin_gg_int = int(float(chanpin_gg))
                                        else:
                                            chanpin_gg_int = int(chanpin_gg[:2])
                                    if order_gg_int < chanpin_gg_int:
                                        tongji[temp][6] += 1
                #  开始统计机房
            # 订单中的在机房和配套中的索引

            order_index = str(order[i][8]) + str(order[i][10])
            # 现场中的在机房和配套中的索引
            temp1 = str(order[i][17]).split('（')[0]
            xianchang_index = str(order[i][16]) + temp1
            if order_index != xianchang_index:
                tongji[temp][3] += 1
                if order_index in jifang.keys() and xianchang_index in jifang.keys():
                    if jifang[order_index] > jifang[xianchang_index]:
                        tongji[temp][4] += 1
            #  计算塔类共享类型
            xianchang_share_ta = len(order[i][18].split('+'))
            order_share_ta = int(order[i][9])
            if xianchang_share_ta != order_share_ta:
                tongji[temp][7] += 1
                if xianchang_share_ta < order_share_ta:
                    tongji[temp][8] += 1
            # 计算机房共享
            xianchang_share_jifang = len(order[i][19].split('+'))
            order_share_jifang = int(order[i][11])
            if xianchang_share_jifang != order_share_jifang:
                tongji[temp][9] += 1
                if xianchang_share_jifang < order_share_jifang:
                    tongji[temp][10] += 1
            # 计算配套共享
            xianchang_share_peitao = len(order[i][20].split('+'))
            order_share_peitao = int(order[i][12])
            if xianchang_share_peitao != order_share_peitao:
                tongji[temp][11] += 1
                if xianchang_share_peitao < order_share_peitao:
                    tongji[temp][12] += 1
    qu_tongji = ['芙蓉区',
                 '开福区',
                 '天心区',
                 '雨花区',
                 '岳麓区',
                 '长沙县',
                 '望城区',
                 '浏阳市',
                 '宁乡县']
    excle_head = ["清理订单数",
                  "塔型不一致",
                  "塔型比铁塔造价低",
                  "机房类型不一致",
                  "机房类型比铁塔造价低",
                  "挂高不一致",
                  "挂高比铁塔挂高低",
                  "塔类共享不一致",
                  "塔类共享比铁塔价低",
                  "机房共享不一致",
                  "机房共享比铁塔价低",
                  "配套共享不一致",
                  "配套共享比铁塔价低"]

    tongji_daochu = pd.read_excel('C://Users//gxl//Desktop//中国移动//2021528//统计导出.xlsx')
    print(tongji_daochu)
    print(tongji)
    for i in range(len(tongji)):
        for j in range(len(tongji[0])):
            #     # tongji[i].insert(0, tongji_daochu[i])

            tongji_daochu[excle_head[j]][i] = tongji[i][j]

    DataFrame(tongji_daochu).to_excel('C://Users//gxl//Desktop//中国移动//202164//统计导出.xlsx', index=False,
                                      header=True)

    print(tongji)


if __name__ == '__main__':
    caltongji()
    # cal_calshare_jifang()
