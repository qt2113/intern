import time
from collections import Counter
#import re
import numpy as np
import pandas as pd

class Format:
    """解析附件并提取数据"""
    def __init__(self):
        # 截止采集时间
        self.END_TIME = '2024-07-05 1:41:00'
        self.GOODS_CODE = ''

        # 关键字
        self.MB_TPL = {
            'goods_code': '基金代码,产品代码,资产代码,协会备案编号,备案编码,TA代码',
            'goods_name': '基金名称,产品名称,资产名称,基金全称,账套名称',
            'net_time': '净值日期,日期,业务日期,估值日期,估值基准日,基金净值日期,期末日期',
            'dw_net': '单位净值,计提前单位净值,资产份额净值(元),实际净值,A级单位净值,虚拟净值提取前单位净值,份额净值,期末单位净值,基金份额净值,试算前单位净值',
            'lj_net': '累计单位净值,资产份额累计净值(元),实际累计净值,资产份额累计净值(元),基金份额累计净值,累计净值,A级单位净值,期末累计净值,试算前累计净值,试算前累计单位净值',
            'our_name':'客户名称,投资者名称,客户姓名,TA账号名称',
            'fene': '发生份额,份额余额,持有份额,持仓份额,客户资产份额,参与计提份额,投资者资产份额,计提份额',
            'xn_net':'虚拟单位净值,计提后单位净值,虚拟净值,虚拟后净值,试算单位净值（扣除业绩报酬后）,虚拟净值提取后单位净值,产品虚拟单位净值,试算后单位净值'
        }

        # 特殊标的白名单，该列表内的标的其净值邮件缺失标的代码
        self.SPE_GOODS = [['尚郡衡盈阿尔法','SQU095']]

    def date_to_strtotime2(self,date):
        """把日期格式转化成时间戳"""
        format = ['%Y%m%d','%Y-%m-%d','%Y/%m/%d','%Y.%m.%d','%Y年%m月%d日','%Y年%m月%d', '%Y%m%d %H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y.%m.%d %H:%M:%S','%Y年%m月%d日 %H:%M:%S', '%Y年%m月%d %H:%M:%S']
        for val in format:
            try:
                strtotime = time.strptime(date, val)
                strtotime = int(time.mktime(strtotime))
                return strtotime
            except:
                continue

    def select_data(self,list1):
        """从邮件原始数据中提取净值数据,返回一个列表"""
        """
        待完善：
        1、邮件提醒
        2、日志
        """

        # 缺少标的代码的标的，通过往numpy数组插值的方式补全数据（代码已实现相关功能，但后来标的代码又同邮件一起发了过来，所以暂时屏蔽当前代码）
        # IS_OPEN = False
        # temp = []
        # for i, x in enumerate(list1):
        #     x_str = str(x)# x等于每一行
        #     ser1 = re.search('产品名称',x_str)
        #     ser2 = re.search(self.SPE_GOODS[0][0],x_str)
        #
        #     if ser1 is not None:
        #         x = np.insert(x,1,'基金代码')
        #         temp.append(x)
        #
        #     if ser2 is not None:
        #         IS_OPEN = True
        #         x = np.insert(x,1,self.SPE_GOODS[0][1])
        #         temp.append(x)
        #
        # if IS_OPEN:
        #     list1 = np.array(temp)

        # 表头坐标容器{'goods_name':[行,列],'goods_code':[行,列]...}
        #print(pd.DataFrame(list1))
        coordinate = {}
        count=[]
        # 查找表头坐标，取的是最后一个附件最后一个满足的excel部分
        for tpl in self.MB_TPL:
            coordinate[tpl]=[]
            for row, val in enumerate(list1):
                for col, v in enumerate(val):
                    tpl_list1 = self.MB_TPL[tpl].split(',')
                    v = ' ' if v is np.nan else v
                    v = str(v) 
                    if v.strip(':：') in tpl_list1:
                        coordinate[tpl].append( [row, col])
                        count.append([row,col])
        # 统计第一位数字出现次数
        first_digits = [item[0] for item in count]
        first_counter = Counter(first_digits)
        most_common_first = first_counter.most_common(1)
        
        # 统计第二位数字出现次数
        second_digits = [item[1] for item in count]
        second_counter = Counter(second_digits)
        most_common_second = second_counter.most_common(1)
        #print(most_common_first[0][1],most_common_second[0][1])

        #max_hang=max(coordinate)
        #首先必须要获取标的代码和标的名称，文多文泰晓日三号A是例外，没有goods_name名称
        if ('goods_code' in coordinate.keys()) is False or ('goods_name' in coordinate.keys()) is False:
            return None
        #判断是横向表格还是纵向表格
        try:
            to_rows = len(list1)# 总行数
            to_cols = len(list1[0])# 总列数
            result = []

            # 区分横向和竖向排列,从不同方向查找数据
            if int(most_common_first[0][1]) > int(most_common_second[0][1]):
                # 横向排列，从上往下查找
                for key, value in coordinate.items():
                    if coordinate[key]==[]:
                        continue
                    else:
                        coordinate[key] = [item for item in value if item[0] == most_common_first[0][0]]
                        coordinate[key] =coordinate[key][0]
                range_data = to_rows# 迭代行
                add_map = 0
            else:
                # 竖向排列，从左往右查找数据
                for key, value in coordinate.items():
                    if coordinate[key]==[]:
                        continue
                    else:
                        coordinate[key] = [item for item in value if item[1] == most_common_second[0][0]]
                        coordinate[key] =coordinate[key][0]
                range_data = to_cols# 迭代列
                add_map = 1
            #print(coordinate)
            for i in range(range_data):# 迭代行/列,后面需要修改
                #print(i)
                if i > coordinate['goods_code'][add_map]:#会取最后一个表头后面的数据
                    temp = {}
                    for t in coordinate:
                        if coordinate[t]==[]:
                            continue
                        coordinate[t][add_map] = i
                        data = list1[coordinate[t][0]][coordinate[t][1]]
                        #print(data)
                        if data and data is not np.nan:
                            if t == 'goods_code' or t == 'goods_name'or t == 'our_name':
                                if isinstance(data,str)==0:
                                    continue
                                temp[t] = data.strip()
                            elif t == 'net_time':
                                temp[t] = self.date_to_strtotime2(str(data).strip())
                            elif t == 'dw_net' or t == 'lj_net' or  t == 'fene'or  t == 'xn_net':
                                if isinstance(data,int) or isinstance(data,float):
                                    temp[t] = data
                                else:
                                    temp[t] = data.strip()
                        #print(data)

                    result.append(temp)
            #print(result)
            if result:
                # 补净值，某些邮件只包含单位净值或累计净值，需要把缺少的净值补全
                for index,val in enumerate(result):
                    # 没有单位净值
                    if ('dw_net' in val.keys()) is False and 'lj_net' in val.keys():
                        result[index]['dw_net'] = result[index]['lj_net']

                    # 没有累计净值
                    if ('lj_net' in val.keys()) is False and 'dw_net' in val.keys():
                        result[index]['lj_net'] = result[index]['dw_net']

                # 特殊邮件处理，如招证净值邮件抓不到净值日期，需要指定坐标抓取
                #print(list1[0][0])
                if list1[0][0] == '资产净值公告':
                    net_time = self.date_to_strtotime2(list1[2][0])
                    result[0]['net_time'] = net_time

            #print('result: ',result)
            return result

        except Exception as e:
            print('mail_analysis错误如下:',e)
            return None

    def format_by_mail1(self,list1):
        """邮件解析专用代码1，标的：文多文泰晓日三号私募证券投资基金-A类专用"""
        result = []
        temp = dict()
        for index,val in enumerate(list1):
            if index == 4:
                temp['goods_name'] = val[0]
            elif index == 5:
                temp['goods_code'] = val[4]
                temp['net_time'] = self.date_to_strtotime2(val[10])
                temp['dw_net'] = str(val[8])
                temp['lj_net'] = str(val[9])
                temp['xn_net'] = str(val[7])
                temp['fene'] = str(val[6])
                temp['our_name'] = str(val[3])
        result.append(temp)
        #print('专用格式抓取到的数据：',result)
        return result

    def index(self,rawdata):
        """数据重构与校验，只有满足要求的才会选出来"""
        # result = self.select_data(rawdata)
        to_str = str(rawdata)
        if to_str.find('文多文泰') > -1:#
            result = self.format_by_mail1(rawdata)
        else:
            result = self.select_data(rawdata)
        if isinstance(result,list) is False:
            return None

        data = []
        data_full = True
        
        for val in result:# 只有满足所有值都获取了之后才会采集这条数据
            temp = {}
            if ('goods_code' in val.keys()) is False or val['goods_code'] is np.nan:
                data_full = False
                continue
            else:
                temp['goods_code'] = val['goods_code']

            if ('goods_name' in val.keys()) is False or val['goods_name'] is np.nan:
                data_full = False
                continue
            else:
                temp['goods_name'] = val['goods_name']

            if ('net_time' in val.keys()) is False or val['net_time'] is np.nan:
                data_full = False
                continue
            else:
                temp['net_time'] = val['net_time']

            if ('dw_net' in val.keys()) is False or val['dw_net'] is np.nan:
                data_full = False
                continue
            else:
                temp['dw_net'] = val['dw_net']

            if ('lj_net' in val.keys()) is False or val['lj_net'] is np.nan:
                data_full = False
                continue
            else:
                temp['lj_net'] = val['lj_net']

            if ('xn_net' in val.keys()) is False or val['xn_net'] is np.nan:
                data_full = False
                continue
            else:
                temp['xn_net'] = val['xn_net']
                
            if ('our_name' in val.keys()) is False or val['our_name'] is np.nan:
                data_full = False
                continue
            else:
                temp['fof_name'] = val['our_name']
                
            if ('fene' in val.keys()) is False or val['fene'] is np.nan:
                data_full = False
                continue
            else:
                fene = str(val['fene'])
                temp['fene'] = fene.replace(',','')

            if data_full:
                temp['fq_net'] = 0
                temp['add_time'] = int(time.time())
                temp['source'] = 3
                data.append(temp)
            
        # print(data)
        return data

