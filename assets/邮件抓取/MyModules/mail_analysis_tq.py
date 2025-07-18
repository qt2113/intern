import time
import re
import numpy as np

class Format:
    """解析附件并提取数据"""
    def __init__(self):
        # 截止采集时间
        self.END_TIME = '2022-03-29 00:00:00'
        self.GOODS_CODE = ''

        # 关键字
        self.MB_TPL = {
            'goods_code': '基金代码,产品代码,资产代码,协会备案编号,备案编码',
            'goods_name': '基金名称,产品名称,资产名称,基金全称,账套名称',
            'net_time': '净值日期,日期,业务日期,估值日期,估值基准日,基金净值日期,期末日期',
            'dw_net': '单位净值,计提前单位净值,资产份额净值(元),实际净值,A级单位净值,虚拟净值提取前单位净值,份额净值,期末单位净值,基金份额净值,试算前单位净值',
            'lj_net': '累计单位净值,资产份额累计净值(元),实际累计净值,资产份额累计净值(元),基金份额累计净值,累计净值,A级单位净值,期末累计净值,试算前累计净值,试算前累计单位净值'
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

    def select_data(self,list):
        """从邮件原始数据中提取净值数据"""
        """
        待完善：
        1、邮件提醒
        2、日志
        """

        # 缺少标的代码的标的，通过往numpy数组插值的方式补全数据（代码已实现相关功能，但后来标的代码又同邮件一起发了过来，所以暂时屏蔽当前代码）
        # IS_OPEN = False
        # temp = []
        # for i, x in enumerate(list):
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
        #     list = np.array(temp)

        # 表头坐标容器{'goods_name':[行,列],'goods_code':[行,列]...}
        coordinate = {}

        # 查找表头坐标
        for tpl in self.MB_TPL:
            for row, val in enumerate(list):
                for col, v in enumerate(val):
                    tpl_list = self.MB_TPL[tpl].split(',')
                    v = ' ' if v is np.nan else v
                    v = str(v) if isinstance(v, int) or isinstance(v, float) else v
                    if v.strip(':：') in tpl_list:
                        coordinate[tpl] = [row, col]

        if ('goods_code' in coordinate.keys()) is False or ('goods_name' in coordinate.keys()) is False:
            return None

        try:
            to_rows = len(list)# 总行数
            to_cols = len(list[0])# 总列数
            result = []

            # 区分横向和竖向排列,从不同方向查找数据
            if coordinate['goods_code'][0] == coordinate['goods_name'][0]:
                # 横向排列，从上往下查找
                range_data = to_rows# 迭代行
                add_map = 0
            else:
                # 竖向排列，从左往右查找数据
                range_data = to_cols# 迭代列
                add_map = 1

            for i in range(range_data):  # 迭代行/列
                if i > coordinate['goods_code'][add_map]:
                    temp = {}
                    for t in coordinate:
                        if coordinate[t]==[]:
                            continue
                        coordinate[t][add_map] = i
                        data = list[coordinate[t][0]][coordinate[t][1]]

                        if data and data is not np.nan:
                            if t == 'goods_code' or t == 'goods_name':
                                temp[t] = data.strip()
                            elif t == 'net_time':
                                temp[t] = self.date_to_strtotime2(data.strip())
                            elif t == 'dw_net' or t == 'lj_net':
                                if isinstance(data,int) or isinstance(data,float):
                                    temp[t] = data
                                else:
                                    temp[t] = data.strip()

                    result.append(temp)

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
                if list[0][0] == '资产净值公告':
                    net_time = self.date_to_strtotime2(list[2][0])
                    result[0]['net_time'] = net_time

            print('result: ',result)
            return result

        except Exception as e:
            print('mail_analysis错误如下:',e)
            return None

    def format_by_mail1(self,list):
        """邮件解析专用代码1，标的：文多文泰晓日三号私募证券投资基金-A类专用"""
        result = []
        temp = dict()
        for index,val in enumerate(list):
            if index == 4:
                temp['goods_name'] = val[0]
            elif index == 5:
                temp['goods_code'] = val[4]
                temp['net_time'] = self.date_to_strtotime2(val[10])
                temp['dw_net'] = str(val[8])
                temp['lj_net'] = str(val[9])
            result.append(temp)
        print('专用格式抓取到的数据：',result)
        return result

    def format_by_mail2(self,list):
        """邮件解析专用代码2，标的：汉盛新千禧壹号私募证券投资基金 专用"""
        result = []
        temp = dict()
        for index,val in enumerate(list):
            if index == 8:
                temp['net_time'] = self.date_to_strtotime2(val[0])
            elif index == 6:
                temp['goods_name'] = val[0]
                temp['goods_code'] = val[1]
                temp['dw_net'] = val[2]
                temp['lj_net'] = val[3]
            result.append(temp)
        return result

    def filtra_code(self,goods_code):
        """去掉代码中的符号或中文"""
        pattern = re.compile(r'[A-za-z0-9]+')
        matches = pattern.findall(goods_code)
        if len(matches) > 0:
            str = ''.join(matches)
            str = str.replace('_', '')
            return str

    def index(self,rawdata):
        """数据校验"""
        # result = self.select_data(rawdata)
        to_str = str(rawdata)
        if to_str.find('文多文泰晓日三号') > -1:
            result = self.format_by_mail1(rawdata)
        elif to_str.find('海通证券') > -1:
            result = self.format_by_mail2(rawdata)
        else:
            result = self.select_data(rawdata)

        if isinstance(result,list) is False:
            return None

        data = []
        data_full = True

        for val in result:
            temp = {}
            if ('goods_code' in val.keys()) is False or val['goods_code'] is np.nan:
                data_full = False
                continue
            else:
                temp['goods_code'] = self.filtra_code(val['goods_code'])

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
                val['dw_net'] = str(val['dw_net'])
                temp['dw_net'] = eval(val['dw_net'])

            if ('lj_net' in val.keys()) is False or val['lj_net'] is np.nan:
                data_full = False
                continue
            else:
                val['lj_net'] = str(val['lj_net'])
                temp['lj_net'] = eval(val['lj_net'])

            if data_full:
                temp['fq_net'] = 0
                temp['add_time'] = int(time.time())
                temp['source'] = 3
                data.append(temp)

        return data
