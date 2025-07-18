import time

class Format:
    """邮件格式解析"""

    def __init__(self):

        # 截止采集时间
        self.END_TIME = '2022-12-28 00:00:00'

        # 各个标的坐标数据
        self.COORDINATE = [
            # 坐标顺序：标的识别代码，标的名、标的代码、净值日期、单位净值、累计净值
            ['SXY662',[2,0],[2,5],[7,0],[2, 9],[2, 10]],# 佳岳-实投8号私募证券投资基金
            ['20172B', [3, 2], [3, 1], [3, 0], [3, 3], [3, 4]],
        ]

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

    def data_analysis(self,rawdata):
        """读取净值数据并做检查"""
        data = {}
        for val in self.COORDINATE:
            try:
                goods_code = rawdata[val[2][0], val[2][1]]
            except Exception as es:
                continue

            if goods_code != val[0]:
                continue

            goods_name  = rawdata[val[1][0],val[1][1]]
            net_time    = rawdata[val[3][0],val[3][1]]
            dw_net      = rawdata[val[4][0],val[4][1]]
            lj_net      = rawdata[val[5][0],val[5][1]]

            goods_name  = goods_name.strip()
            goods_code  = goods_code.strip()
            net_time    = net_time.strip()

            try:
                dw_net = dw_net.strip()
                lj_net = lj_net.strip()
                dw_net = eval(dw_net)
                lj_net = eval(lj_net)
                nav_data_correct = True
            except Exception as e:
                print('错误如下：'+e)
                nav_data_correct = False

            if goods_name != '' and goods_name != 'nan' and goods_name is not None:
                data['goods_name'] = goods_name
            else:
                data['goods_name'] = False

            if goods_code != '' and goods_code != 'nan' and goods_code is not None:
                data['goods_code'] = goods_code
            else:
                data['goods_code'] = False

            if net_time != '' and net_time != 'nan' and net_time is not None:
                data['net_time'] = self.date_to_strtotime2(net_time)
            else:
                data['net_time'] = False

            if nav_data_correct is True:
                if isinstance(dw_net,int) or isinstance(dw_net,float):
                    data['dw_net'] = dw_net
                else:
                    data['dw_net'] = False

                if isinstance(lj_net,int) or isinstance(lj_net,float):
                    data['lj_net'] = lj_net
                else:
                    data['lj_net'] = False
            else:
                data['dw_net'] = False
                data['lj_net'] = False

            # 只要有一个字段的数据不正确都不允许进库
            for k in data:
                if data[k] is False:
                    data = {}
                    break
            data['add_time'] = int(time.time())
            data['fq_net'] = 0
            break

        return data