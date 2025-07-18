# -*- coding: utf-8 -*-
"""
Created on Tue Sep 10 13:01:24 2024
@author: rls
"""
import imaplib
import email
import os
import rule1
from bs4 import BeautifulSoup
import time
import pandas as pd
#import datetime
#import chardet
#from email.utils import parseaddr
from email.header import decode_header
from email import utils
from MyModules.logger import mylogger
from MyModules.db_MySQL import MySQLClient
from mail_config import *


import warnings
import sys
import traceback
import traceback
warnings.filterwarnings('ignore')
# from func_timeout import func_set_timeout
imaplib._MAXLINE  =  10000000
class Email:
    """本类是从邮箱抓取私募基金净值正文"""
    def __init__(self, email_address, password, IMAP4_server="imap.exmail.qq.com"):# IMAP4_server="imap.exmail.qq.com"
        CURRENT_TIME = time.strftime('%Y-%m-%d_%H', time.localtime(time.time()))
        log_filename = os.path.splitext(os.path.basename(__file__))[0] + f'_{CURRENT_TIME}_mail.log'
        LOGS_FILEPATH = './logs'
        os.makedirs(LOGS_FILEPATH, exist_ok=True)
        self.logger = mylogger(os.path.join(LOGS_FILEPATH, log_filename))
        # try:
        #     self.db = MySQLClient(MYSQL_HOST, MYSQL_PORT, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DB,self.logger)  # touyan库（即净值库）
        # except:
        #     print('数据库链接失败！')

        self.server = imaplib.IMAP4_SSL(IMAP4_server)
        self.server.login(email_address, password)
        #print(self.server.list()) 看文件夹的名字
        self.server.select("INBOX")# 选择邮箱测试文件夹    &UXZO1mWHTvZZOQ-/TESTBOX
        # 搜索匹配的邮件，第一个参数是字符集，None默认就是ASCII编码，第二个参数是查询条件，这里的ALL就是查找全部
        type, data = self.server.search(None, 'ALL')
       
        # 邮件列表,使用空格分割得到邮件索引
        self.msgList = data[0].split()
        self.MAIL_TIME = 'mail_time_tq_body.txt'

        # 读取截止采集时间
        if self.read_mail_data(self.MAIL_TIME) is not None and self.read_mail_data(self.MAIL_TIME) != '':
            self.end_time = int(self.read_mail_data(self.MAIL_TIME))
        else:
            self.end_time = self.date_to_strtotime('2020-01-01')

        self.end_time = 1577808000  # 测试代码

        print('采集截止时间：')
        print(self.timeStamp_to_date(self.end_time))

        #print(self.msgList)
        # 实例化配置信息0
        # self.begin_time=self.date_to_strtotime('2024-10-17 00:00:00')
        # self.end_time=self.date_to_strtotime('2024-10-1 00:00:00')
        # print('采集开始时间：')
        # print(self.timeStamp_to_date(self.begin_time))
        # print('采集截止时间：')
        # print(self.timeStamp_to_date(self.end_time))
    # 标准化指标名
        self.mapping = {
            'goods_code': ['基金代码','产品代码','资产代码','协会备案编号','备案编码','TA代码','协会备案编码（Fund Filling Code）'],
            'goods_name': ['基金名称','产品名称','资产名称','基金全称','账套名称','产品名称（Fund Name）'],
            'net_time': ['净值日期','日期','业务日期','估值日期','估值基准日','基金净值日期','期末日期','日期（NAV As Of Date）'],
            'dw_net': ['单位净值','计提前单位净值','资产份额净值(元)','实际净值','A级单位净值','虚拟净值提取前单位净值','份额净值','期末单位净值','基金份额净值','试算前单位净值','单位净值（NAV/Share）'],
            'lj_net': ['最新净值','累计单位净值','资产份额累计净值(元)','实际累计净值','资产份额累计净值(元)','基金份额累计净值','累计净值','A级单位净值','期末累计净值','试算前累计净值','试算前累计单位净值','累计单位净值（Accumulated NAV/Share）']
        }
    # 增加了“最新净值”

    # 邮件编码转换
    def decode_str(self,s):
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value

    def date_to_strtotime(self,date):
        """把日期格式转化成时间戳"""
        """把日期格式转化成时间戳"""
        format = ['%Y%m%d', '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y年%m月%d日', '%Y年%m月%d', '%Y%m%d %H:%M:%S','%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y.%m.%d %H:%M:%S', '%Y年%m月%d日 %H:%M:%S','%Y年%m月%d %H:%M:%S']

        for val in format:
            try:
                strtotime = time.strptime(date, val)
                strtotime = int(time.mktime(strtotime))
                return strtotime
            except:
                continue
        return None

    def timeStamp_to_date(self,timestamp):
        """把时间戳转换成日期格式"""
        timeArray = time.localtime(timestamp)
        otherStyleTime = time.strftime('%Y-%m-%d %H:%M:%S',timeArray)
        return otherStyleTime
       
    def write_mail_data(self,date_time,filename):
        """写入截止采集日期"""
        try:
            if date_time:
                with open(filename,'w') as f:
                    f.write(str(date_time))
                    f.close()
        except:
            print('写入文件数据失败！')

    def read_mail_data(self,filename):
        """读取截止采集日期"""
        try:
            with open(filename,'r') as f:
                times = f.read()
                f.close()
                return times
        except Exception as e:
            print('读取文件数据失败，错误如下：',e)
            return None


    def get_result(self,df1):
        global res
       
        res=pd.concat([res,df1])
        return res
    
    # 解码
    def decodeStr(self,s):
        try:
            subject = email.header.decode_header(s)
        except:
            return None
        sub_bytes = subject[0][0]
        sub_charset = subject[0][1]
        if None == sub_charset:
            subject = sub_bytes
        elif 'unknown-8bit' == sub_charset:
            subject = str(sub_bytes, 'utf8')
        else:
            subject = str(sub_bytes, sub_charset)
        return subject


    # 检测编码
    def guessCharset(self,msg):
        charset = msg.get_charset()
        if charset is None:
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                charset = content_type[pos + 8:].strip()
        return charset
    
    # 解析正文
    def parseBody(self,msg, indent=0):
        # 如果是多重结构，则进入该结构
        if msg.is_multipart():
            parts = msg.get_payload()
            for n, part in enumerate(parts):
                self.parseBody(part, indent + 1)
        else:
            content_type = msg.get_content_type()
            # 表格的格式为html，因此查找表头为‘text/html’的内容
            if content_type == 'text/html':
                content = msg.get_payload(decode=True)
                charset = self.guessCharset(msg)
                if charset:
                    content = content.decode(charset)
                global htmltext
                htmltext = content


    # 对指标名称进行映射
    def map_column(self,col):
        for new_col, old_cols in self.mapping.items():
            if col in old_cols:
                return new_col
        return col

    # 获得标题
    def getHeading(self,msg):
        sub = msg.get('Subject')
        return self.decodeStr(sub)

    def showMaxFactor(self,num):
        count = num / 2
        while count > 1:
            if num % count == 0:
                return count
                break
            count -= 1

    # 提取数据并存储到表格中
    def extractData(self, msg):
        global htmltext
        self.parseBody(msg)
        soup = BeautifulSoup(htmltext, 'html.parser')
        table = soup.find('table')
        title = self.getHeading(msg)
        received = msg.get('Date')
        try:
            sent_ts = utils.parsedate(received)
            mail_time = self.timeStamp_to_date(int(time.mktime(sent_ts)))
        except Exception:
            mail_time = received

        if not table:
            self._append_result(mail_time, title, "未识别到表格")
            return None

        rows = table.find_all('tr')
        if len(rows) < 2:
            print('表格行数不足')
            return None

        header = [cell.get_text(strip=True) for cell in rows[0].find_all(['th', 'td'])]
        data = []
        for row in rows[1:]:
            cells = [cell.get_text(strip=True) for cell in row.find_all(['td', 'th'])]
            if not cells:
                continue
            if len(cells) < len(header):
                cells += [None] * (len(header) - len(cells))
            data.append(cells[:len(header)])

        if not data:
            print('表头或数据行为空')
            return None

        df = pd.DataFrame(data, columns=[self.map_column(col) for col in header])
        required = ['net_time', 'goods_code', 'goods_name', 'dw_net', 'lj_net']
        for col in required:
            if col not in df.columns:
                df[col] = None
        result = df[required]

        if result.empty or all(result.iloc[0][col] in [None, '', 'None'] for col in required):
            self._append_result(mail_time, title, "未能识别到净值相关字段")
            return None

        print(f'解析成功，列名: {result.columns.tolist()}, 数据: \n{result.to_string()}')
        return result

    def _append_result(self, mail_time, title, remark):
        output = {
            "net_time": mail_time,
            "goods_code": None,
            "goods_name": None,
            "dw_net": None,
            "lj_net": None,
            "备注": f"{remark}，邮件标题: {title}, 日期: {mail_time}"
        }
        if not hasattr(self, 'all_results'):
            self.all_results = []
        self.all_results.append(output)
        folder_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        excel_path = os.path.join(folder_path, "邮件正文净值抓取.xlsx")
        pd.DataFrame(self.all_results).to_excel(excel_path, index=False)
        print(f"已实时输出到: {excel_path}")

    def insert_data(self, result):
        if result is None or result.empty:
            return False
        n_time = str(result.loc[0, 'net_time'])
        net_time = self.date_to_strtotime(n_time)
        goods_code = str(result.loc[0, 'goods_code'])
        goods_name = str(result.loc[0, 'goods_name'])
        dw_net = str(result.loc[0, 'dw_net'])
        lj_net = str(result.loc[0, 'lj_net'])
        _insert_data = {}

        if not net_time: return False
        _insert_data['net_time'] = net_time
        if not goods_code: return False
        _insert_data['goods_code'] = goods_code
        if not goods_name: return False
        _insert_data['goods_name'] = goods_name
        if not dw_net: return False
        _insert_data['dw_net'] = eval(dw_net)
        if not lj_net: return False
        _insert_data['lj_net'] = eval(lj_net)

        _insert_data.update({'fq_net': 0, 'add_time': int(time.time()), 'source': 3, 'is_sync': 0})
        is_exists = self.db.query_one(TQ_CJ_NET_TABLE, 'count(*) as count',
                                      'where goods_code="%s" and net_time=%d' % (goods_code, _insert_data['net_time']))
        print('is_exists: ', is_exists)
        if is_exists[0] > 0:
            self.db.update(TQ_CJ_NET_TABLE, _insert_data,
                           'where goods_code="%s" and net_time=%d' % (goods_code, _insert_data['net_time']))
        else:
            self.db.insert_one(TQ_CJ_NET_TABLE, _insert_data)
        print('净值更新完成~')

    def get_mail(self):
        try:
            is_write_mail_time = False
            all_results = getattr(self, 'all_results', [])
            for i in range(len(self.msgList), 0, -1):
                latest = self.msgList[i-1]
                print(i, latest)
                typ, datas = self.server.fetch(latest, '(RFC822)')
                if not datas or datas == 'nan':
                    print("邮件数据获取失败，跳过")
                    continue
                try:
                    raw = email.message_from_bytes(datas[0][1])
                except Exception as es:
                    print('解析邮件原始内容失败:', es)
                    continue
                try:
                    subject = email.header.decode_header(raw.get('subject'))
                    mail_coding = subject[0][1] or 'gbk'
                    text = datas[0][1].decode(mail_coding, 'ignore')
                except Exception as es:
                    print('编码出错', es)
                    continue
                message = email.message_from_string(text)
                title = self.decode_str(message.get('Subject', ''))
                print('当前邮件是：' + str(title))
                try:
                    received = message['Date']
                    sent_ts = utils.parsedate(received)
                    mail_time = int(time.mktime(sent_ts))
                    mail_time_str = self.timeStamp_to_date(mail_time)
                    print('当前邮件接收时间是：' + mail_time_str)
                except Exception as es:
                    print('邮件接收时间出错', es)
                    continue
                if mail_time is not None and mail_time <= self.end_time:
                    print('到达采集终点，停止采集！')
                    break
                if not is_write_mail_time and mail_time is not None:
                    self.write_mail_data(mail_time, self.MAIL_TIME)
                    is_write_mail_time = True
                try:
                    self.parseBody(message)
                    result = self.extractData(message)
                    print('从extractData函数返回的数据：\n')
                    if result is not None and not result.empty:
                        print('\t'.join(result.columns))
                        for _, row in result.iterrows():
                            print('\t'.join(str(x) for x in row))
                        for _, row in result.iterrows():
                            output = {col: str(row.get(col, None)) for col in ["net_time", "goods_code", "goods_name", "dw_net", "lj_net"]}
                            output["备注"] = "抓取成功"
                            all_results.append(output)
                    else:
                        output = {
                            "net_time": mail_time_str,
                            "goods_code": None,
                            "goods_name": None,
                            "dw_net": None,
                            "lj_net": None,
                            "备注": f"extractData返回空或None，邮件标题: {title}, 日期: {mail_time_str}"
                        }
                        all_results.append(output)
                    folder_path = os.path.dirname(os.path.abspath(sys.argv[0]))
                    excel_path = os.path.join(folder_path, "邮件正文净值抓取.xlsx")
                    pd.DataFrame(all_results).to_excel(excel_path, index=False)
                    print(f"已实时输出到: {excel_path}")
                except Exception as e:
                    print('报错了，错误信息如下1：')
                    print(e)
                    traceback.print_exc()
                    continue
                time.sleep(1)
                print('=' * 70)
            self.all_results = all_results
        except Exception as e:
            print('报错了，错误信息如下2：')
            print(e)
            traceback.print_exc()
        self.server.close()
        self.server.logout()


if __name__ == '__main__':
    # res=pd.DataFrame(columns=["goods_code","goods_name","net_time","dw_net","lj_net"])#用来存结果
    while True:
        # 注：Email对象实例化如果放在while外面只会初始化一次邮箱登录、数据库连接、采集截止时间读取，为保证句柄资源可用最好放在循环体内！
        #Email("1832703548@qq.com", "hqmutbdipgghecei")("1349115004@qq.com", "wkmryuorksodihch")("jz@dingshi.net", "FRaYwArtZMhhuCAP")
        email_obj =  Email("jz@dingshi.net", "uJLygjCC9sDhxg4q")
        email_obj.get_mail()


        print('本次采集结束，等待5分钟后重新采集！')
        print('本次采集时间是：',time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
        time.sleep(300)
        #res.to_excel('resultaaa.xlsx')





