import getpass,imaplib
import email
import os
import time
import pandas as pd
import datetime
import chardet
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr,formataddr
from email.header import decode_header
import datetime
import random
from email import utils
from MyModules.mail_analysis_tq import Format
from MyModules.logger import mylogger
from MyModules.thread_pool import BoundedThreadPoolExecutor
from MyModules.user_agent_list import USER_AGENT_LIST
from MyModules.db_MySQL import MySQLClient
from mail_config import *
import warnings
warnings.filterwarnings('ignore')
imaplib._MAXLINE = 20000000

class Email:
    """本类是从jz@dingshi.net邮箱抓取基金净值（属于投前跟踪标的的净值）"""
    def __init__(self, email_address, password, IMAP4_server="imap.exmail.qq.com"):
        CURRENT_TIME = time.strftime('%Y-%m-%d_%H', time.localtime(time.time()))
        log_filename = os.path.splitext(os.path.basename(__file__))[0] + f'_{CURRENT_TIME}_mail.log'
        LOGS_FILEPATH = './logs'
        os.makedirs(LOGS_FILEPATH, exist_ok=True)
        self.logger = mylogger(os.path.join(LOGS_FILEPATH, log_filename))
        # try:
        #     self.db = MySQLClient(MYSQL_HOST, MYSQL_PORT, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DB, self.logger)# touyan库（即净值库）
        # except:
        #     print('数据库链接失败！')

        self.server = imaplib.IMAP4_SSL(IMAP4_server)
        self.server.login(email_address, password)
        self.server.select("INBOX")# 选择邮箱默认文件夹
        # 搜索匹配的邮件，第一个参数是字符集，None默认就是ASCII编码，第二个参数是查询条件，这里的ALL就是查找全部
        type, data = self.server.search(None, "ALL")
        # 邮件列表,使用空格分割得到邮件索引
        self.msgList = data[0].split()
        # 实例化配置信息0
        self.mc = Format()
        self.MAIL_TIME = 'mail_time_tq.txt'
        # 读取截止采集时间
        if self.read_mail_data(self.MAIL_TIME) is not None and self.read_mail_data(self.MAIL_TIME) != '':
            self.end_time = int(self.read_mail_data(self.MAIL_TIME))
        else:
            self.end_time = self.date_to_strtotime(self.mc.END_TIME)
        
        self.end_time = 1577808000# 测试代码
        print('采集截止时间：')
        print(self.timeStamp_to_date(self.end_time))

    # 邮件编码转换
    def decode_str(self,s):
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value

    def trans_format(self,time_str):
        GMT_FORMAT = "%a, %d %b %Y %H:%M:%S %z"
        return datetime.strptime(time_str, GMT_FORMAT)


    def date_to_strtotime_bak(self, date):
        """把日期格式转化成时间戳"""
        if isinstance(date, datetime):
            date = date.strftime('%Y-%m-%d %H:%M:%S')

        strtotime = time.strptime(date, '%Y-%m-%d %H:%M:%S')
        return int(time.mktime(strtotime))

    def date_to_strtotime(self, date, index=0):
        """把日期格式转化成时间戳"""
        format = [
            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d',
            '%Y年%m月%d日', '%Y年%m月%d', '%Y%m%d %H:%M:%S', '%Y%m%d',
            '%Y/%m/%d %H:%M:%S', '%Y.%m.%d %H:%M:%S',
            '%Y年%m月%d日 %H:%M:%S', '%Y年%m月%d %H:%M:%S'
        ]

        # 如果 date 是 datetime 对象，先转成字符串
        if isinstance(date, datetime):
            date = date.strftime('%Y-%m-%d %H:%M:%S')

        if index >= len(format):
            return None

        try:
            strtotime = time.strptime(date, format[index])
            return int(time.mktime(strtotime))
        except:
            return self.date_to_strtotime(date, index + 1)
        

    def timeStamp_to_date(self,timestamp):
        """把时间戳转换成日期格式"""
        timeArray = time.localtime(timestamp)
        otherStyleTime = time.strftime('%Y-%m-%d %H:%M:%S',timeArray)
        return otherStyleTime

    # 检测邮件编码
    def guess_charset(self,msg):
        # 先从msg对象获取编码
        charset = msg.get_charset()
        if charset is None:
            # 如果获取不到，再从Content-Type字段获取
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                # 去掉尾部不代表编码的字段
                charset = content_type[pos + 8:].strip('; format=flowed; delsp=yes')
        return charset

    # 使用全局变量来保存邮件内容
    mail_content = '\n'

    def detect_encoding(self,file_path):
        """利用chardet猜测文件编码"""
        with open(file_path,'rb') as f:
            result = chardet.detect(f.read())
            return result['encoding']


    # indent用于缩进显示:
    def print_info(self,msg, indent=0):
        """循环遍历读取邮件"""
        global mail_content
        if indent == 0:
            for header in ['From', 'To', 'Subject']:
                value = msg.get(header, '')
                if value:
                    if header == 'Subject':
                        value = self.decode_str(value)
                    else:
                        hdr, addr = parseaddr(value)
                        name = self.decode_str(hdr)
                        value = u'%s <%s>' % (name, addr)
                mail_content += '%s%s: %s' % ('  ' * indent, header, value) + '\n'

        parts = msg.get_payload()

        for n, part in enumerate(parts):
            content_type = part.get_content_type()
            if content_type == 'text/plain':
                content = part.get_payload(decode=True)
                charset = self.guess_charset(msg)
                # charset = 'utf-8'
                if charset:
                    content = content.decode(charset)
                mail_content += '%sText:\n %s' % (' ' * indent, content)
            else:
                # 这里没有读取非text/plain类型的内容，只是读取了其格式，一般为text/html
                mail_content += '%sAttachment: %s' % ('  ' * indent, content_type)
        return mail_content


    def get_mail_file(self, message):
        """提取邮件中的Excel附件内容，登记非Excel附件。"""
        result = []
        non_excel_attachments = []
        has_excel = False

        subject = self.decode_str(message.get('Subject', ''))

        mail_time = ''
        try:
            received = message['Received'].split(';')[-1]
            sent_ts = email.utils.parsedate(received)
            mail_time = time.strftime('%Y/%m/%d %H:%M', time.localtime(time.mktime(sent_ts)))
        except Exception:
            pass

        parts_info = []
        for part in message.walk():
            filename = part.get_filename()
            if filename:
                filename_decoded = self.decode_str(filename)
                is_excel = filename_decoded.lower().endswith(('.xls', '.xlsx'))
                parts_info.append((part, filename_decoded, is_excel))
                if is_excel:
                    has_excel = True

        if has_excel:
            for part, filename_decoded, is_excel in parts_info:
                if is_excel:
                    try:
                        from io import BytesIO
                        sheets = pd.read_excel(BytesIO(part.get_payload(decode=True)), sheet_name=None, header=None)

                        for sheet_name, raw_data in sheets.items():
                            raw_data = raw_data.applymap(
                                lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if isinstance(x, datetime.datetime)
                                else (str(x).strip() if pd.notna(x) else '')
                            )

                            parsed_data = self.mc.index(raw_data.values)
                            
                            # 处理缺陷2：只保留每个 goods_name 最新日期的记录
                            latest_data = {}
                            for item in parsed_data:
                                if not isinstance(item, dict):
                                    continue
                                name = item.get('goods_name')
                                date = item.get('net_time')
                                if name and date:
                                    if name not in latest_data or date > latest_data[name]['net_time']:
                                        latest_data[name] = item
                            
                            if latest_data:
                                result.append(list(latest_data.values()))
                                print(f"成功解析Excel附件 {filename_decoded} - Sheet: {sheet_name}: {list(latest_data.values())}")
                            else:
                                print(f"解析后无有效数据：{filename_decoded} - Sheet: {sheet_name}")

                    except Exception as e:
                        print(f"读取Excel失败: {e}")
            return result

        # 没有Excel，登记非excel附件
        for part, filename_decoded, is_excel in parts_info:
            if not is_excel:
                non_excel_attachments.append({
                    '标的/邮件标题': subject,
                    '附件类型': os.path.splitext(filename_decoded)[-1][1:],
                    '邮件接收时间': mail_time
                })

        if non_excel_attachments:
            output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '非excel附件登记表.xlsx')
            df_new = pd.DataFrame(non_excel_attachments)
            try:
                if os.path.exists(output_path):
                    df_existing = pd.read_excel(output_path)
                    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                else:
                    df_combined = df_new
                df_combined.to_excel(output_path, index=False)
                print(f"非excel附件已登记输出到: {output_path}")
            except Exception as e:
                print(f"写入Excel失败: {e}")

        return result

    

    def write_mail_data(self, date_time, filename):
        """写入截止采集日期"""
        try:
            if date_time:
                with open(filename, 'w') as f:
                    f.write(str(date_time))
        except Exception:
            print('写入文件数据失败！')

    def read_mail_data(self, filename):
        """读取截止采集日期"""
        try:
            with open(filename, 'r') as f:
                return f.read()
        except Exception as e:
            print('读取文件数据失败，错误如下：')
            print(e)
            return None

    def sort_data(self, lists):
        """对净值列表升序排序"""
        temp = [v for val in lists for v in (val if isinstance(val, list) else [val])]
        for i in range(len(temp)):
            for j in range(i + 1, len(temp)):
                try:
                    if temp[i]['net_time'] > temp[j]['net_time']:
                        temp[i], temp[j] = temp[j], temp[i]
                except Exception as e:
                    print('排序函数中,日期比较报错,错误信息如下:')
                    print(e)
        return temp

    def get_mail(self):
        """读取邮件，从第5375封开始，向前（更早）检索时间早过它的（即5375, 5374, 5373...）"""
        try:
            is_write_mail_time = False
            # 用于收集所有有效数据的列表
            all_valid_data = []
            
            for i in range(len(self.msgList), 0, -1):
                latest = self.msgList[i - 1]  # UIDS
                print(i)  # 打印邮件索引
                type, datas = self.server.fetch(latest, '(RFC822)')

                if datas is None or datas == 'nan':
                    continue

                try:
                    raw = email.message_from_bytes(datas[0][1])
                except Exception as es:
                    raw = None
                    print(es)

                if raw is None:
                    continue

                try:
                    subject = email.header.decode_header(raw.get('subject'))
                    mail_coding = subject[0][1]  # 获得当前邮件编码
                    text = datas[0][1].decode(mail_coding, 'ignore')
                except Exception as es:
                    print('编码出错', es)
                message = email.message_from_string(text)
                title = self.decode_str(message.get('Subject', ''))
                print('当前邮件是：' + title)

                try:
                    received = message['Received'].split(';')[-1]
                    sent_ts = utils.parsedate(received)
                    mail_time = int(time.mktime(sent_ts))
                except Exception as te:
                    print(te, '读不到邮件时间，跳过~')
                    self.log_failed_email(title, 0, f"Failed to read mail time: {te}")
                    continue

                print('当前邮件接收时间是：' + self.timeStamp_to_date(mail_time))

                if mail_time <= self.end_time:
                    print('到达采集终点，停止采集！')
                    break

                if not is_write_mail_time:
                    self.write_mail_data(mail_time, self.MAIL_TIME)
                    is_write_mail_time = True

                result = self.get_mail_file(message)
                if not result or (len(result) == 1 and result[0] is None):
                    has_attachment = any(
                        part.get_content_disposition() == 'attachment' or part.get_filename()
                        for part in message.walk()
                    )
                    if not has_attachment:
                        self.log_failed_email(title, mail_time, "邮件没有任何附件")
                        non_excel_attachments = [{
                            '标的/邮件标题': title,
                            '附件类型': '未检测到附件',
                            '邮件接收时间': self.timeStamp_to_date(mail_time)
                        }]
                        output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '非excel附件登记表.xlsx')
                        df_new = pd.DataFrame(non_excel_attachments)
                        try:
                            if os.path.exists(output_path):
                                df_existing = pd.read_excel(output_path)
                                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                            else:
                                df_combined = df_new
                            df_combined.to_excel(output_path, index=False)
                            print(f"未检测到附件，已登记输出到: {output_path}")
                        except Exception as e:
                            print(f"写入Excel失败: {e}")
                    else:
                        has_excel = any(
                            self.decode_str(part.get_filename()).lower().endswith(('.xls', '.xlsx'))
                            for part in message.walk() if part.get_filename()
                        )
                        if not has_excel:
                            self.log_failed_email(title, mail_time, "邮件有附件但没有Excel文件")
                        else:
                            self.log_failed_email(title, mail_time, "邮件包含Excel文件，但内容无法解析或格式不正确")
                else:
                    print('当前result变量值是:', result)
                    # 清洗 result 数据，提取有效记录
                    valid_data = []
                    for sublist in result:
                        if isinstance(sublist, list):
                            for item in sublist:
                                if (isinstance(item, dict) and 
                                    item.get('goods_code') and 
                                    item.get('goods_name') and 
                                    item.get('net_time') and 
                                    item.get('dw_net') and 
                                    item.get('lj_net')):
                                    valid_data.append({
                                        'net_time': self.timeStamp_to_date(item['net_time']),
                                        'goods_code': item['goods_code'],
                                        'goods_name': item['goods_name'],
                                        'dw_net': item['dw_net'],
                                        'lj_net': item['lj_net']
                                    })
                    if valid_data:
                        all_valid_data.extend(valid_data)
                        print(f"提取到有效数据: {valid_data}")

                        # 追加写入Excel，每封邮件处理一次就写一次
                        output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '邮件附件净值抓取.xlsx')
                        df_new = pd.DataFrame(valid_data, columns=['net_time', 'goods_code', 'goods_name', 'dw_net', 'lj_net'])
                        try:
                            if os.path.exists(output_path):
                                # 读取已存在的数据，追加后去重（如需去重可加）
                                df_existing = pd.read_excel(output_path)
                                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                                # 可选：去重
                                df_combined.drop_duplicates(subset=['net_time', 'goods_code'], inplace=True)
                            else:
                                df_combined = df_new
                            df_combined.to_excel(output_path, index=False)
                            print(f"有效数据已追加写入Excel: {output_path}")
                        except Exception as e:
                            print(f"写入Excel失败: {e}")

                time.sleep(1)
                print('=' * 70)
        except Exception as e:
            print(f"get_mail方法发生异常: {e}")


    def log_failed_email(self, subject, mail_time, reason):
        """Log failed emails to an Excel file."""
        output_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(output_dir, 'failed_emails.xlsx')
        
        failure_data = {
            'Email Subject': subject,
            'Receiving Time': self.timeStamp_to_date(mail_time) if mail_time else 'Unknown',
            'Failure Reason': reason
        }
        
        df_new = pd.DataFrame([failure_data])
        
        try:
            if os.path.exists(output_path):
                df_existing = pd.read_excel(output_path)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            else:
                df_combined = df_new
            df_combined.to_excel(output_path, index=False)
            self.logger.info(f"Logged failed email '{subject}' to {output_path}")
        except Exception as e:
            self.logger.error(f"Failed to write to failed_emails.xlsx: {e}")


    def update_data(self,lists):
        """数据进库，插入ty_cj_net_tq表"""
        if lists is None or len(lists) <= 0:
            return False
        print('待入库净值数据:',lists)
        for val in lists:
            if val is None or len(val) <= 0:
                continue

            for v in val:
                where_sql = ' WHERE goods_code="'+v['goods_code']+'" and net_time='+str(v['net_time'])
                check_result = self.db.query_one(TQ_CJ_NET_TABLE,'*',where_sql)
                if check_result is None:
                    self.db.insert_one(TQ_CJ_NET_TABLE,v)
                else:
                    self.db.update(TQ_CJ_NET_TABLE, v,where_sql)

if __name__ == '__main__':
    while True:
        try:
            # 注：Email对象实例化如果放在while外面只会初始化一次邮箱登录、数据库连接、采集截止时间读取，为保证句柄资源可用最好放在循环体内！
            email_obj = Email("jz@dingshi.net", "uJLygjCC9sDhxg4q")
            email_obj.get_mail()
            print('本次采集结束，等待10分钟后重新采集！')
            print('本次采集时间是：')
            print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            time.sleep(300)
        except Exception as e:
            print(e)
