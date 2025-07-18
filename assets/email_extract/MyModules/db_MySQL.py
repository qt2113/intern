import pymysql
import traceback
import time
import os
import re
from MyModules.logger import mylogger


class MySQLClient:

    def __init__(self, host, port, user, password, db, logger=None):
        if logger is None:
            CURRENT_TIME = time.strftime('%Y-%m-%d_%H', time.localtime(time.time()))
            log_filename = os.path.splitext(os.path.basename(__file__))[0] + f'_{CURRENT_TIME}.log'
            self.logger = mylogger(log_filename)
        else:
            self.logger = logger

        try:
            self.db = pymysql.connect(host=host, port=port, user=user, password=password, db=db, autocommit=True)
            self.cursor = self.db.cursor()

        except Exception:
            self.logger.error(traceback.format_exc())

    def save_one(self, table, detail: dict):
        """
        插入单条数据
        如果为 MyIsam 引擎，则更新数据时不会影响自增 id，若为 InnoDB 引擎，则更新一条数据时自增 id 会 +1
        :param table:
        :param detail:
        :return:
        """
        keys = ', '.join(detail.keys())
        values = ', '.join(['%s'] * len(detail))

        sql = f"INSERT INTO {table} ({keys}) VALUES ({values}) ON DUPLICATE KEY UPDATE "
        update = ', '.join([f"{key} = VALUES({key})" for key in detail if key != 'update_time'])    # 如果数据已存在，则更新除了 add_time 以外的字段
        sql += update

        try:
            self.cursor.execute(sql, tuple(detail.values()))
            print(detail)

        except Exception:
            self.logger.error(traceback.format_exc())
            self.db.rollback()

    def save_many(self, table, detail_list: list):
        """批量插入数据表"""
        keys = ', '.join(detail_list[0].keys())
        values = ', '.join(['%s'] * len(detail_list[0]))

        datas = [tuple(detail.values()) for detail in detail_list]

        sql = f"INSERT INTO {table} ({keys}) VALUES ({values}) ON DUPLICATE KEY UPDATE "
        update = ', '.join([f"{key} = VALUES({key})" for key in detail_list[0] if key != 'add_time'])
        sql += update
        print(sql)

        # try:
        #     self.cursor.executemany(sql, datas)
        #     print(detail_list)
        # except Exception:
        #     self.logger.error(traceback.format_exc())
        #     self.db.rollback()

    def insert_one(self, table, detail: dict):
        keys = ', '.join(detail.keys())
        values = ', '.join(['%s'] * len(detail))

        sql = f"INSERT INTO {table} ({keys}) VALUES ({values})"
        try:
            self.cursor.execute(sql, tuple(detail.values()))
            print('INSERT', detail)

            last_id = self.cursor.lastrowid
            return last_id

        except Exception as e:
            if e.args[0] == 2006:
                self.db.ping()
            else:
                self.logger.error(traceback.format_exc() + '\n' + 'sql: ' + sql + '\n\n' + 'detail: ' + str(detail))
                self.db.rollback()

    def insert_many(self, table, detail_list: list):
        detail = detail_list[0]
        keys = ', '.join(detail.keys())
        values = ', '.join(['%s'] * len(detail))

        sql = f"INSERT INTO {table} ({keys}) VALUES ({values})"
        try:
            values = [tuple(detail.values()) for detail in detail_list]
            return self.cursor.executemany(sql, values)
            # print(f'INSERT [{len(detail_list)}]', detail_list)

        except Exception as e:
            if e.args[0] == 2006:
                self.db.ping()
            else:
                self.logger.error(traceback.format_exc() + '\n' + 'sql: ' + sql + '\n\n' + 'detail_list: ' + str(detail_list))
                self.db.rollback()

    def update(self, table, detail, condition=''):
        """更新数据表指定字段（detail中的key）"""
        if 'add_time' in detail:
            detail.pop('add_time')

        upsets = ', '.join(["{key} = %s".format(key=key) for key in detail])
        sql = f'UPDATE {table} SET {upsets} {condition}'

        try:
            result = self.cursor.execute(sql, tuple(detail.values()))
            # print('UPDATE', detail)
            return result   # 成功为1，失败为0

        except Exception as e:
            if e.args[0] == 2006:
                self.db.ping()
            else:
                self.logger.error(traceback.format_exc() + '\n' + 'sql: ' + sql + '\n\n' + 'detail: ' + str(detail))
                self.db.rollback()
                return 0

    def update_many(self, table, detail_list: list, condition=''):
        _detail_list = []
        for detail in detail_list:
            if 'add_time' in detail:
                detail.pop('add_time')
            _detail_list.append(detail)

        upsets = ', '.join(["{key} = %s".format(key=key) for key in _detail_list[0]])
        sql = f'UPDATE {table} SET {upsets} {condition}'

        try:
            field = re.search(r'WHERE\s(.*?)\s', condition).group(1).strip()
            values = [tuple(detail.values()) + (detail[field],) for detail in _detail_list]
            result = self.cursor.executemany(sql, values)
            print(f'UPDATE [{len(_detail_list)}]', _detail_list)
            return result   # 成功为1，失败为0

        except Exception as e:
            if e.args[0] == 2006:
                self.db.ping()
            else:
                self.logger.error(traceback.format_exc() + '\n' + 'sql: ' + sql + '\n\n' + 'detail_list: ' + str(detail_list))
                self.db.rollback()
                return 0

    def query_one(self, table, field, condition=''):
        sql = f'SELECT {field} FROM {table} {condition} LIMIT 1'

        try:
            self.cursor.execute(sql)
            result = self.cursor.fetchone()
            return result

        except Exception as e:
            if e.args[0] == 2006:
                self.db.ping()
            else:
                self.logger.error(traceback.format_exc() + '\n' + 'sql: ' + sql)
                return None

    def query_many(self, table, field, condition=''):
        sql = f'SELECT {field} FROM {table} {condition}'

        try:
            self.cursor.execute(sql)
            result = self.cursor.fetchall()
            return result

        except Exception as e:
            if e.args[0] == 2006:
                self.db.ping()
            else:
                self.logger.error(traceback.format_exc() + '\n' + 'sql: ' + sql)
                return None

    def execute(self, sql):
        try:
            return self.cursor.execute(sql)

        except Exception as e:
            if e.args[0] == 2006:
                self.db.ping()
            else:
                self.logger.error(traceback.format_exc() + '\n' + 'sql: ' + sql)
                return None