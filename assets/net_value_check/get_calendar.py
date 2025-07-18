'''此代码用于从百度接口抓取节假日日历数据，并保存为CSV文件。'''

import requests as req
import datetime
import json
import random
import time
import pandas as pd
import os

# 保存的CSV文件名
CALENDAR_CSV = "calendar_data.csv"

# 从百度接口中获取节假日日历数据
class Calendar:

    def __init__(self):
        self.dataTime = str(random.random())
        self.header = {
            "Content-Type": "application/json;charset=UTF-8"
        }
        self.param = {
            "query": '',
            "resource_id": "39043",
            "t": self.dataTime,
            "ie": "utf8",
            "oe": "gbk",
            "format": "json",
            "tn": "wisetpl",
            "cb": ""
        }

        if os.path.exists(CALENDAR_CSV):
            self.df = pd.read_csv(CALENDAR_CSV, dtype={"year": str, "month": str, "day": str})
        else:
            self.df = pd.DataFrame(columns=[
                'year', 'month', 'day', 'timestamp', 'status',
                'cnDay', 'festival', 'is_sd', 'is_trading'
            ])

    def time_to_fmt(self, times, fmt='%Y-%m-%d %H:%M:%S'):
        """把时间戳转成日期格式"""
        now = int(times)
        time_arr = time.localtime(now)
        fmt_time = time.strftime(fmt, time_arr)
        return fmt_time

    def save_calendar(self):
        """保存 DataFrame 到 CSV 文件"""
        self.df.to_csv(CALENDAR_CSV, index=False)
        print(f"已保存至 {CALENDAR_CSV}")

    def catch_url_from_baidu(self, calcultaion_year, month):
        """从百度接口抓取指定年月的日历信息"""
        self.param['query'] = f"{calcultaion_year}年{month}月"
        print(f"抓取百度日历：{self.param['query']}...")

        r = req.get(
            url="https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php",
            headers=self.header,
            params=self.param
        ).text

        month_data = json.loads(r)["data"][0]["almanac"]

        for data in month_data:
            insert_data = {
                'year': str(data['year']),
                'month': str(data['month']),
                'day': str(data['day']),
                'timestamp': int(data['timestamp']),
                'status': int(data['status']) if 'status' in data else 0,
                'cnDay': data['cnDay'],
                'festival': data.get('festival', ''),
                'is_sd': 0,
                'is_trading': 0
            }

            # 判断是否已存在（按年月日查重）
            mask = (
                (self.df['year'] == insert_data['year']) &
                (self.df['month'] == insert_data['month']) &
                (self.df['day'] == insert_data['day'])
            )

            if mask.any():
                self.df.loc[mask, list(insert_data.keys())] = list(insert_data.values())
                print(f"{insert_data['year']}-{insert_data['month']}-{insert_data['day']} | 更新成功")
            else:
                self.df = pd.concat([self.df, pd.DataFrame([insert_data])], ignore_index=True)
                print(f"{insert_data['year']}-{insert_data['month']}-{insert_data['day']} | 新增成功")

    def find_trading(self):
        """标记交易日：工作日且非周六日"""
        df = self.df.copy()
        df = df[df['status'] == 0]
        df = df[~df['cnDay'].isin(['六', '日'])]

        for idx, row in df.iterrows():
            ts = int(row['timestamp'])
            date_obj = datetime.date.fromtimestamp(ts)
            ts_normalized = int(time.mktime(date_obj.timetuple()))
            self.df.loc[self.df['timestamp'] == ts, 'is_trading'] = 1

        print("交易日更新完成！")
        self.save_calendar()


if __name__ == '__main__':
    # 只执行一次，获取2023~2025的数据
    obj = Calendar()
    for year in [2019, 2020, 2021, 2022, 2023, 2024, 2025]: #可自定义需要抓取的年份数据
        for month in ["2", "5", "8", "11"]:
            obj.catch_url_from_baidu(year, month)
            time.sleep(1)

    obj.find_trading()
    print("历史节假日抓取完毕！运行时间：", obj.time_to_fmt(time.time()))
