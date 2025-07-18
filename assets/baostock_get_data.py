import baostock as bs
import pandas as pd
import time
from datetime import datetime

# 指数代码列表
index_codes = {
    '中证500': 'sh.000905',
    '中证1000': 'sh.000852',
    '沪深300': 'sh.000300'
}

# 要获取的字段
fields = "date,code,open,high,low,close,volume"

# 设置起止时间
start_date = '2024-01-01'
end_date = datetime.now().strftime('%Y-%m-%d')  # 每次用今天的日期

# 定义抓取一次数据的函数
def fetch_index_data():
    # 登录
    bs.login()
    all_data = pd.DataFrame()

    for name, code in index_codes.items():
        print(f"\n正在获取 {name}（{code}）的数据...")
        rs = bs.query_history_k_data_plus(code,
                                          fields,
                                          start_date=start_date,
                                          end_date=end_date,
                                          frequency="d",
                                          adjustflag="3")

        if rs.error_code != '0':
            print(f"获取失败：{rs.error_msg}")
            continue

        data = rs.get_data()
        if data.empty:
            print(f"{name}（{code}）无数据返回！")
            continue

        df = pd.DataFrame(data)
        df['指数名称'] = name
        all_data = pd.concat([all_data, df], ignore_index=True)
        print(f"{name} 数据获取成功，共 {len(df)} 条")

    # 登出
    bs.logout()

    # 保存
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"index_data.xlsx"
    all_data.to_excel(filename, index=False)
    print(f"\n本次数据保存为 {filename}")

# 每 5 分钟执行一次
interval_minutes = 5
print(f"\n开始定时抓取，每 {interval_minutes} 分钟执行一次（Ctrl+C 可停止）")

try:
    while True:
        print(f"\n当前时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        fetch_index_data()
        print(f"等待 {interval_minutes} 分钟后继续...\n")
        time.sleep(interval_minutes * 60)
except KeyboardInterrupt:
    print("\n手动停止程序")
