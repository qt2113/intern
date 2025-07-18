import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import schedule
import time
import os
#需操作读取指定路径的节假日日历数据；假设节假日日历数据已存储在 calendar_data.csv 中；具体的节假日数据需要运行get_calendar.py脚本获取
calendar_df = pd.read_csv(r"C:\Users\TQY\Desktop\净值检测\calendar_data.csv", dtype={'year': str, 'month': str, 'day': str})
calendar_df['date'] = pd.to_datetime(calendar_df['year'] + '-' + calendar_df['month'] + '-' + calendar_df['day'])
calendar_df['should_check_net_value'] = calendar_df['status'].isin([0, 2]) # 标注是否应检测净值的日子（status 为 0 或 2）

def load_data(file_path='net_value_data.xlsx'):
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
    else:
        raise ValueError("仅支持csv或xlsx格式")

    df.columns = df.columns.str.strip()
    df['净值日期'] = pd.to_datetime(df['净值日期'])
    df = df.dropna(subset=['标的代码', '净值日期'])
    return df


def infer_frequency_segmented(dates):
    """
    对净值日期进行滑动窗口频率推断，返回每段的频率及其时间段
    """
    dates = dates.sort_values().drop_duplicates()
    segments = []
    if len(dates) < 5:
        return [{'start': dates.min(), 'end': dates.max(), 'freq': '未知'}]

    window_size = 10
    step = 5
    for i in range(0, len(dates) - window_size + 1, step):
        window = dates[i:i + window_size]
        diffs = window.diff().dt.days.dropna()
        avg_gap = diffs.mean()
        if avg_gap <= 2:
            freq = '日频'
        elif 3 <= avg_gap <= 9:
            freq = '周频'
        elif avg_gap >= 20:
            freq = '月频'
        else:
            freq = '未知'
        segments.append({
            'start': window.min(),
            'end': window.max(),
            'freq': freq
        })

    # 合并相邻相同频率段
    merged = []
    for seg in segments:
        if not merged or merged[-1]['freq'] != seg['freq']:
            merged.append(seg)
        else:
            merged[-1]['end'] = seg['end']
    return merged


def check_completeness(group):
    code = group['标的代码'].iloc[0]
    group = group.sort_values('净值日期')
    dates = group['净值日期']
    today = pd.to_datetime('today').normalize()

    # 缺字段检查
    missing_fields = group[
        group[['单位净值', '累计净值', '净值日期', '标的代码']].isnull().any(axis=1)
    ]

    # 推断不同时间段的频率
    freq_segments = infer_frequency_segmented(dates)
    missing_dates = []

    for seg in freq_segments:
        freq = seg['freq']
        seg_start = seg['start']
        seg_end = seg['end']
        seg_dates = dates[(dates >= seg_start) & (dates <= seg_end)]

        valid_days = calendar_df[
            (calendar_df['date'] >= seg_start) &
            (calendar_df['date'] <= seg_end) &
            (calendar_df['should_check_net_value']) &
            (calendar_df['date'] <= today)
        ]['date']

        if freq == '日频':
            valid_days_filtered = valid_days[valid_days.dt.weekday < 5]
            for d in valid_days_filtered:
                if d not in seg_dates.values:
                    missing_dates.append(d.strftime('%Y-%m-%d'))

        elif freq == '周频':
            valid_weeks = valid_days.dt.to_period('W').unique()
            actual_weeks = seg_dates.dt.to_period('W').unique()
            for week in valid_weeks:
                if week not in actual_weeks:
                    week_days = valid_days[valid_days.dt.to_period('W') == week]
                    if (week_days.dt.weekday < 5).sum() >= 3:
                        missing_dates.append(week_days.max().strftime('%Y-%m-%d'))

        elif freq == '月频':
            valid_months = valid_days.dt.to_period('M').unique()
            actual_months = seg_dates.dt.to_period('M').unique()
            for month in valid_months:
                if month not in actual_months:
                    month_days = valid_days[valid_days.dt.to_period('M') == month]
                    if (month_days.dt.weekday < 5).sum() >= 7:
                        missing_dates.append(month_days.max().strftime('%Y-%m-%d'))

        # 其他频率不做处理

    return {
        '标的代码': code,
        '频率段数': len(freq_segments),
        '缺失字段记录数': len(missing_fields),
        '缺失日期数': len(missing_dates),
        '缺失日期示例': missing_dates[:10]
    }


def main_check(file_path, target_code=None, output_path=None):
    print(f"\n[{datetime.now()}] 净值完整性检测开始")
    df = load_data(file_path)

    if target_code:
        if target_code not in df['标的代码'].unique():
            print(f"❌ 标的代码 '{target_code}' 不存在于数据中")
            return
        df = df[df['标的代码'] == target_code]

    all_codes = df['标的代码'].unique()
    incomplete_list = []

    for code in all_codes:
        group = df[df['标的代码'] == code].copy()
        result = check_completeness(group)
        print(f"\n标的代码：{result['标的代码']}")
        print(f"频率段数：{result['频率段数']}")
        if result['缺失字段记录数'] > 0 or result['缺失日期数'] > 0:
            print(f"❗缺失字段记录数：{result['缺失字段记录数']}")
            print(f"❗缺失日期数：{result['缺失日期数']}")
            print(f"❗缺失日期示例：{result['缺失日期示例']}")
            incomplete_list.append(result)
        else:
            print("✅ 净值数据完整")

    print("\n--- 检测结束 ---")

    if incomplete_list:
        out_df = pd.DataFrame(incomplete_list)
        out_df['缺失日期示例'] = out_df['缺失日期示例'].apply(lambda x: ', '.join(x))

        if output_path is None:
            base_dir = os.path.dirname(file_path)
            output_path = os.path.join(base_dir, '净值异常数据.xlsx')
        out_df.to_excel(output_path, index=False)
        print(f"\n缺失信息已保存至：{output_path}")
    else:
        print("\n所有标的净值数据完整，无缺失。")
        # 若无异常数据，仍然覆盖输出文件为空表
        if output_path is None:
            base_dir = os.path.dirname(file_path)
            output_path = os.path.join(base_dir, '净值异常数据.xlsx')
        pd.DataFrame(columns=['标的代码', '频率段数', '缺失字段记录数', '缺失日期数', '缺失日期示例']).to_excel(output_path, index=False)
        print(f"\n输出文件已清空：{output_path}")


def setup_schedule(interval_min=5, file_path='net_value_data.xlsx'):
    base_dir = os.path.dirname(file_path)
    output_path = os.path.join(base_dir, '净值异常数据.xlsx')
    schedule.every(interval_min).minutes.do(main_check, file_path=file_path, output_path=output_path)
    print(f"\n[定时检测已启动，每{interval_min}分钟执行一次]")
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == '__main__':
    # 启动一次立即检测
    # main_check(file_path = r"C:\Users\TQY\Desktop\to_goods_net.xlsx", target_code='', output_path=r"C:\Users\TQY\Desktop\净值异常数据.xlsx") # 检测特定标的代码
    main_check(file_path = r"C:\Users\TQY\Desktop\净值检测\to_goods_net.xlsx", output_path=r"C:\Users\TQY\Desktop\净值检测\净值异常数据.xlsx")  #检测所有标的代码

    #然后每5分钟定时执行
    setup_schedule(5, r"C:\Users\TQY\Desktop\净值检测\to_goods_net.xlsx")
