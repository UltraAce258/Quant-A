import os
import pandas as pd
import tushare as ts

# 设定工作目录
work_dir = os.path.abspath(os.curdir)
os.chdir(work_dir)

# 配置文件夹路径
stock_folder = os.path.join(work_dir, '涉及的股票')
price_folder = os.path.join(work_dir, '股价原始数据')
os.makedirs(price_folder, exist_ok=True)

# tushare初始化
ts.set_token('1770cab9e56c81e69e188f6ff930d3060d79a6db63bca59dcf7c3f65')  # 替换为你的tushare token
pro = ts.pro_api()

# 需要获取的日期
target_dates = [
    '2024-12-31', '2024-09-30', '2024-06-28', '2024-03-29',
    '2023-12-29', '2023-09-28', '2023-06-30', '2023-03-31',
    '2022-12-30', '2022-09-30', '2022-06-30', '2022-03-31',
    '2021-12-31', '2021-09-30', '2021-06-30', '2021-03-31',
    '2020-12-31', '2020-09-30', '2020-06-30', '2020-03-31'
]

# 转换为交易所格式日历日期
def find_trade_date(date_str):
    """
    对于不是交易日的日期，向前寻找最近交易日
    """
    trade_cal = pro.trade_cal(exchange='SSE', start_date='20200101', end_date='20251231')
    trade_cal = trade_cal[trade_cal['is_open']==1]['cal_date'].sort_values()
    date = pd.to_datetime(date_str)
    trade_dates = pd.to_datetime(trade_cal.values)
    date_use = trade_dates[trade_dates <= date].max()
    return date_use.strftime('%Y%m%d')

target_trade_dates = [find_trade_date(d) for d in target_dates]

# 准备股票代码查找（名称->代码）
def get_stock_code_map():
    # 取全市场A股
    stocks = pro.stock_basic(exchange='', list_status='L', fields='ts_code,name')
    return dict(zip(stocks['name'], stocks['ts_code']))

stock_name2code = get_stock_code_map()

# 遍历‘涉及的股票’下的所有行业表
for fname in os.listdir(stock_folder):
    if fname.endswith('.xlsx'):
        industry = fname.split('_')[0]
        fpath = os.path.join(stock_folder, fname)
        df = pd.read_excel(fpath)
        result = []
        # 对每个股票逐个查找
        for idx, row in df.iterrows():
            stock_name = str(row['证券名称'])
            ts_code = stock_name2code.get(stock_name)
            if not ts_code:
                print(f"[WARN] 找不到股票代码: {stock_name}")
                continue
            for d in target_trade_dates:
                try:
                    price_df = pro.daily(ts_code=ts_code, trade_date=d)
                    if not price_df.empty:
                        price_df['industry'] = industry
                        price_df['stock_name'] = stock_name
                        result.append(price_df)
                except Exception as e:
                    print(f"[ERROR] 抓取 {ts_code} {stock_name} {d} 异常: {e}")

        if result:
            price_res = pd.concat(result, ignore_index=True)
            out_path = os.path.join(price_folder, f"{industry}_股价.xlsx")
            price_res.to_excel(out_path, index=False)
            print(f"[INFO] {industry} 行业股价已保存至 {out_path}")
        else:
            print(f"[INFO] {industry} 行业没有抓到有效数据")

print("全部行业股票抓取完毕！")