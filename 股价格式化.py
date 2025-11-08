import os
import pandas as pd
from glob import glob

# ==============================
# 配置文件路径
# ==============================
raw_data_dir = '股价原始数据'
output_data_dir = '股价整理数据'
os.makedirs(output_data_dir, exist_ok=True)

# ==============================
# 处理一个行业的原始文件
# ==============================
def process_industry_file(raw_file_path, output_dir):
    industry = os.path.basename(raw_file_path).replace('_股价.xlsx', '')
    df = pd.read_excel(raw_file_path, dtype={'trade_date': str})

    # 获取所有股票名，确保顺序一致
    stock_names = df['stock_name'].unique().tolist()
    stock_names.sort()
    
    # 按日期、股票透视，通过"收盘价"
    pivot_df = df.pivot_table(
        index='trade_date',
        columns='stock_name',
        values='close',
        aggfunc='first'
    )

    # 日期升序排列，并转为yyyy-mm-dd格式
    pivot_df.index = pd.to_datetime(pivot_df.index, format='%Y%m%d').strftime('%Y-%m-%d')
    pivot_df = pivot_df.sort_index(ascending=False)   # 最新日期在上

    # 列顺序按照stock_names
    pivot_df = pivot_df[stock_names]
    pivot_df.reset_index(inplace=True)
    pivot_df.rename(columns={'trade_date': '日期'}, inplace=True)

    # 列名中文不用处理，首列为日期
    pivot_df.rename(columns={'index': '日期'}, inplace=True)

    # 输出文件名
    out_path = os.path.join(output_dir, f'{industry}_股价整理.xlsx')
    pivot_df.to_excel(out_path, index=False)

# ==============================
# 处理目录下所有行业文件
# ==============================
def main():
    for file_path in glob(os.path.join(raw_data_dir, '*_股价.xlsx')):
        process_industry_file(file_path, output_data_dir)
    print("所有行业股价数据已整理完成。")

if __name__ == "__main__":
    main()