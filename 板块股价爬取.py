import os
import pandas as pd
import akshare as ak

# ------------ 配置区域 --------------
work_dir = os.path.abspath(os.curdir)
os.chdir(work_dir)
output_folder = os.path.join(work_dir, '板块原始数据')
os.makedirs(output_folder, exist_ok=True)

# 板块中文名（与同花顺官网/akshare一致）
industry_names = ["通信服务", "通信设备", "银行", "非银行金融"]

# 目标日期（为筛选用，最终数据全部日K，不论是否是最后一个交易日，后续可按此筛选）
target_dates = [
    '2024-12-31','2024-09-30','2024-06-28','2024-03-29',
    '2023-12-29','2023-09-28','2023-06-30','2023-03-31',
    '2022-12-30','2022-09-30','2022-06-30','2022-03-31',
    '2021-12-31','2021-09-30','2021-06-30','2021-03-31',
    '2020-12-31','2020-09-30','2020-06-30','2020-03-31'
]
target_dates_set = set(target_dates)

# ------- 工具函数 --------
def find_board_code(name: str):
    """
    自动用合适的akshare接口，获取同花顺行业板块代码
    """
    import akshare as ak
    try:
        df = ak.ths_board_industry_listing()
    except AttributeError:
        df = ak.stock_board_industry_name_ths()
    row = df[df['名称'] == name]
    if row.empty:
        print(f"[WARN] 未找到板块：{name}")
        return None
    return row.iloc[0]['代码']

# -------- 主程序 --------
for ind_name in industry_names:
    print(f"\n【{ind_name}】——获取板块指数行情 ...")
    board_code = find_board_code(ind_name, board_type="industry")
    if not board_code:
        print(f"[WARN] {ind_name} 无法获取代码，跳过")
        continue

    try:
        df = ak.ths_index_daily(symbol=board_code)  # 全部历史日K
        if not df.empty:
            # 仅保留目标日期的记录，若你只关心特定日期
            df['date'] = pd.to_datetime(df['日期'], errors='coerce').dt.strftime('%Y-%m-%d')
            # 若仅需要target_dates
            df_select = df[df['date'].isin(target_dates_set)].copy()
            out_path = os.path.join(output_folder, f"{ind_name}_板块指数日K.xlsx")
            df.to_excel(out_path, index=False)
            print(f"  [INFO] 板块【{ind_name}】已保存全部日K（含所有历史日期）至：{out_path}")
            out_path2 = os.path.join(output_folder, f"{ind_name}_板块指数日K_目标日期筛选.xlsx")
            df_select.to_excel(out_path2, index=False)
            print(f"  [INFO] 板块【{ind_name}】目标交易日数据另存为：{out_path2}")
        else:
            print(f"  [WARN] 板块【{ind_name}】行情为空")
    except Exception as e:
        print(f"  [ERROR] 板块【{ind_name}】行情获取出错: {e}")

print("全部同花顺行业板块指数行情抓取完毕！")