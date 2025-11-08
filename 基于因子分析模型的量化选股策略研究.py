import os
import pandas as pd
import numpy as np
import re
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import FactorAnalysis
import matplotlib.pyplot as plt
import seaborn as sns
from pandas.plotting import table

# ==============================================================================
# --- 1. 全局配置区 ---
# ==============================================================================
FUNDAMENTAL_DATA_DIR = "初步清洗"
PRICE_DATA_DIR = "股价整理数据"
OUTPUT_PROJECT_NAME = "因子分析量化策略研究"
BACKTEST_START_DATE = "2021-03-31"
BACKTEST_END_DATE = "2024-12-31" 
INITIAL_CAPITAL = 1_000_000

TOP_N_STOCKS = 5

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# ==============================================================================
# --- 2. 核心功能函数 ---
# ==============================================================================

def rolling_factor_analysis_and_ranking(window_data: pd.DataFrame, min_cum_var: float = 0.8, max_factor: int = 10):
    # max_factor 可根据变量/经验选合适上限
    stock_info = window_data[['证券名称']].copy()
    features = window_data.drop(columns=['证券名称'])
    n_variables = features.shape[1]
    fit_factor = min(max_factor, n_variables)

    scaler = StandardScaler()
    features_scaled = np.nan_to_num(scaler.fit_transform(features))
    
    # 先最多做 max_factor 个
    fa = FactorAnalysis(n_components=fit_factor, rotation='varimax', random_state=0)
    try:
        factor_scores = fa.fit_transform(features_scaled)
        eig_vals = np.sum(fa.components_**2, axis=1)
        contrib_ratio = eig_vals / np.sum(eig_vals)
        cumsum = np.cumsum(contrib_ratio)
        n_factors = np.searchsorted(cumsum, min_cum_var) + 1  # 满足累计贡献率的最小个数
        # 最终再fit一次
        fa = FactorAnalysis(n_components=n_factors, rotation='varimax', random_state=0)
        factor_scores = fa.fit_transform(features_scaled)
        eig_vals = np.sum(fa.components_**2, axis=1)
        weights = eig_vals / np.sum(eig_vals)
    except Exception:
        return None, None

    total_scores = np.dot(factor_scores, weights)
    loadings_df = pd.DataFrame(fa.components_.T, index=features.columns,
                               columns=[f'因子{i+1}' for i in range(n_factors)])
    ranked_stocks = stock_info.copy()
    ranked_stocks['综合得分'] = total_scores
    ranked_stocks.sort_values(by='综合得分', ascending=False, inplace=True)
    return loadings_df, ranked_stocks.reset_index(drop=True)

def get_base_indicator_name(column_name):
    return re.sub(r'\n.*', '', column_name, flags=re.DOTALL).strip()

def backtest_and_analyze(industry_name: str, fundamental_df: pd.DataFrame, price_df: pd.DataFrame = None):
    print(f"\n{'='*25} 开始处理【{industry_name}】行业 {'='*25}")

    # 初始化存储器
    all_loadings, all_top_stocks, portfolio_history = {}, {}, {}
    
    # 预处理股价数据：列名、索引全部为“证券名称”及“日期”
    price_ts = None
    if price_df is not None:
        try:
            price_ts = price_df.copy()
            price_ts['日期'] = pd.to_datetime(price_ts['日期'])
            price_ts.set_index('日期', inplace=True)
            price_ts.sort_index(inplace=True)
        except Exception as e:
            print(f"    - 警告: 股价数据预处理失败: {e}。将仅执行因子分析。")
            price_ts = None

    # 回测变量
    current_portfolio = {}
    cash = INITIAL_CAPITAL
    quarters = pd.to_datetime(pd.date_range(start=BACKTEST_START_DATE, end=BACKTEST_END_DATE, freq='QS-JAN'))

    for i, trade_date in enumerate(quarters):
        quarter_str = f"{trade_date.year}Q{trade_date.quarter}"
        print(f"\n--- 季度: {quarter_str} ---")

        # --- 1. 期初资产 ---
        if price_ts is not None:
            try:
                actual_trade_date = price_ts.index[price_ts.index >= trade_date][0]
                prices_on_trade_date = price_ts.loc[actual_trade_date]
            except IndexError:
                print(f"    - 警告: 找不到 {trade_date.date()} 或之后的股价，无法交易。")
                continue

            if i == 0:
                total_asset_value = INITIAL_CAPITAL
            else:
                stock_value = sum(shares * prices_on_trade_date.get(name, 0) for name, shares in current_portfolio.items())
                total_asset_value = stock_value + cash
            
            portfolio_history[quarter_str] = {'start_asset': total_asset_value}
            print(f"    - 期初总资产: {total_asset_value:,.2f}")
            cash = total_asset_value
            current_portfolio = {}

        # --- 2. 因子分析与选股 ---
        t2 = trade_date - pd.DateOffset(months=6)
        t5 = t2 - pd.DateOffset(months=9)
        
        window_cols = []
        report_map = {"一季": 3, "中报": 6, "三季": 9, "年报": 12}
        for col in fundamental_df.columns[1:]:  # 跳过“证券名称”
            clean_col = col.replace('\n', ' ')
            match = re.search(r'(\d{4}).*?(一季|中报|三季|年报)', clean_col)
            if match:
                year, report_type = match.groups()
                month = report_map.get(report_type)
                if year and month:
                    report_date = pd.to_datetime(f'{year}-{month}-01') + pd.offsets.MonthEnd(0)
                    if t5 <= report_date <= t2:
                        window_cols.append(col)

        if not window_cols:
            print("    - 警告: 找不到足够的财务数据，跳过本季度。")
            continue

        base_cols = ['证券名称']
        window_df = fundamental_df[base_cols + window_cols].copy()
        for col in window_cols:
            window_df[col] = pd.to_numeric(window_df[col], errors='coerce')
        
        mean_df = window_df[base_cols].copy()
        base_indicators = {get_base_indicator_name(col) for col in window_cols}
        for indicator in sorted(list(base_indicators)):
            cols_for_mean = [col for col in window_cols if get_base_indicator_name(col) == indicator]
            mean_df[indicator] = window_df[cols_for_mean].mean(axis=1)

        loadings, ranked_stocks = rolling_factor_analysis_and_ranking(mean_df.dropna())

        if loadings is not None and ranked_stocks is not None:
            all_loadings[quarter_str] = loadings
            all_top_stocks[quarter_str] = ranked_stocks.head(TOP_N_STOCKS)
            print(f"    - 本期选股 (Top {TOP_N_STOCKS}):\n{ranked_stocks.head(TOP_N_STOCKS)[['证券名称']].to_string(index=False)}")
        else:
            print("    - 警告: 本期因子分析失败，无选股结果。")
            continue

        # --- 3. 建立新持仓 ---
        if price_ts is not None:
            if trade_date != quarters[-1]:
                target_names = all_top_stocks[quarter_str]['证券名称'].tolist()
                tradable_names = [name for name in target_names
                                  if name in prices_on_trade_date.index or name in prices_on_trade_date.index.tolist() or name in prices_on_trade_date.keys()]
                tradable_names = [name for name in target_names
                                  if name in prices_on_trade_date.index or name in prices_on_trade_date.keys()]
                # 保险：用keys()兼容Series类型
                tradable_names = [name for name in target_names
                                  if name in prices_on_trade_date.index or name in prices_on_trade_date.keys()]
                # 并且价格不为nan且大于0
                tradable_names = [name for name in tradable_names
                                  if pd.notna(prices_on_trade_date.get(name, np.nan)) and prices_on_trade_date.get(name, 0) > 0]
                if tradable_names:
                    investment_per_stock = cash / len(tradable_names)
                    for name in tradable_names:
                        price = prices_on_trade_date.get(name, np.nan)
                        shares = investment_per_stock / price
                        current_portfolio[name] = shares
                        cash -= investment_per_stock
                    print(f"    - 建立新持仓后，剩余现金: {cash:,.2f}")
                else:
                    print("    - 警告: 目标股票均无法交易，本季度空仓。")

    # --- 4. 最终清算 ---
    if price_ts is not None and portfolio_history:
        last_q_str = f"{quarters[-1].year}Q{quarters[-1].quarter}"
        final_asset_value = cash
        if current_portfolio:
            try:
                final_prices = price_ts.iloc[-1]
                final_proceeds = sum(shares * final_prices.get(name, 0) for name, shares in current_portfolio.items())
                final_asset_value = cash + final_proceeds
            except IndexError:
                pass
        
        if last_q_str in portfolio_history:
            portfolio_history[last_q_str]['end_asset'] = final_asset_value
        print(f"\n回测结束，最终资产价值: {final_asset_value:,.2f}")

    return all_loadings, all_top_stocks, portfolio_history

def generate_visualizations(industry_name: str, loadings: dict, top_stocks: dict, history: dict):
    print(f"\n--- 为【{industry_name}】生成图表 ---")
    output_dir = os.path.join(os.getcwd(), OUTPUT_PROJECT_NAME)
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. 收益回测明细
    if history:
        history_df = pd.DataFrame.from_dict(history, orient='index').reset_index().rename(columns={'index': '季度', 'start_asset': '期初总资产'})
        history_df['季度收益率(%)'] = (history_df['期初总资产'].pct_change(fill_method=None) * 100).fillna(0)
        history_df['期末总资产'] = [history[q].get('end_asset') for q in history_df['季度']]
        history_df = history_df[['季度', '期初总资产', '期末总资产', '季度收益率(%)']]
        excel_path = os.path.join(output_dir, f"{industry_name}_收益回测明细.xlsx")
        history_df.to_excel(excel_path, index=False, float_format="%.2f")
        print(f"  - 已保存收益回测明细: {excel_path}")

    # 2. 因子载荷矩阵
    if loadings:
        for year in sorted({q[:4] for q in loadings.keys()}):
            quarters_in_year = [q for q in loadings if q.startswith(year)]
            fig, axes = plt.subplots(1, len(quarters_in_year), figsize=(5 * len(quarters_in_year), 8), squeeze=False)
            fig.suptitle(f'【{industry_name}】{year}年 因子载荷矩阵', fontsize=16)
            for i, q in enumerate(quarters_in_year):
                sns.heatmap(loadings[q], ax=axes[0, i], cmap='viridis', annot=True, fmt=".2f")
                axes[0, i].set_title(q)
            plt.tight_layout(rect=[0, 0.03, 1, 0.95])
            plt.savefig(os.path.join(output_dir, f"{industry_name}_{year}年_因子载荷矩阵.png"))
            plt.close(fig)
        print(f"  - 已保存因子载荷矩阵图。")

    # 3. 选股策略
    if top_stocks:
        result_df = pd.concat([df.assign(季度=q) for q, df in top_stocks.items()])[['季度', '证券名称', '综合得分']]
        result_df.to_excel(os.path.join(output_dir, f"{industry_name}_每季度选股策略.xlsx"), index=False)
        fig, ax = plt.subplots(figsize=(12, max(5, 0.4 * len(result_df))))
        ax.axis('off'); ax.set_title(f'【{industry_name}】每季度选股策略 (Top {TOP_N_STOCKS})', fontsize=16, pad=20)
        tbl = table(ax, result_df, loc='center', cellLoc='center', colWidths=[0.15, 0.2, 0.2])
        tbl.auto_set_font_size(False); tbl.set_fontsize(10); tbl.scale(1.2, 1.2)
        plt.savefig(os.path.join(output_dir, f"{industry_name}_每季度选股策略.png"), bbox_inches='tight', pad_inches=0.1)
        plt.close(fig)
        print(f"  - 已保存选股策略图表。")
        
    # 4. 净值曲线
    if history and len(history) > 1:
        sorted_q = sorted(history.keys())
        dates = [pd.to_datetime(f"{q[:4]}-{ (int(q[5:])-1)*3 + 1 }-01") for q in sorted_q]
        net_values = [history[q]['start_asset'] / INITIAL_CAPITAL for q in sorted_q]
        if 'end_asset' in history[sorted_q[-1]] and pd.notna(history[sorted_q[-1]]['end_asset']):
            dates.append(dates[-1] + pd.DateOffset(months=3))
            net_values.append(history[sorted_q[-1]]['end_asset'] / INITIAL_CAPITAL)
        plt.figure(figsize=(14, 7))
        plt.plot(dates, net_values, marker='o', linestyle='-')
        total_return = (net_values[-1] - 1) * 100
        plt.title(f'【{industry_name}】策略净值曲线 (总收益率: {total_return:.2f}%)', fontsize=16)
        plt.xlabel('日期'); plt.ylabel('策略净值 (初始为1)'); plt.grid(True); plt.tight_layout()
        plt.savefig(os.path.join(output_dir, f"{industry_name}_策略净值曲线.png"))
        plt.close()
        print(f"  - 已保存策略净值曲线图。")

def plot_multi_industry_nav_comparison(output_project_folder, initial_capital=INITIAL_CAPITAL):
    """
    汇总output_project_folder中所有“_收益回测明细.xlsx”，
    绘制多行业策略净值对比曲线，图片保存到同文件夹
    """
    import glob
    color_list = plt.get_cmap('tab10').colors
    plt.figure(figsize=(13, 7))
    files = glob.glob(os.path.join(output_project_folder, "*_收益回测明细.xlsx"))
    legend = []
    for i, path in enumerate(files):
        try:
            df = pd.read_excel(path)
            industry = os.path.basename(path).replace("_收益回测明细.xlsx", "")
            # 用期初总资产为净值
            dates = df['季度']
            nav = df['期初总资产'] / initial_capital
            if isinstance(dates.iloc[0], float):  # 防止是数字
                dates = [f'{int(str(d)[:4])}Q{int(str(d)[5:])}' for d in df['季度']]
            plt.plot(dates, nav, marker='o', label=industry, color=color_list[i % len(color_list)])
            legend.append(industry)
        except Exception as e:
            print(f"【多行业净值曲线】{path} 处理失败: {e}")
            continue
    plt.title("多行业策略净值对比曲线", fontsize=17)
    plt.xlabel("季度")
    plt.ylabel("策略净值(初始为1)")
    plt.xticks(rotation=40)
    plt.legend(loc="best")
    plt.tight_layout()
    save_path = os.path.join(output_project_folder, "多行业策略净值对比曲线.png")
    plt.savefig(save_path)
    plt.close()
    print(f"【多行业净值对比】已保存：{save_path}")


    
def main():
    if not os.path.isdir(FUNDAMENTAL_DATA_DIR):
        print(f"错误: 财务数据文件夹 '{FUNDAMENTAL_DATA_DIR}' 不存在。"); return

    for filename in os.listdir(FUNDAMENTAL_DATA_DIR):
        if filename.endswith("_清洗后.xlsx") and not filename.startswith("~"):
            industry_name = filename.replace("_清洗后.xlsx", "")
            fundamental_file_path = os.path.join(FUNDAMENTAL_DATA_DIR, filename)
            price_file_path = os.path.join(PRICE_DATA_DIR, f"{industry_name}_股价整理.xlsx")
            
            try:
                fundamental_df = pd.read_excel(fundamental_file_path)
            except Exception as e:
                print(f"读取文件 {filename} 失败: {e}"); continue
            
            price_df = None
            if os.path.exists(price_file_path):
                try:
                    price_df = pd.read_excel(price_file_path)
                    print(f"成功加载【{industry_name}】的股价数据。")
                except Exception as e:
                    print(f"警告: 读取股价文件 {price_file_path} 失败: {e}")
            else:
                print(f"提示: 未找到【{industry_name}】的股价数据文件。")

            loadings, top_stocks, history = backtest_and_analyze(industry_name, fundamental_df, price_df)
            generate_visualizations(industry_name, loadings, top_stocks, history)

    plot_multi_industry_nav_comparison(OUTPUT_PROJECT_NAME)

if __name__ == "__main__":
    main()