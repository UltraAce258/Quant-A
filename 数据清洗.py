import os
import pandas as pd
import re

# --- 参数配置 ---
# 输入和输出文件夹名称
INPUT_DIR = "原始数据"
OUTPUT_DIR = "初步清洗"

# 清洗阈值 (可根据经验调整)
# 1. 指标列删除阈值：如果一个指标（如ROA）在整个行业（所有股票、所有季度）中的数据点缺失超过80%，则删除该指标对应的所有列
INDICATOR_DROP_THRESHOLD = 0.8 
# 2. 个股行删除阈值：在删除了无用指标后，如果一只股票的所有剩余数据点中，有超过50%是缺失的，则剔除该股票
STOCK_DROP_THRESHOLD = 0.5


def get_base_indicator_name(column_name):
    """
    从复杂的列名中提取基础指标名称。
    例如: "净资产收益率ROE\n[报告期] 2020一季..." -> "净资产收益率ROE"
    """
    if column_name in ["证券代码", "证券名称"]:
        return column_name
    # 移除换行符和方括号内的所有内容
    base_name = re.sub(r'\n\[.*?\]', '', column_name, flags=re.DOTALL)
    # 移除末尾的单位信息等
    base_name = re.sub(r'\n.*', '', base_name)
    return base_name.strip()

def clean_data(file_path):
    """
    对单个行业数据文件进行清洗。
    """
    print(f"\n--- 正在处理文件: {os.path.basename(file_path)} ---")

    # 读取数据，并将 "--" 视为空值
    try:
        df = pd.read_excel(file_path, na_values="--")
    except FileNotFoundError:
        print(f"错误: 文件未找到 {file_path}")
        return None
        
    print(f"原始数据维度: {df.shape[0]}只股票, {df.shape[1]}列")

    # 1. (要求 3) 移除ST股
    initial_stock_count = len(df)
    df = df[~df['证券名称'].str.contains('ST', na=False)].copy()
    print(f"步骤1: 移除了 {initial_stock_count - len(df)} 只ST股票。剩余: {len(df)}只。")

    # 2. (新的要求 4) 如果某个指标在全行业普遍缺失，则删除该指标的所有列
    print(f"步骤2: 移除全行业普遍缺失的指标列 (阈值 > {INDICATOR_DROP_THRESHOLD:.0%})")
    
    # 提取所有数据列的基础指标名称
    data_columns = df.columns[2:]
    base_indicator_map = {col: get_base_indicator_name(col) for col in data_columns}
    unique_indicators = sorted(list(set(base_indicator_map.values())))
    
    cols_to_drop = []
    for indicator in unique_indicators:
        # 找到属于当前基础指标的所有列
        indicator_cols = [col for col, base_name in base_indicator_map.items() if base_name == indicator]
        
        # 提取这个指标的数据块
        indicator_block = df[indicator_cols]
        
        # 计算整个块的缺失率
        total_cells = indicator_block.size
        missing_cells = indicator_block.isnull().sum().sum()
        block_missing_rate = missing_cells / total_cells if total_cells > 0 else 0
        
        # 如果缺失率超过阈值，则将这些列加入待删除列表
        if block_missing_rate > INDICATOR_DROP_THRESHOLD:
            print(f"  - 指标 '{indicator}' 缺失率高达 {block_missing_rate:.1%}，将被从该行业数据中移除。")
            cols_to_drop.extend(indicator_cols)
            
    # 一次性删除所有标记的列
    if cols_to_drop:
        df.drop(columns=cols_to_drop, inplace=True)
        print(f"  > 共移除了 {len(cols_to_drop)} 个指标列。")

    # 3. (要求 5) 在此基础上，如果个股缺失数据过多，则移除该个股
    print(f"步骤3: 移除数据缺失严重的个股 (阈值 > {STOCK_DROP_THRESHOLD:.0%})")
    
    # 计算每只股票在所有剩余指标列中的总缺失率
    if len(df.columns) > 2:
        total_missing_rate = df.iloc[:, 2:].isnull().mean(axis=1)
        
        # 找出要移除的股票
        stocks_to_remove = total_missing_rate[total_missing_rate > STOCK_DROP_THRESHOLD].index
        
        if not stocks_to_remove.empty:
            print(f"  - 发现 {len(stocks_to_remove)} 只股票在剩余指标中数据质量过低，将被移除。")
            df = df.drop(index=stocks_to_remove)
    else:
        print("  - 没有足够的数据列进行个股缺失率检查。")

    print(f"清洗后数据维度: {df.shape[0]}只股票, {df.shape[1]}列")
    return df


def main():
    """
    主函数：设置目录，遍历文件，执行清洗并保存。
    """
    # 2. 设置工作目录 (在现代环境中通常是自动的，但显式设置更稳健)
    base_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
    os.chdir(base_dir)

    input_path = os.path.join(base_dir, INPUT_DIR)
    output_path = os.path.join(base_dir, OUTPUT_DIR)

    if not os.path.isdir(input_path):
        print(f"错误: 输入文件夹 '{INPUT_DIR}' 不存在。请在脚本同目录下创建它并放入数据文件。")
        return

    os.makedirs(output_path, exist_ok=True)
    print(f"数据将从 '{INPUT_DIR}' 读取，清洗后存入 '{OUTPUT_DIR}'。")

    # 1. 遍历输入目录中的所有xlsx文件
    for filename in os.listdir(input_path):
        if filename.endswith(".xlsx") and not filename.startswith("~"):
            file_path = os.path.join(input_path, filename)
            
            cleaned_df = clean_data(file_path)
            
            if cleaned_df is not None and not cleaned_df.empty:
                # 6. 构建新的文件名并保存
                industry_name = filename.replace(".xlsx", "")
                output_filename = f"{industry_name}_清洗后.xlsx"
                output_file_path = os.path.join(output_path, output_filename)
                
                cleaned_df.to_excel(output_file_path, index=False)
                print(f"成功保存清洗后的文件到: {output_file_path}")
            else:
                print(f"文件 {filename} 清洗后为空，不进行保存。")

if __name__ == "__main__":
    main()