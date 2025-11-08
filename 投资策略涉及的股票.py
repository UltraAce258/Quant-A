import os
import pandas as pd
from collections import Counter
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontManager

def set_chinese_font():
    """
    自动查找并设置matplotlib的中文字体。
    """
    try:
        # 常见中文字体列表
        font_names = ['SimHei', 'Microsoft YaHei', 'PingFang SC', 'Heiti TC', 'WenQuanYi Micro Hei']
        fm = FontManager()
        
        # 查找系统中可用的中文字体
        available_fonts = [f.name for f in fm.ttflist]
        found_font = None
        for font_name in font_names:
            if font_name in available_fonts:
                found_font = font_name
                break
        
        if found_font:
            print(f"找到并设置中文字体: {found_font}")
            plt.rcParams['font.sans-serif'] = [found_font]
        else:
            print("警告：未找到推荐的中文字体（如黑体、微软雅黑）。图表中的中文可能无法正常显示。")
            print("请尝试安装 'SimHei' 或 'Microsoft YaHei' 字体。")
            # 使用系统默认的 sans-serif 字体作为备选
            plt.rcParams['font.sans-serif'] = ['sans-serif']

        # 解决负号显示问题
        plt.rcParams['axes.unicode_minus'] = False
    except Exception as e:
        print(f"设置中文字体时出错: {e}")
        print("将使用默认字体，中文可能无法显示。")

def create_and_save_plot(data_df, industry, output_path):
    """
    根据DataFrame数据创建并保存柱状图。

    :param data_df: 包含'证券名称'和'出现次数'的DataFrame。
    :param industry: 行业名称。
    :param output_path: 图表保存路径。
    """
    if data_df.empty:
        print(f"行业 '{industry}' 数据为空，跳过图表生成。")
        return

    # 为了图表美观，只展示出现次数最多的前30只股票
    plot_df = data_df.head(30)

    plt.figure(figsize=(16, 9))
    plt.bar(plot_df['证券名称'], plot_df['出现次数'], color='skyblue')

    plt.title(f'{industry} 行业股票出现次数排行', fontsize=20, pad=20)
    plt.xlabel('证券名称', fontsize=14)
    plt.ylabel('出现次数', fontsize=14)
    plt.xticks(rotation=45, ha='right')  # 旋转X轴标签以防重叠
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()  # 自动调整布局

    try:
        plt.savefig(output_path, dpi=300)
        print(f"图表已保存到: {output_path}")
    except Exception as e:
        print(f"保存图表 '{output_path}' 时出错: {e}")
    
    plt.close() # 关闭当前图形，释放内存


def analyze_stocks_by_industry(base_path):
    """
    分析行业股票数据，输出Excel统计表和可视化图表。

    :param base_path: 工作目录的根路径。
    """
    # 1. 定义输入和输出目录
    input_dir = os.path.join(base_path, '因子分析量化策略研究')
    excel_output_dir = os.path.join(base_path, '涉及的股票')
    plot_output_dir = os.path.join(base_path, '涉及的股票_图表')

    # 2. 检查输入目录
    if not os.path.isdir(input_dir):
        print(f"错误：输入目录 '{input_dir}' 不存在。")
        return

    # 3. 创建输出目录
    os.makedirs(excel_output_dir, exist_ok=True)
    os.makedirs(plot_output_dir, exist_ok=True)
    print(f"Excel输出目录: {os.path.abspath(excel_output_dir)}")
    print(f"图表输出目录: {os.path.abspath(plot_output_dir)}")

    # 4. 读取和汇总数据
    industry_stocks = {}
    for filename in os.listdir(input_dir):
        if filename.endswith(('.xlsx', '.xls')):
            try:
                industry_name = filename.split('_')[0]
                file_path = os.path.join(input_dir, filename)
                print(f"正在处理文件: {filename}，行业: {industry_name}")
                df = pd.read_excel(file_path)

                if '证券名称' in df.columns:
                    if industry_name not in industry_stocks:
                        industry_stocks[industry_name] = []
                    industry_stocks[industry_name].extend(df['证券名称'].dropna().tolist())
                else:
                    print(f"警告：文件 '{filename}' 中未找到 '证券名称' 列。")
            except Exception as e:
                print(f"处理文件 '{filename}' 时出错: {e}")

    # 5. 处理每个行业的数据
    if not industry_stocks:
        print("未找到任何股票数据进行处理。")
        return

    # 设置中文字体
    set_chinese_font()

    for industry, stocks in industry_stocks.items():
        # 统计频率
        stock_counts = Counter(stocks)
        result_df = pd.DataFrame(stock_counts.items(), columns=['证券名称', '出现次数'])
        result_df = result_df.sort_values(by='出现次数', ascending=False).reset_index(drop=True)
        
        # --- 保存Excel文件 ---
        excel_filename = f"{industry}_涉及的股票.xlsx"
        excel_filepath = os.path.join(excel_output_dir, excel_filename)
        print(f"正在将行业 '{industry}' 的统计结果写入到: {excel_filepath}")
        result_df.to_excel(excel_filepath, index=False, engine='openpyxl')
        
        # --- 创建并保存图表 ---
        plot_filename = f"{industry}_股票频率图.png"
        plot_filepath = os.path.join(plot_output_dir, plot_filename)
        create_and_save_plot(result_df, industry, plot_filepath)
    
    print("\n所有文件处理完成！")


if __name__ == '__main__':
    # 将当前目录设置为工作目录
    work_directory = os.getcwd()
    print(f"当前工作目录已设置为: {work_directory}")
    
    # 执行主函数
    analyze_stocks_by_industry(work_directory)