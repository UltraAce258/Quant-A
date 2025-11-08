import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os

# ==============================================================================
# --- 1. 环境设置 ---
# ==============================================================================
try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print(f"工作目录已成功设置为: {os.getcwd()}")
except NameError:
    print(f"在交互式环境中运行，使用当前工作目录: {os.getcwd()}")

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# ==============================================================================
# --- 2. 核心绘图函数 (全新重构) ---
# ==============================================================================
def draw_custom_table(ax, data, year_title):
    """手动在给定的子图(ax)上绘制一个美化的表格"""
    ax.axis('off')
    ax.set_title(f"--- {year_title} ---", fontsize=18, pad=20, weight='bold')

    if data.empty:
        ax.text(0.5, 0.5, '本年度无数据', ha='center', va='center', fontsize=14)
        return

    # --- 表格参数 ---
    header = data.columns
    col_widths = [0.15, 0.55, 0.3]  # 列宽比例: 季度, 股票名称, 综合得分
    row_height = 0.08
    header_color = '#E0E6F1'
    row_color_even = '#F7F7F7'
    row_color_odd = 'w'
    font_size_header = 14
    font_size_data = 13
    
    # --- 绘制表头 ---
    y_pos = 0.9
    for i, (col, width) in enumerate(zip(header, col_widths)):
        x_pos = sum(col_widths[:i])
        # 绘制表头背景
        ax.add_patch(patches.Rectangle((x_pos, y_pos - row_height), width, row_height, facecolor=header_color, edgecolor='none'))
        # 打印表头文字
        ax.text(x_pos + width / 2, y_pos - row_height / 2, col, ha='center', va='center', fontsize=font_size_header, weight='bold')

    # --- 绘制数据行 ---
    y_pos -= row_height
    for row_idx, row_data in enumerate(data.itertuples(index=False)):
        bg_color = row_color_even if row_idx % 2 == 0 else row_color_odd
        for i, (cell_data, width) in enumerate(zip(row_data, col_widths)):
            x_pos = sum(col_widths[:i])
            # 绘制数据行背景
            ax.add_patch(patches.Rectangle((x_pos, y_pos - row_height), width, row_height, facecolor=bg_color, edgecolor='none'))
            # 打印数据文字
            ha = 'left' if i == 1 else 'center' # 股票名称左对齐，其他居中
            text_x = x_pos + 0.02 if i == 1 else x_pos + width / 2
            ax.text(text_x, y_pos - row_height / 2, str(cell_data), ha=ha, va='center', fontsize=font_size_data)
        y_pos -= row_height


def create_pixel_perfect_grid(input_path: str, output_path: str, title: str):
    """
    读取数据并创建一个布局完美的2x3网格表格图。
    """
    print(f"\n--- 开始生成像素级完美的网格图表 ---")
    
    if not os.path.exists(input_path):
        print(f"错误: 找不到输入文件 '{input_path}'。")
        return

    df = pd.read_excel(input_path)
    df['年份'] = df['季度'].str.slice(0, 4)
    df['季度号'] = "Q" + df['季度'].str.slice(5) # 格式化为 Q1, Q2...
    df['综合得分'] = df['综合得分'].round(2)
    df_processed = df[['年份', '季度号', '证券名称', '综合得分']].copy()
    df_processed.rename(columns={'季度号': '季度', '证券名称': '股票名称'}, inplace=True)
    
    years = sorted(df_processed['年份'].unique())[:5] # 最多显示5年
    num_years = len(years)

    print(f"正在为 {num_years} 个年份的数据生成图表: {', '.join(years)}")

    # 创建一个更宽的画布来容纳网格
    fig, axes = plt.subplots(2, 3, figsize=(22, 12))
    fig.suptitle(title, fontsize=26, y=0.97)
    flat_axes = axes.flatten()

    for i, year in enumerate(years):
        year_data = df_processed[df_processed['年份'] == year].drop(columns=['年份'])
        draw_custom_table(flat_axes[i], year_data, year)

    # 隐藏多余的子图格子
    for i in range(num_years, len(flat_axes)):
        flat_axes[i].axis('off')
        
    # 调整子图间的间距，防止重叠
    plt.subplots_adjust(wspace=0.4, hspace=0.4)
    
    try:
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        print(f"成功保存最终版图表至: {os.path.join(os.getcwd(), output_path)}")
    except Exception as e:
        print(f"保存图片时出错: {e}")
        
    plt.close(fig)


# ==============================================================================
# --- 4. 主程序入口 ---
# ==============================================================================
if __name__ == "__main__":
    for industry in ['通信设备', '通信服务', '银行', '非银行金融机构']:
        INPUT_EXCEL_NAME = f"output/{industry}_每季度选股策略.xlsx"
        OUTPUT_IMAGE_NAME = f"output/{industry}_年度选股策略图_最终版.png"
        CHART_TITLE = f"{industry}行业：年度选股策略回顾"
        create_pixel_perfect_grid(
            input_path=INPUT_EXCEL_NAME,
            output_path=OUTPUT_IMAGE_NAME,
            title=CHART_TITLE
        )