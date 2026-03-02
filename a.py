import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import matplotlib.dates as mdates
from collections import Counter
import warnings
import os
from matplotlib import font_manager
warnings.filterwarnings('ignore')

# ----------------------
# 核心修复：强制加载并使用SimHei字体
# ----------------------
def setup_chinese_font():
    """配置中文字体，优先加载仓库中的SimHei.ttf"""
    # 1. 优先加载仓库中的SimHei.ttf字体文件
    font_path = os.path.join(os.path.dirname(__file__), 'SimHei.ttf')
    font_name = None
    
    if os.path.exists(font_path):
        try:
            # 加载自定义字体
            custom_font = font_manager.FontProperties(fname=font_path)
            font_name = custom_font.get_name()
            
            # 全局设置字体
            plt.rcParams['font.family'] = font_name
            plt.rcParams['font.sans-serif'] = [font_name]
            plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
            st.success("✅ 成功加载中文字体 SimHei")
            return custom_font
        except Exception as e:
            st.warning(f"⚠️ 加载自定义字体失败: {str(e)}")
    
    # 2. 兜底方案：使用系统字体
    font_options = ['WenQuanYi Zen Hei', 'DejaVu Sans']
    plt.rcParams['font.sans-serif'] = font_options
    plt.rcParams['axes.unicode_minus'] = False
    st.info("ℹ️ 使用系统兜底字体，部分中文可能显示异常")
    
    # 返回空字体对象
    return None

# 执行字体配置并获取字体对象
chinese_font = setup_chinese_font()

# 设置页面配置
st.set_page_config(
    page_title="酒店预约入住数据分析系统",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------
# 节假日配置
# ----------------------
HOLIDAYS = {
    2026: {
        1: [1,2,3,4,5,6,7],  # 元旦
        2: [10,11,12,13,14,15,16],  # 春节
        4: [4,5,6],  # 清明
        5: [1,2,3,4,5],  # 五一
        6: [7,8,9],  # 端午
        9: [13,14,15],  # 中秋
        10: [1,2,3,4,5,6,7]  # 国庆
    }
}

def is_holiday(date):
    """判断日期是否为节假日"""
    try:
        year = date.year
        month = date.month
        day = date.day
        if year in HOLIDAYS and month in HOLIDAYS[year]:
            return day in HOLIDAYS[year][month]
        return False
    except:
        return False

# ----------------------
# 数据加载和预处理
# ----------------------
@st.cache_data
def load_and_preprocess_data(file_path):
    """加载并预处理数据"""
    # 读取数据
    df = pd.read_excel(file_path)
    
    # 转换日期时间格式
    date_columns = ['下单时间', '支付时间', '预约时间', '入住日期', '离店日期']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # 计算入住天数（间晚数）
    if '入住天数' not in df.columns and '离店日期' in df.columns and '入住日期' in df.columns:
        df['入住天数'] = (df['离店日期'] - df['入住日期']).dt.days
    
    # 提取时间特征
    if '下单时间' in df.columns:
        df['下单日期'] = df['下单时间'].dt.date
        df['下单小时'] = df['下单时间'].dt.hour
        df['下单星期'] = df['下单时间'].dt.day_name()
        df['下单星期中文'] = df['下单星期'].map({
            'Monday': '星期一', 'Tuesday': '星期二', 'Wednesday': '星期三',
            'Thursday': '星期四', 'Friday': '星期五', 'Saturday': '星期六', 'Sunday': '星期日'
        })
    
    if '预约时间' in df.columns:
        df['预约日期'] = df['预约时间'].dt.date
        df['预约小时'] = df['预约时间'].dt.hour
        df['预约星期'] = df['预约时间'].dt.day_name()
        df['预约星期中文'] = df['预约星期'].map({
            'Monday': '星期一', 'Tuesday': '星期二', 'Wednesday': '星期三',
            'Thursday': '星期四', 'Friday': '星期五', 'Saturday': '星期六', 'Sunday': '星期日'
        })
    
    if '入住日期' in df.columns:
        df['入住星期'] = df['入住日期'].dt.day_name()
        df['入住星期中文'] = df['入住星期'].map({
            'Monday': '星期一', 'Tuesday': '星期二', 'Wednesday': '星期三',
            'Thursday': '星期四', 'Friday': '星期五', 'Saturday': '星期六', 'Sunday': '星期日'
        })
        df['入住月份'] = df['入住日期'].dt.month
        df['入住日'] = df['入住日期'].dt.day
        # 标记节假日
        df['是否节假日'] = df['入住日期'].apply(is_holiday)
    
    # 计算ADR（平均每日房价）
    if '订单实收' in df.columns and '入住天数' in df.columns:
        df['ADR'] = df['订单实收'] / df['入住天数']
        df['ADR'] = df['ADR'].replace([np.inf, -np.inf], np.nan)
    
    # 计算购买到预约的时间差
    if '下单时间' in df.columns and '预约时间' in df.columns:
        df['购买到预约时间差'] = (df['预约时间'] - df['下单时间']).dt.total_seconds() / 3600  # 小时
        df['购买到预约天数差'] = df['购买到预约时间差'] / 24
    
    # 筛选有效数据
    required_cols = ['预约房型', '入住日期', '预约时间']
    valid_df = df.dropna(subset=[col for col in required_cols if col in df.columns])
    
    return df, valid_df

# ----------------------
# 酒店基础数据配置
# ----------------------
def configure_hotel_rooms():
    """使用表单重写动态配置，避免removeChild报错"""
    st.sidebar.subheader("🏨 酒店房间配置")

    # 总房量
    total_rooms = st.sidebar.number_input(
        "酒店总房间数",
        min_value=1,
        value=100,
        step=1,
        help="输入酒店的总房间数量"
    )

    st.sidebar.markdown("### 各房型房间数配置")
    st.sidebar.info("💡 点击'保存配置'后生效，避免频繁刷新")

    # 初始化房型列表（如果session_state中没有）
    if 'room_types' not in st.session_state:
        st.session_state.room_types = [
            {'name': '花园双床房', 'rooms': 20},
            {'name': '花园大床房', 'rooms': 20},
            {'name': '花园家庭房', 'rooms': 20},
            {'name': '花园豪华双床房', 'rooms': 20},
            {'name': '花园豪华大床房', 'rooms': 20}
        ]

    # 使用表单批量处理房型配置
    with st.sidebar.form("room_config_form", clear_on_submit=False):
        # 显示现有房型
        room_config_temp = []
        for i, room in enumerate(st.session_state.room_types):
            col1, col2 = st.columns([3, 1])
            with col1:
                new_name = st.text_input(f"房型名称 {i+1}", value=room['name'], key=f"room_name_{i}")
            with col2:
                new_count = st.number_input(f"房间数 {i+1}", min_value=1, value=room['rooms'], key=f"room_count_{i}")
            room_config_temp.append({'name': new_name, 'rooms': new_count})
        
        # 操作按钮
        col_ops1, col_ops2 = st.columns(2)
        with col_ops1:
            add_btn = st.form_submit_button("➕ 添加新房型")
        with col_ops2:
            del_btn = st.form_submit_button("🗑️ 删除最后一个房型")
        
        # 保存配置按钮
        save_btn = st.form_submit_button("💾 保存所有配置")

    # 处理添加/删除操作（单次rerun，避免循环）
    if add_btn:
        st.session_state.room_types.append({'name': f'新房型{len(st.session_state.room_types)+1}', 'rooms': 10})
        st.rerun()
    
    if del_btn and len(st.session_state.room_types) > 1:
        st.session_state.room_types.pop()
        st.rerun()
    
    # 保存最终配置
    if save_btn:
        st.session_state.room_types = room_config_temp
    
    # 构建房型配置字典
    room_config = {room['name']: room['rooms'] for room in st.session_state.room_types}
    
    # 保存到session state
    st.session_state['hotel_config'] = {
        'total_rooms': total_rooms,
        'room_config': room_config
    }
    
    return total_rooms, room_config

# ----------------------
# 订单状态筛选器
# ----------------------
def add_order_status_filter(df):
    """添加订单状态筛选器"""
    if '预约状态' in df.columns:
        status_options = df['预约状态'].unique().tolist()
        selected_status = st.sidebar.multiselect(
            "筛选订单状态",
            status_options,
            default=status_options,
            help="选择要分析的订单状态"
        )
        return selected_status
    return None

# ----------------------
# 核心分析函数
# ----------------------
def calculate_occ(room_nights, total_rooms, date_range_days):
    """计算入住率OCC"""
    if total_rooms == 0 or date_range_days == 0:
        return 0
    occ = (room_nights / (total_rooms * date_range_days)) * 100
    return round(occ, 2)

def analyze_room_type_performance(df, room_config):
    """房型深度分析（适配动态房型）"""
    if '预约房型' not in df.columns or '入住日期' not in df.columns:
        return None
    
    # 基础统计
    room_analysis = df.groupby('预约房型').agg({
        '入住天数': ['sum', 'mean', 'count'],  # 总间晚数、平均间晚数、订单数
        '订单实收': ['sum', 'mean'] if '订单实收' in df.columns else ['count'],  # 总收入、平均订单金额
        'ADR': ['mean', 'median'] if 'ADR' in df.columns else ['count'],  # 平均ADR、中位ADR
        '入住日期': ['min', 'max']  # 最早/最晚入住日期
    }).round(2)
    
    # 统一列名
    cols = []
    if '入住天数' in room_analysis.columns.levels[0]:
        cols.extend(['总间晚数', '平均间晚数', '订单数'])
    if '订单实收' in room_analysis.columns.levels[0]:
        cols.extend(['总收入', '平均订单金额'])
    if 'ADR' in room_analysis.columns.levels[0]:
        cols.extend(['平均ADR', '中位ADR'])
    cols.extend(['最早入住日期', '最晚入住日期'])
    room_analysis.columns = cols
    
    # 计算日期范围
    date_start = df['入住日期'].min()
    date_end = df['入住日期'].max()
    date_range_days = max(1, (date_end - date_start).days + 1)
    
    # 合并动态配置的房型
    all_rooms = list(room_config.keys())
    for room in all_rooms:
        if room not in room_analysis.index:
            empty_row = [0]*len(room_analysis.columns)
            empty_row[-2:] = [pd.NaT, pd.NaT]  # 日期列设为NaT
            room_analysis.loc[room] = empty_row
    
    # 添加房型配置和OCC
    room_analysis['房型房量'] = room_analysis.index.map(lambda x: room_config.get(x, 0))
    room_analysis['房型OCC(%)'] = room_analysis.apply(
        lambda row: calculate_occ(row['总间晚数'], row['房型房量'], date_range_days),
        axis=1
    )
    
    # 计算占比
    total_room_nights = room_analysis['总间晚数'].sum()
    total_revenue = room_analysis['总收入'].sum() if '总收入' in room_analysis.columns else 1
    if total_room_nights > 0:
        room_analysis['间晚占比(%)'] = (room_analysis['总间晚数'] / total_room_nights * 100).round(2)
    else:
        room_analysis['间晚占比(%)'] = 0
    
    if total_revenue > 0 and '总收入' in room_analysis.columns:
        room_analysis['收入占比(%)'] = (room_analysis['总收入'] / total_revenue * 100).round(2)
    else:
        room_analysis['收入占比(%)'] = 0
    
    # 格式化日期
    for col in ['最早入住日期', '最晚入住日期']:
        room_analysis[col] = room_analysis[col].dt.strftime('%Y-%m-%d')
    
    return room_analysis

def plot_booking_trend(df):
    """预订趋势分析图表 - 强制指定字体"""
    if '预约时间' not in df.columns:
        return None
    
    # 按日期统计预订量
    trend_data = df.groupby(df['预约时间'].dt.date).size()
    
    # 创建图表
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.plot(trend_data.index, trend_data.values, 'o-', color='#FF6B6B', linewidth=2, markersize=4)
    
    # 样式设置 - 强制指定中文字体
    title = '预约量趋势（按日期）'
    xlabel = '日期'
    ylabel = '预约订单数'
    
    ax.set_title(title, fontproperties=chinese_font, fontsize=14, fontweight='bold', pad=20)
    ax.set_xlabel(xlabel, fontproperties=chinese_font, fontsize=12)
    ax.set_ylabel(ylabel, fontproperties=chinese_font, fontsize=12)
    ax.grid(True, alpha=0.3)
    
    # 日期格式化
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=2))
    plt.xticks(rotation=45)
    
    # 为刻度标签设置字体
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    return fig

def plot_room_night_analysis(df):
    """入住间夜分析 - 强制指定字体"""
    if '入住日期' not in df.columns or '入住天数' not in df.columns:
        return None
    
    # 按日期统计间夜数
    night_data = df.groupby(df['入住日期'].dt.date)['入住天数'].sum()
    
    # 创建图表
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.bar(night_data.index, night_data.values, color='#4ECDC4', alpha=0.8, width=0.8)
    
    # 样式设置 - 强制指定中文字体
    title = '每日入住间夜数'
    xlabel = '日期'
    ylabel = '间夜数'
    
    ax.set_title(title, fontproperties=chinese_font, fontsize=14, fontweight='bold', pad=20)
    ax.set_xlabel(xlabel, fontproperties=chinese_font, fontsize=12)
    ax.set_ylabel(ylabel, fontproperties=chinese_font, fontsize=12)
    ax.grid(True, alpha=0.3, axis='y')
    
    # 日期格式化
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=2))
    plt.xticks(rotation=45)
    
    # 为刻度标签设置字体
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    return fig

def plot_checkin_time_analysis(df):
    """入住时间分析 - 强制指定字体"""
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(14, 10))
    
    # 1. 按星期分析
    if '入住星期中文' in df.columns:
        week_data = df['入住星期中文'].value_counts()
        week_order = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
        week_data = week_data.reindex(week_order, fill_value=0)
        ax1.bar(week_data.index, week_data.values, color='#45B7D1', alpha=0.8)
        ax1.set_title('入住星期分布', fontproperties=chinese_font, fontweight='bold')
        ax1.tick_params(axis='x', rotation=45)
        # 设置刻度字体
        for label in ax1.get_xticklabels() + ax1.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    # 2. 按月分析
    if '入住月份' in df.columns:
        month_data = df['入住月份'].value_counts().sort_index()
        ax2.bar(month_data.index, month_data.values, color='#96CEB4', alpha=0.8)
        ax2.set_title('入住月份分布', fontproperties=chinese_font, fontweight='bold')
        ax2.set_xlabel('月份', fontproperties=chinese_font)
        ax2.set_xticks(range(1, 13))
        # 设置刻度字体
        for label in ax2.get_xticklabels() + ax2.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    # 3. 节假日vs工作日
    if '是否节假日' in df.columns:
        holiday_data = df['是否节假日'].value_counts()
        labels = ['工作日', '节假日']
        values = [holiday_data.get(False, 0), holiday_data.get(True, 0)]
        ax3.pie(values, labels=labels, autopct='%1.1f%%', colors=['#FFEAA7', '#FD79A8'], startangle=90)
        ax3.set_title('节假日vs工作日入住占比', fontproperties=chinese_font, fontweight='bold')
        # 设置标签字体
        for text in ax3.texts:
            text.set_fontproperties(chinese_font)
    
    # 4. 按小时分析（预约时间）
    if '预约小时' in df.columns:
        hour_data = df['预约小时'].value_counts().sort_index()
        ax4.plot(hour_data.index, hour_data.values, 'o-', color='#6C5CE7', linewidth=2)
        ax4.set_title('预约小时分布', fontproperties=chinese_font, fontweight='bold')
        ax4.set_xlabel('小时', fontproperties=chinese_font)
        ax4.set_xticks(range(0, 24, 2))
        # 设置刻度字体
        for label in ax4.get_xticklabels() + ax4.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    return fig

def plot_purchase_booking_gap(df):
    """购买预约时间差分析 - 强制指定字体"""
    if '购买到预约时间差' not in df.columns:
        return None
    
    # 过滤有效数据
    gap_data = df['购买到预约时间差'].dropna()
    gap_data = gap_data[gap_data >= 0]  # 只保留正数
    
    if len(gap_data) == 0:
        return None
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    
    # 1. 直方图
    ax1.hist(gap_data, bins=20, color='#A29BFE', alpha=0.7, edgecolor='black')
    ax1.set_title('购买到预约时间差分布（小时）', fontproperties=chinese_font, fontweight='bold')
    ax1.set_xlabel('小时', fontproperties=chinese_font, fontsize=12)
    ax1.set_ylabel('订单数', fontproperties=chinese_font, fontsize=12)
    ax1.grid(True, alpha=0.3)
    # 设置刻度字体
    for label in ax1.get_xticklabels() + ax1.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    # 2. 按天数分组
    gap_days = gap_data / 24
    bins = [0, 1, 3, 7, 15, 30, 90, float('inf')]
    labels = ['0-1天', '1-3天', '3-7天', '7-15天', '15-30天', '30-90天', '90天以上']
    gap_groups = pd.cut(gap_days, bins=bins, labels=labels, right=False)
    group_counts = gap_groups.value_counts()
    
    ax2.bar(group_counts.index, group_counts.values, color='#FD79A8', alpha=0.8)
    ax2.set_title('购买到预约时间差分组统计', fontproperties=chinese_font, fontweight='bold')
    ax2.tick_params(axis='x', rotation=45)
    # 设置刻度字体
    for label in ax2.get_xticklabels() + ax2.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    return fig

def plot_occ_analysis(df, total_rooms):
    """OCC分析 - 强制指定字体"""
    if '入住日期' not in df.columns or '入住天数' not in df.columns:
        return None
    
    # 按日期计算OCC
    date_start = df['入住日期'].min()
    date_end = df['入住日期'].max()
    date_range = pd.date_range(start=date_start, end=date_end)
    
    occ_data = []
    for date in date_range:
        daily_df = df[df['入住日期'].dt.date == date.date()]
        daily_nights = daily_df['入住天数'].sum()
        daily_occ = calculate_occ(daily_nights, total_rooms, 1)
        occ_data.append(daily_occ)
    
    # 创建图表
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))
    
    # 1. 每日OCC趋势
    ax1.plot(date_range, occ_data, 'o-', color='#E17055', linewidth=2, markersize=4)
    ax1.set_title('每日OCC趋势', fontproperties=chinese_font, fontsize=14, fontweight='bold')
    ax1.set_ylabel('OCC (%)', fontproperties=chinese_font, fontsize=12)
    ax1.set_ylim(0, 100)
    ax1.grid(True, alpha=0.3)
    
    # 添加平均线
    avg_occ = np.mean(occ_data)
    ax1.axhline(y=avg_occ, color='red', linestyle='--', alpha=0.7, 
                label=f'平均OCC: {avg_occ:.1f}%')
    ax1.legend(prop=chinese_font)
    
    # 设置刻度字体
    for label in ax1.get_xticklabels() + ax1.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    # 2. OCC分布
    ax2.hist(occ_data, bins=15, color='#00B894', alpha=0.7, edgecolor='black')
    ax2.set_title('OCC值分布', fontproperties=chinese_font, fontsize=14, fontweight='bold')
    ax2.set_xlabel('OCC (%)', fontproperties=chinese_font, fontsize=12)
    ax2.set_ylabel('天数', fontproperties=chinese_font, fontsize=12)
    ax2.grid(True, alpha=0.3)
    
    # 设置刻度字体
    for label in ax2.get_xticklabels() + ax2.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    return fig

def analyze_packages(df):
    """按商品ID分析套餐"""
    if '商品ID' not in df.columns:
        return None, None
    
    # 按商品ID分组
    agg_dict = {
        '商品名称': 'first',  # 套餐名称
        '订单ID': 'count',    # 订单数
        '用户实付金额': ['sum', 'mean', 'median'] if '用户实付金额' in df.columns else ['count'],  # 收入和价格
        '入住天数': ['sum', 'mean'] if '入住天数' in df.columns else ['count'],  # 间晚数
        '预约时间': ['min', 'max']  # 时间范围
    }
    
    package_analysis = df.groupby('商品ID').agg(agg_dict).round(2)
    
    # 整理列名
    cols = []
    cols.extend(['套餐名称', '订单数'])
    if '用户实付金额' in package_analysis.columns.levels[0]:
        cols.extend(['总收入', '平均价格', '中位价格'])
    if '入住天数' in package_analysis.columns.levels[0]:
        cols.extend(['总间晚数', '平均间晚数'])
    cols.extend(['最早预约时间', '最晚预约时间'])
    package_analysis.columns = cols
    
    # 计算占比
    total_orders = package_analysis['订单数'].sum()
    total_revenue = package_analysis['总收入'].sum() if '总收入' in package_analysis.columns else 1
    
    if total_orders > 0:
        package_analysis['订单占比(%)'] = (package_analysis['订单数'] / total_orders * 100).round(2)
    else:
        package_analysis['订单占比(%)'] = 0
    
    if total_revenue > 0 and '总收入' in package_analysis.columns:
        package_analysis['收入占比(%)'] = (package_analysis['总收入'] / total_revenue * 100).round(2)
    else:
        package_analysis['收入占比(%)'] = 0
    
    # 格式化时间
    for col in ['最早预约时间', '最晚预约时间']:
        package_analysis[col] = package_analysis[col].dt.strftime('%Y-%m-%d %H:%M')
    
    # 套餐趋势分析
    package_trend = None
    if '预约时间' in df.columns:
        package_trend = df.groupby([df['预约时间'].dt.date, '商品ID']).size().unstack(fill_value=0)
    
    return package_analysis, package_trend

def plot_package_analysis(package_analysis, package_trend):
    """绘制套餐分析图表 - 强制指定字体"""
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
    
    # 1. 套餐订单数和收入占比
    x = range(len(package_analysis))
    width = 0.35
    
    bars1 = ax1.bar([i - width/2 for i in x], package_analysis['订单数'], width, 
                   label='订单数', color='#FF6B6B', alpha=0.8)
    ax1_twin = ax1.twinx()
    line1 = ax1_twin.plot(x, package_analysis['收入占比(%)'], 'o-', color='#4ECDC4', linewidth=2, 
                          label='收入占比(%)')
    
    ax1.set_title('各套餐订单数与收入占比', fontproperties=chinese_font, fontweight='bold', fontsize=12)
    ax1.set_xlabel('套餐ID', fontproperties=chinese_font, fontsize=12)
    ax1.set_ylabel('订单数', fontproperties=chinese_font, color='#FF6B6B', fontsize=12)
    ax1_twin.set_ylabel('收入占比 (%)', fontproperties=chinese_font, color='#4ECDC4', fontsize=12)
    ax1.set_xticks(x)
    ax1.set_xticklabels([f"套餐{pid}" for pid in package_analysis.index], rotation=45, ha='right')
    ax1.grid(True, alpha=0.3)
    
    # 设置字体
    ax1.legend(prop=chinese_font)
    ax1_twin.legend(prop=chinese_font)
    for label in ax1.get_xticklabels() + ax1.get_yticklabels() + ax1_twin.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    # 2. 套餐平均价格和间晚数
    if '平均价格' in package_analysis.columns and '总间晚数' in package_analysis.columns:
        bars2 = ax2.bar([i - width/2 for i in x], package_analysis['平均价格'], width, 
                       label='平均价格', color='#45B7D1', alpha=0.8)
        ax2_twin = ax2.twinx()
        line2 = ax2_twin.plot(x, package_analysis['总间晚数'], 's-', color='#96CEB4', linewidth=2, 
                              label='总间晚数')
        
        ax2.set_title('各套餐平均价格与总间晚数', fontproperties=chinese_font, fontweight='bold', fontsize=12)
        ax2.set_xlabel('套餐ID', fontproperties=chinese_font, fontsize=12)
        ax2.set_ylabel('平均价格 (¥)', fontproperties=chinese_font, color='#45B7D1', fontsize=12)
        ax2_twin.set_ylabel('总间晚数', fontproperties=chinese_font, color='#96CEB4', fontsize=12)
        ax2.set_xticks(x)
        ax2.set_xticklabels([f"套餐{pid}" for pid in package_analysis.index], rotation=45, ha='right')
        ax2.grid(True, alpha=0.3)
        
        # 设置字体
        ax2.legend(prop=chinese_font)
        ax2_twin.legend(prop=chinese_font)
        for label in ax2.get_xticklabels() + ax2.get_yticklabels() + ax2_twin.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    # 3. 套餐收入占比饼图
    colors = plt.cm.Set3(np.linspace(0, 1, len(package_analysis)))
    wedges, texts, autotexts = ax3.pie(package_analysis['收入占比(%)'], 
                                       labels=[f"套餐{pid}" for pid in package_analysis.index], 
                                       autopct='%1.1f%%', colors=colors, startangle=90)
    ax3.set_title('各套餐收入占比', fontproperties=chinese_font, fontweight='bold', fontsize=12)
    
    # 设置饼图字体
    for text in texts + autotexts:
        text.set_fontproperties(chinese_font)
    
    # 4. 套餐预订趋势
    if package_trend is not None:
        for i, col in enumerate(package_trend.columns):
            ax4.plot(package_trend.index, package_trend[col], 'o-', linewidth=2, 
                    label=f"套餐{col}", color=colors[i], alpha=0.8)
        ax4.set_title('套餐预订趋势', fontproperties=chinese_font, fontweight='bold', fontsize=12)
        ax4.set_xlabel('日期', fontproperties=chinese_font, fontsize=12)
        ax4.set_ylabel('订单数', fontproperties=chinese_font, fontsize=12)
        ax4.legend(prop=chinese_font, bbox_to_anchor=(1.05, 1), loc='upper left')
        ax4.grid(True, alpha=0.3)
        ax4.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
        ax4.xaxis.set_major_locator(mdates.DayLocator(interval=2))
        plt.setp(ax4.xaxis.get_majorticklabels(), rotation=45)
        
        # 设置字体
        for label in ax4.get_xticklabels() + ax4.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    return fig

# ----------------------
# 新增：预约-入住时差分析函数
# ----------------------
def analyze_booking_checkin_gap(df):
    """分析预约时间与入住时间的差"""
    if '预约时间' not in df.columns or '入住日期' not in df.columns:
        return None
    
    # 计算时差（天）
    df['预约入住时差_天'] = (df['入住日期'].dt.date - df['预约时间'].dt.date).dt.days
    df['预约入住时差_天'] = df['预约入住时差_天'].clip(lower=0)  # 排除负数
    
    # 分组统计
    bins = [0, 1, 3, 7, 15, 30, 90, float('inf')]
    labels = ['0-1天', '1-3天', '3-7天', '7-15天', '15-30天', '30-90天', '90天以上']
    df['时差分组'] = pd.cut(df['预约入住时差_天'], bins=bins, labels=labels, right=False)
    
    # 统计结果
    gap_stats = df['时差分组'].value_counts().sort_index()
    gap_summary = {
        '平均时差(天)': round(df['预约入住时差_天'].mean(), 1),
        '中位时差(天)': df['预约入住时差_天'].median(),
        '最集中时段': gap_stats.idxmax(),
        '总订单数': len(df)
    }
    
    return df, gap_stats, gap_summary

def plot_booking_checkin_gap_analysis(df, gap_stats):
    """绘制预约-入住时差分析图表"""
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
    
    # 1. 时差分组分布（柱状图）
    ax1.bar(gap_stats.index, gap_stats.values, color='#6C5CE7', alpha=0.8)
    ax1.set_title('预约-入住时差分组分布', fontproperties=chinese_font, fontsize=14, fontweight='bold')
    ax1.set_xlabel('时差分组', fontproperties=chinese_font, fontsize=12)
    ax1.set_ylabel('订单数', fontproperties=chinese_font, fontsize=12)
    ax1.tick_params(axis='x', rotation=45)
    for label in ax1.get_xticklabels() + ax1.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    # 2. 时差分布直方图
    gap_data = df['预约入住时差_天'].dropna()
    gap_data = gap_data[gap_data >= 0]
    ax2.hist(gap_data, bins=20, color='#FD79A8', alpha=0.7, edgecolor='black')
    ax2.set_title('预约-入住时差分布（天）', fontproperties=chinese_font, fontsize=14, fontweight='bold')
    ax2.set_xlabel('时差（天）', fontproperties=chinese_font, fontsize=12)
    ax2.set_ylabel('订单数', fontproperties=chinese_font, fontsize=12)
    for label in ax2.get_xticklabels() + ax2.get_yticklabels():
        label.set_fontproperties(chinese_font)
    
    # 3. 按星期分析时差
    if '入住星期中文' in df.columns:
        week_gap = df.groupby('入住星期中文')['预约入住时差_天'].mean().reindex(
            ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
        )
        ax3.bar(week_gap.index, week_gap.values, color='#00B894', alpha=0.8)
        ax3.set_title('各星期平均预约-入住时差', fontproperties=chinese_font, fontsize=14, fontweight='bold')
        ax3.set_xlabel('星期', fontproperties=chinese_font, fontsize=12)
        ax3.set_ylabel('平均时差（天）', fontproperties=chinese_font, fontsize=12)
        ax3.tick_params(axis='x', rotation=45)
        for label in ax3.get_xticklabels() + ax3.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    # 4. 按月分析时差
    if '入住月份' in df.columns:
        month_gap = df.groupby('入住月份')['预约入住时差_天'].mean()
        ax4.plot(month_gap.index, month_gap.values, 'o-', color='#E17055', linewidth=2, markersize=6)
        ax4.set_title('各月平均预约-入住时差', fontproperties=chinese_font, fontsize=14, fontweight='bold')
        ax4.set_xlabel('月份', fontproperties=chinese_font, fontsize=12)
        ax4.set_ylabel('平均时差（天）', fontproperties=chinese_font, fontsize=12)
        ax4.set_xticks(range(1, 13))
        for label in ax4.get_xticklabels() + ax4.get_yticklabels():
            label.set_fontproperties(chinese_font)
    
    plt.tight_layout()
    return fig

# ----------------------
# 主页面内容
# ----------------------
def main():
    # 初始化session state
    if 'hotel_config' not in st.session_state:
        st.session_state['hotel_config'] = {
            'total_rooms': 100,
            'room_config': {}
        }
    
    # 侧边栏
    st.sidebar.title("🏨 酒店预约数据分析系统")
    st.sidebar.markdown("### 数据文件上传")
    
    # 文件上传
    uploaded_file = st.sidebar.file_uploader(
        "上传预约入住明细Excel文件", 
        type=['xlsx', 'xls'],
        help="请上传包含预约入住明细的Excel文件"
    )
    
    # 酒店配置（修复版）
    total_rooms, room_config = configure_hotel_rooms()
    
    if uploaded_file is None:
        st.warning("请在左侧边栏上传Excel数据文件，并配置酒店房间信息")
        st.stop()
    
    # 加载数据
    with st.spinner("正在加载和预处理数据..."):
        df, valid_df = load_and_preprocess_data(uploaded_file)
        st.session_state['valid_df'] = valid_df
    
    # 订单状态筛选
    selected_status = add_order_status_filter(valid_df)
    if selected_status is not None and len(selected_status) > 0:
        valid_df = valid_df[valid_df['预约状态'].isin(selected_status)]
    
    # 侧边栏数据概览
    st.sidebar.markdown("### 📊 数据概览")
    st.sidebar.write(f"总记录数: **{len(df):,}**")
    st.sidebar.write(f"有效记录数: **{len(valid_df):,}**")
    
    if len(valid_df) > 0:
        st.sidebar.write(f"数据时间范围:")
        if '预约时间' in valid_df.columns:
            st.sidebar.write(f"- 预约时间: {valid_df['预约时间'].min().strftime('%Y-%m-%d')} 至 {valid_df['预约时间'].max().strftime('%Y-%m-%d')}")
        if '入住日期' in valid_df.columns:
            st.sidebar.write(f"- 入住时间: {valid_df['入住日期'].min().strftime('%Y-%m-%d')} 至 {valid_df['入住日期'].max().strftime('%Y-%m-%d')}")
        
        # 计算整体OCC
        if '入住天数' in valid_df.columns and '入住日期' in valid_df.columns:
            total_room_nights = valid_df['入住天数'].sum()
            date_start = valid_df['入住日期'].min()
            date_end = valid_df['入住日期'].max()
            date_range_days = max(1, (date_end - date_start).days + 1)
            overall_occ = calculate_occ(total_room_nights, total_rooms, date_range_days)
            st.sidebar.metric("整体OCC", f"{overall_occ}%")
    
    # 侧边栏分析选项（新增预约-入住时差分析）
    st.sidebar.markdown("### 📈 分析选项")
    analysis_options = [
        "数据概览", "房型深度分析", "预订趋势分析", "入住间夜分析", 
        "入住时间分析", "购买预约时间差", "OCC分析", "套餐分析", 
        "预约-入住时差分析", "详细数据"
    ]
    analysis_type = st.sidebar.radio(
        "选择分析类型",
        analysis_options
    )
    
    # 主内容区域
    st.title("🏨 酒店预约入住数据分析系统")
    st.markdown("---")
    
    # 数据为空判断
    if len(valid_df) == 0:
        st.error("⚠️ 没有有效数据，请检查上传的文件或筛选条件")
        st.stop()
    
    # ----------------------
    # 各分析模块实现
    # ----------------------
    if analysis_type == "数据概览":
        st.header("📊 数据整体概览")
        
        # 关键指标卡片
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("总预约单数", f"{len(valid_df):,}")
        
        with col2:
            total_room_nights = valid_df['入住天数'].sum() if '入住天数' in valid_df.columns else 0
            st.metric("总间夜数", f"{total_room_nights:,}")
        
        with col3:
            avg_adr = valid_df['ADR'].mean() if 'ADR' in valid_df.columns and not pd.isna(valid_df['ADR']).all() else 0
            st.metric("平均ADR", f"¥{avg_adr:.2f}")
        
        with col4:
            total_revenue = valid_df['订单实收'].sum() if '订单实收' in valid_df.columns else 0
            st.metric("总实收金额", f"¥{total_revenue:,.2f}")
        
        with col5:
            if '入住天数' in valid_df.columns and '入住日期' in valid_df.columns:
                date_start = valid_df['入住日期'].min()
                date_end = valid_df['入住日期'].max()
                date_range_days = max(1, (date_end - date_start).days + 1)
                overall_occ = calculate_occ(total_room_nights, total_rooms, date_range_days)
                st.metric("整体OCC", f"{overall_occ}%")
            else:
                st.metric("整体OCC", "N/A")
        
        st.markdown("---")
        
        # 核心统计信息
        st.subheader("核心运营指标")
        stats_data = {
            "指标": [
                "预约订单数", "总间夜数", "平均间晚数", "平均ADR",
                "总实收金额", "整体OCC", "主要房型", "入住日期范围",
                "节假日订单占比", "平均购买-预约时间差"
            ],
            "数值": [
                f"{len(valid_df):,}",
                f"{valid_df['入住天数'].sum():,}" if '入住天数' in valid_df.columns else "N/A",
                f"{valid_df['入住天数'].mean():.1f} 天/单" if '入住天数' in valid_df.columns else "N/A",
                f"¥{valid_df['ADR'].mean():.2f}" if 'ADR' in valid_df.columns and not pd.isna(valid_df['ADR']).all() else "N/A",
                f"¥{valid_df['订单实收'].sum():,.2f}" if '订单实收' in valid_df.columns else "N/A",
                f"{overall_occ}%" if '入住天数' in valid_df.columns and '入住日期' in valid_df.columns else "N/A",
                f"{valid_df['预约房型'].value_counts().index[0]}" if '预约房型' in valid_df.columns else "N/A",
                f"{valid_df['入住日期'].min().strftime('%Y-%m-%d')} 至 {valid_df['入住日期'].max().strftime('%Y-%m-%d')}" if '入住日期' in valid_df.columns else "N/A",
                f"{valid_df['是否节假日'].sum() / len(valid_df) * 100:.1f}%" if '是否节假日' in valid_df.columns else "N/A",
                f"{valid_df['购买到预约时间差'].mean():.1f} 小时" if '购买到预约时间差' in valid_df.columns else "N/A"
            ]
        }
        stats_df = pd.DataFrame(stats_data)
        st.dataframe(stats_df, use_container_width=True)
    
    elif analysis_type == "房型深度分析":
        st.header("🏠 房型深度分析")
        
        # 执行分析
        room_analysis = analyze_room_type_performance(valid_df, room_config)
        
        if room_analysis is None:
            st.warning("⚠️ 缺少房型分析所需字段（预约房型/入住日期）")
        else:
            # 显示数据表格
            st.subheader("1. 房型性能指标汇总")
            st.dataframe(room_analysis, use_container_width=True)
            
            # 可视化
            st.subheader("2. 房型分析可视化")
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
            
            # 房型订单数对比
            room_orders = room_analysis['订单数'].sort_values(ascending=False)
            ax1.bar(room_orders.index, room_orders.values, color='#FF6B6B', alpha=0.8)
            ax1.set_title('各房型订单数', fontproperties=chinese_font, fontweight='bold')
            ax1.set_xlabel('房型', fontproperties=chinese_font)
            ax1.set_ylabel('订单数', fontproperties=chinese_font)
            ax1.tick_params(axis='x', rotation=45)
            
            # 设置刻度字体
            for label in ax1.get_xticklabels() + ax1.get_yticklabels():
                label.set_fontproperties(chinese_font)
            
            # 房型OCC对比
            room_occ = room_analysis['房型OCC(%)'].sort_values(ascending=False)
            ax2.bar(room_occ.index, room_occ.values, color='#4ECDC4', alpha=0.8)
            ax2.set_title('各房型OCC', fontproperties=chinese_font, fontweight='bold')
            ax2.set_xlabel('房型', fontproperties=chinese_font)
            ax2.set_ylabel('OCC (%)', fontproperties=chinese_font)
            ax2.tick_params(axis='x', rotation=45)
            
            # 设置刻度字体
            for label in ax2.get_xticklabels() + ax2.get_yticklabels():
                label.set_fontproperties(chinese_font)
            
            plt.tight_layout()
            st.pyplot(fig)
    
    elif analysis_type == "预订趋势分析":
        st.header("📈 预订趋势分析")
        
        fig = plot_booking_trend(valid_df)
        if fig is None:
            st.warning("⚠️ 缺少预约时间字段，无法生成趋势图")
        else:
            st.pyplot(fig)
            
            # 额外统计
            st.subheader("预订趋势关键指标")
            if '预约时间' in valid_df.columns:
                daily_avg = valid_df.groupby(valid_df['预约时间'].dt.date).size().mean()
                peak_day = valid_df.groupby(valid_df['预约时间'].dt.date).size().idxmax()
                peak_value = valid_df.groupby(valid_df['预约时间'].dt.date).size().max()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("日均预约量", f"{daily_avg:.1f}")
                with col2:
                    st.metric("预约峰值日期", peak_day.strftime('%Y-%m-%d'))
                with col3:
                    st.metric("峰值预约量", f"{peak_value:,}")
    
    elif analysis_type == "入住间夜分析":
        st.header("🌙 入住间夜分析")
        
        fig = plot_room_night_analysis(valid_df)
        if fig is None:
            st.warning("⚠️ 缺少入住日期/入住天数字段，无法生成间夜分析图")
        else:
            st.pyplot(fig)
            
            # 额外统计
            st.subheader("间夜分析关键指标")
            if '入住天数' in valid_df.columns:
                avg_nights = valid_df['入住天数'].mean()
                total_nights = valid_df['入住天数'].sum()
                max_single_night = valid_df['入住天数'].max()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("平均每单间晚数", f"{avg_nights:.1f}")
                with col2:
                    st.metric("总间晚数", f"{total_nights:,}")
                with col3:
                    st.metric("最大单次间夜数", f"{max_single_night}")
    
    elif analysis_type == "入住时间分析":
        st.header("⏰ 入住时间分析")
        
        fig = plot_checkin_time_analysis(valid_df)
        st.pyplot(fig)
    
    elif analysis_type == "购买预约时间差":
        st.header("🕒 购买-预约时间差分析")
        
        fig = plot_purchase_booking_gap(valid_df)
        if fig is None:
            st.warning("⚠️ 缺少下单时间/预约时间字段，无法生成时间差分析图")
        else:
            st.pyplot(fig)
            
            # 关键指标
            if '购买到预约时间差' in valid_df.columns:
                gap_data = valid_df['购买到预约时间差'].dropna()
                if len(gap_data) > 0:
                    avg_gap = gap_data.mean()
                    median_gap = gap_data.median()
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("平均购买-预约时间差", f"{avg_gap:.1f} 小时")
                    with col2:
                        st.metric("中位购买-预约时间差", f"{median_gap:.1f} 小时")
    
    elif analysis_type == "OCC分析":
        st.header("📊 OCC（入住率）分析")
        
        fig = plot_occ_analysis(valid_df, total_rooms)
        if fig is None:
            st.warning("⚠️ 缺少入住日期/入住天数字段，无法生成OCC分析图")
        else:
            st.pyplot(fig)
            
            # OCC关键指标
            if '入住天数' in valid_df.columns and '入住日期' in valid_df.columns:
                total_room_nights = valid_df['入住天数'].sum()
                date_start = valid_df['入住日期'].min()
                date_end = valid_df['入住日期'].max()
                date_range_days = max(1, (date_end - date_start).days + 1)
                overall_occ = calculate_occ(total_room_nights, total_rooms, date_range_days)
                
                # 按星期计算OCC
                week_occ = {}
                week_names = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
                for week in week_names:
                    week_df = valid_df[valid_df['入住星期中文'] == week]
                    week_nights = week_df['入住天数'].sum()
                    week_days = len(week_df['入住日期'].dt.date.unique())
                    week_occ[week] = calculate_occ(week_nights, total_rooms, week_days)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("整体OCC", f"{overall_occ:.1f}%")
                with col2:
                    max_week = max(week_occ.items(), key=lambda x: x[1])
                    st.metric("OCC最高的星期", f"{max_week[0]} ({max_week[1]:.1f}%)")
    
    elif analysis_type == "套餐分析":
        st.header("📦 套餐分析（按商品ID）")
        
        if '商品ID' not in valid_df.columns:
            st.warning("⚠️ 数据中缺少商品ID字段，无法进行套餐分析")
        else:
            # 套餐筛选
            package_ids = valid_df['商品ID'].unique()
            selected_packages = st.multiselect(
                "选择套餐（按商品ID）",
                package_ids,
                default=package_ids,
                help="选择要分析的套餐"
            )
            
            filtered_df = valid_df[valid_df['商品ID'].isin(selected_packages)]
            
            if len(filtered_df) == 0:
                st.warning("⚠️ 没有符合筛选条件的套餐数据")
            else:
                # 执行分析
                package_analysis, package_trend = analyze_packages(filtered_df)
                
                if package_analysis is None:
                    st.warning("⚠️ 套餐分析失败，请检查数据格式")
                else:
                    # 显示分析结果
                    st.subheader("1. 套餐性能指标汇总")
                    st.dataframe(package_analysis, use_container_width=True)
                    
                    # 可视化
                    st.subheader("2. 套餐分析可视化")
                    fig = plot_package_analysis(package_analysis, package_trend)
                    st.pyplot(fig)
                    
                    # 关键洞察
                    st.subheader("3. 关键洞察")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        top_package = package_analysis['订单数'].idxmax()
                        st.info(f"""
                        **最受欢迎套餐**  
                        套餐ID: {top_package}  
                        套餐名称: {package_analysis.loc[top_package, '套餐名称']}  
                        订单数: {package_analysis.loc[top_package, '订单数']:,}  
                        订单占比: {package_analysis.loc[top_package, '订单占比(%)']}%
                        """)
                    
                    with col2:
                        if '总收入' in package_analysis.columns:
                            highest_rev_package = package_analysis['总收入'].idxmax()
                            st.info(f"""
                            **收入最高套餐**  
                            套餐ID: {highest_rev_package}  
                            套餐名称: {package_analysis.loc[highest_rev_package, '套餐名称']}  
                            总收入: ¥{package_analysis.loc[highest_rev_package, '总收入']:,.2f}  
                            收入占比: {package_analysis.loc[highest_rev_package, '收入占比(%)']}%
                            """)
                    
                    with col3:
                        if '平均价格' in package_analysis.columns:
                            highest_price_package = package_analysis['平均价格'].idxmax()
                            st.info(f"""
                            **单价最高套餐**  
                            套餐ID: {highest_price_package}  
                            套餐名称: {package_analysis.loc[highest_price_package, '套餐名称']}  
                            平均价格: ¥{package_analysis.loc[highest_price_package, '平均价格']:.2f}  
                            平均间晚数: {package_analysis.loc[highest_price_package, '平均间晚数']:.1f}
                            """)
    
    # ----------------------
    # 新增：预约-入住时差分析模块
    # ----------------------
    elif analysis_type == "预约-入住时差分析":
        st.header("⏱️ 预约-入住时差分析")
        
        if '预约时间' not in valid_df.columns or '入住日期' not in valid_df.columns:
            st.warning("⚠️ 缺少预约时间或入住日期字段，无法进行时差分析")
        else:
            # 执行分析
            df_with_gap, gap_stats, gap_summary = analyze_booking_checkin_gap(valid_df)
            
            # 显示关键指标
            st.subheader("1. 时差核心指标")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("平均时差(天)", gap_summary['平均时差(天)'])
            with col2:
                st.metric("中位时差(天)", gap_summary['中位时差(天)'])
            with col3:
                st.metric("最集中时段", gap_summary['最集中时段'])
            with col4:
                st.metric("总订单数", gap_summary['总订单数'])
            
            # 可视化
            st.subheader("2. 时差分析可视化")
            fig = plot_booking_checkin_gap_analysis(df_with_gap, gap_stats)
            st.pyplot(fig)
            
            # 分组详情
            st.subheader("3. 时差分组详情")
            gap_detail = pd.DataFrame({
                '时差分组': gap_stats.index,
                '订单数': gap_stats.values,
                '占比(%)': (gap_stats.values / gap_stats.sum() * 100).round(1)
            })
            st.dataframe(gap_detail, use_container_width=True)
    
    elif analysis_type == "详细数据":
        st.header("📋 详细数据查看")
        
        # 数据筛选
        st.subheader("数据筛选")
        col1, col2 = st.columns(2)
        with col1:
            date_start = st.date_input("开始日期", 
                                      value=valid_df['入住日期'].min().date() if '入住日期' in valid_df.columns else datetime.now().date())
        with col2:
            date_end = st.date_input("结束日期", 
                                      value=valid_df['入住日期'].max().date() if '入住日期' in valid_df.columns else datetime.now().date())
        
        # 筛选数据
        filtered_data = valid_df.copy()
        if '入住日期' in filtered_data.columns:
            filtered_data = filtered_data[(filtered_data['入住日期'].dt.date >= date_start) & 
                                         (filtered_data['入住日期'].dt.date <= date_end)]
        
        # 显示数据
        st.subheader(f"筛选后数据（共 {len(filtered_data):,} 条）")
        st.dataframe(filtered_data, use_container_width=True)
        
        # 数据导出
        @st.cache_data
        def convert_df_to_csv(df):
            return df.to_csv(index=False, encoding='utf-8-sig')
        
        csv_data = convert_df_to_csv(filtered_data)
        st.download_button(
            label="📥 导出筛选后数据为CSV",
            data=csv_data,
            file_name=f"酒店数据_{date_start}_{date_end}.csv",
            mime="text/csv"
        )

# ----------------------
# 运行应用
# ----------------------
if __name__ == "__main__":
    main()