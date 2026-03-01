import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import matplotlib.dates as mdates
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 设置页面配置
st.set_page_config(
    page_title="酒店预约入住数据分析系统",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------
# 节假日配置（可根据需要扩展）
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
# 1. 数据加载和预处理
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
    valid_df = df.dropna(subset=['预约房型', '入住日期', '预约时间'])
    
    return df, valid_df

# ----------------------
# 2. 酒店基础数据配置
# ----------------------
def configure_hotel_rooms():
    """配置酒店房间数据"""
    st.sidebar.subheader("🏨 酒店房间配置")
    
    # 总房量
    total_rooms = st.sidebar.number_input(
        "酒店总房间数",
        min_value=1,
        value=100,
        step=1,
        help="输入酒店的总房间数量"
    )
    
    # 房型配置
    st.sidebar.markdown("### 各房型房间数配置")
    room_config = {}
    
    # 先从数据中获取房型列表（如果已上传数据）
    if 'valid_df' in st.session_state and len(st.session_state.valid_df) > 0:
        room_types = st.session_state.valid_df['预约房型'].unique()
        for room_type in room_types:
            room_config[room_type] = st.sidebar.number_input(
                f"{room_type}房间数",
                min_value=1,
                value=20,
                step=1,
                key=f"room_{room_type}"
            )
    else:
        # 默认房型配置
        default_rooms = ['花园双床房', '花园大床房', '花园家庭房', '花园豪华双床房', '花园豪华大床房']
        for room_type in default_rooms:
            room_config[room_type] = st.sidebar.number_input(
                f"{room_type}房间数",
                min_value=1,
                value=20,
                step=1,
                key=f"room_{room_type}"
            )
    
    # 保存配置到session state
    st.session_state['hotel_config'] = {
        'total_rooms': total_rooms,
        'room_config': room_config
    }
    
    return total_rooms, room_config

# ----------------------
# 3. 核心分析函数
# ----------------------
def calculate_occ(room_nights, total_rooms, date_range_days):
    """计算入住率OCC"""
    if total_rooms == 0 or date_range_days == 0:
        return 0
    occ = (room_nights / (total_rooms * date_range_days)) * 100
    return round(occ, 2)

def analyze_room_type_performance(df, room_config):
    """房型深度分析"""
    # 基础统计
    room_analysis = df.groupby('预约房型').agg({
        '入住天数': ['sum', 'mean', 'count'],  # 总间晚数、平均间晚数、订单数
        '订单实收': ['sum', 'mean'],  # 总收入、平均订单金额
        'ADR': ['mean', 'median'],  # 平均ADR、中位ADR
        '入住日期': ['min', 'max']  # 最早/最晚入住日期
    }).round(2)
    
    room_analysis.columns = ['总间晚数', '平均间晚数', '订单数', '总收入', '平均订单金额', 
                           '平均ADR', '中位ADR', '最早入住日期', '最晚入住日期']
    
    # 计算OCC
    date_start = df['入住日期'].min()
    date_end = df['入住日期'].max()
    date_range_days = (date_end - date_start).days + 1
    
    room_analysis['房型房量'] = room_analysis.index.map(lambda x: room_config.get(x, 0))
    room_analysis['房型OCC(%)'] = room_analysis.apply(
        lambda row: calculate_occ(row['总间晚数'], row['房型房量'], date_range_days),
        axis=1
    )
    
    # 计算占比
    total_room_nights = room_analysis['总间晚数'].sum()
    total_revenue = room_analysis['总收入'].sum()
    room_analysis['间晚占比(%)'] = (room_analysis['总间晚数'] / total_room_nights * 100).round(2)
    room_analysis['收入占比(%)'] = (room_analysis['总收入'] / total_revenue * 100).round(2)
    
    # 格式化日期
    room_analysis['最早入住日期'] = room_analysis['最早入住日期'].dt.strftime('%Y-%m-%d')
    room_analysis['最晚入住日期'] = room_analysis['最晚入住日期'].dt.strftime('%Y-%m-%d')
    
    return room_analysis

def analyze_booking_trend(df, time_unit='day'):
    """分析预订趋势"""
    if time_unit == 'day':
        trend_df = df.groupby(df['预约时间'].dt.date).agg({
            '预约房型': 'count',
            '入住天数': 'sum'
        })
        trend_df.index = pd.to_datetime(trend_df.index)
        trend_title = '按日预订趋势'
    elif time_unit == 'week':
        trend_df = df.groupby(df['预约时间'].dt.isocalendar().week).agg({
            '预约房型': 'count',
            '入住天数': 'sum'
        })
        trend_title = '按周预订趋势'
    
    # 按房型拆分的趋势
    room_trends = {}
    for room_type in df['预约房型'].unique():
        room_df = df[df['预约房型'] == room_type]
        if time_unit == 'day':
            room_trend = room_df.groupby(room_df['预约时间'].dt.date)[['入住天数']].sum()
            room_trend.index = pd.to_datetime(room_trend.index)
        else:
            room_trend = room_df.groupby(room_df['预约时间'].dt.isocalendar().week)[['入住天数']].sum()
        room_trends[room_type] = room_trend
    
    return trend_df, room_trends, trend_title

def analyze_daily_room_nights(df):
    """分析每日入住间夜及房型构成"""
    # 按日期和房型统计间夜数
    daily_room_nights = df.groupby([df['入住日期'].dt.date, '预约房型'])['入住天数'].sum().unstack(fill_value=0)
    
    # 计算每日总间夜
    daily_room_nights['总计'] = daily_room_nights.sum(axis=1)
    
    # 标记节假日
    daily_room_nights['是否节假日'] = [is_holiday(pd.to_datetime(date)) for date in daily_room_nights.index]
    
    return daily_room_nights

def analyze_weekday_distribution(df):
    """分析星期几入住分布"""
    weekday_dist = df.groupby('入住星期中文').agg({
        '入住天数': ['sum', 'mean', 'count'],
        '订单实收': 'sum',
        'ADR': 'mean'
    }).round(2)
    
    # 重新排序
    weekday_order = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    weekday_dist = weekday_dist.reindex(weekday_order)
    
    weekday_dist.columns = ['总间晚数', '平均间晚数', '订单数', '总收入', '平均ADR']
    
    # 计算占比
    total_room_nights = weekday_dist['总间晚数'].sum()
    total_orders = weekday_dist['订单数'].sum()
    total_revenue = weekday_dist['总收入'].sum()
    
    weekday_dist['间晚占比(%)'] = (weekday_dist['总间晚数'] / total_room_nights * 100).round(2)
    weekday_dist['订单占比(%)'] = (weekday_dist['订单数'] / total_orders * 100).round(2)
    weekday_dist['收入占比(%)'] = (weekday_dist['总收入'] / total_revenue * 100).round(2)
    
    return weekday_dist

def analyze_purchase_to_booking(df):
    """分析购买到预约的时间差"""
    time_diff_analysis = df.dropna(subset=['购买到预约时间差']).agg({
        '购买到预约时间差': ['count', 'mean', 'median', 'min', 'max', 'std'],
        '购买到预约天数差': ['mean', 'median', 'min', 'max', 'std']
    }).round(2)
    
    # 时间差区间分析
    bins = [-1, 0, 1, 6, 12, 24, 48, 1000]
    labels = ['即时预约(≤0h)', '0-1小时', '1-6小时', '6-12小时', '12-24小时', '1-2天', '>2天']
    df['时间差区间'] = pd.cut(df['购买到预约时间差'], bins=bins, labels=labels)
    
    time_diff_by_range = df.groupby('时间差区间').agg({
        '购买到预约时间差': ['count', 'mean'],
        '订单实收': 'mean',
        'ADR': 'mean'
    }).round(2)
    
    time_diff_by_range.columns = ['订单数', '平均时间差(小时)', '平均订单金额', '平均ADR']
    time_diff_by_range['订单占比(%)'] = (time_diff_by_range['订单数'] / time_diff_by_range['订单数'].sum() * 100).round(2)
    
    return time_diff_analysis, time_diff_by_range

# ----------------------
# 4. 可视化函数
# ----------------------
def plot_room_type_performance(room_analysis):
    """绘制房型性能分析图表"""
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
    
    # 1. 各房型总间晚数和平均ADR
    x = range(len(room_analysis))
    width = 0.35
    
    bars1 = ax1.bar([i - width/2 for i in x], room_analysis['总间晚数'], width, 
                   label='总间晚数', color='#FF6B6B', alpha=0.8)
    ax1_twin = ax1.twinx()
    line1 = ax1_twin.plot(x, room_analysis['平均ADR'], 'o-', color='#4ECDC4', linewidth=2, label='平均ADR')
    
    ax1.set_title('各房型总间晚数与平均ADR对比', fontweight='bold', fontsize=12)
    ax1.set_xlabel('房型')
    ax1.set_ylabel('总间晚数', color='#FF6B6B')
    ax1_twin.set_ylabel('平均ADR (¥)', color='#4ECDC4')
    ax1.set_xticks(x)
    ax1.set_xticklabels(room_analysis.index, rotation=45, ha='right')
    ax1.grid(True, alpha=0.3)
    
    # 添加数值标签
    for bar in bars1:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height + 5,
                f'{int(height)}', ha='center', va='bottom', fontsize=9)
    
    # 2. 各房型OCC和间晚占比
    bars2 = ax2.bar([i - width/2 for i in x], room_analysis['房型OCC(%)'], width, 
                   label='房型OCC(%)', color='#45B7D1', alpha=0.8)
    ax2_twin = ax2.twinx()
    line2 = ax2_twin.plot(x, room_analysis['间晚占比(%)'], 's-', color='#96CEB4', linewidth=2, label='间晚占比(%)')
    
    ax2.set_title('各房型OCC与间晚占比', fontweight='bold', fontsize=12)
    ax2.set_xlabel('房型')
    ax2.set_ylabel('房型OCC (%)', color='#45B7D1')
    ax2_twin.set_ylabel('间晚占比 (%)', color='#96CEB4')
    ax2.set_xticks(x)
    ax2.set_xticklabels(room_analysis.index, rotation=45, ha='right')
    ax2.grid(True, alpha=0.3)
    
    # 3. 各房型收入占比饼图
    colors = plt.cm.Set3(np.linspace(0, 1, len(room_analysis)))
    wedges, texts, autotexts = ax3.pie(room_analysis['收入占比(%)'], labels=room_analysis.index, 
                                       autopct='%1.1f%%', colors=colors, startangle=90)
    ax3.set_title('各房型收入占比', fontweight='bold', fontsize=12)
    
    # 4. 平均间晚数vs平均ADR散点图
    scatter = ax4.scatter(room_analysis['平均间晚数'], room_analysis['平均ADR'], 
                         s=room_analysis['订单数']/2, alpha=0.7, c=colors)
    ax4.set_title('平均间晚数 vs 平均ADR（气泡大小=订单数）', fontweight='bold', fontsize=12)
    ax4.set_xlabel('平均间晚数')
    ax4.set_ylabel('平均ADR (¥)')
    ax4.grid(True, alpha=0.3)
    
    # 添加房型标签
    for i, room in enumerate(room_analysis.index):
        ax4.annotate(room, (room_analysis['平均间晚数'].iloc[i], room_analysis['平均ADR'].iloc[i]),
                    xytext=(5, 5), textcoords='offset points', fontsize=8)
    
    plt.tight_layout()
    return fig

def plot_booking_trend(trend_df, room_trends, trend_title):
    """绘制预订趋势图表"""
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(16, 10))
    
    # 1. 总体预订趋势
    ax1.plot(trend_df.index, trend_df['预约房型'], 'o-', color='#FF6B6B', linewidth=2, label='订单数')
    ax1_twin = ax1.twinx()
    ax1_twin.plot(trend_df.index, trend_df['入住天数'], 's-', color='#4ECDC4', linewidth=2, label='间晚数')
    
    ax1.set_title(f'{trend_title} - 总体', fontweight='bold', fontsize=12)
    ax1.set_xlabel('时间')
    ax1.set_ylabel('订单数', color='#FF6B6B')
    ax1_twin.set_ylabel('间晚数', color='#4ECDC4')
    ax1.grid(True, alpha=0.3)
    ax1.legend(loc='upper left')
    ax1_twin.legend(loc='upper right')
    
    # 2. 各房型预订趋势
    colors = plt.cm.Set2(np.linspace(0, 1, len(room_trends)))
    for i, (room_type, room_trend) in enumerate(room_trends.items()):
        ax2.plot(room_trend.index, room_trend['入住天数'], 'o-', linewidth=2, 
                label=room_type, color=colors[i], alpha=0.8)
    
    ax2.set_title(f'{trend_title} - 各房型间晚数', fontweight='bold', fontsize=12)
    ax2.set_xlabel('时间')
    ax2.set_ylabel('间晚数')
    ax2.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    ax2.grid(True, alpha=0.3)
    
    plt.tight_layout()
    return fig

def plot_daily_room_nights(daily_room_nights):
    """绘制每日间夜及房型构成图表"""
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(16, 10))
    
    # 1. 每日总间夜（节假日标红）
    dates = pd.to_datetime(daily_room_nights.index)
    holiday_mask = daily_room_nights['是否节假日']
    
    # 绘制非节假日
    ax1.plot(dates[~holiday_mask], daily_room_nights.loc[~holiday_mask, '总计'], 
            'o-', color='#4ECDC4', linewidth=2, label='非节假日', alpha=0.8)
    # 绘制节假日（标红）
    if holiday_mask.any():
        ax1.plot(dates[holiday_mask], daily_room_nights.loc[holiday_mask, '总计'], 
                'o-', color='red', linewidth=2, markersize=8, label='节假日', alpha=0.8)
    
    ax1.set_title('每日入住总间夜数（节假日标红）', fontweight='bold', fontsize=12)
    ax1.set_xlabel('日期')
    ax1.set_ylabel('总间夜数')
    ax1.grid(True, alpha=0.3)
    ax1.legend()
    
    # 格式化x轴日期
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    ax1.xaxis.set_major_locator(mdates.DayLocator(interval=2))
    plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)
    
    # 2. 每日房型构成堆叠图
    room_columns = [col for col in daily_room_nights.columns if col not in ['总计', '是否节假日']]
    daily_room_nights[room_columns].plot(kind='area', stacked=True, ax=ax2, 
                                        colormap='Set2', alpha=0.8)
    
    ax2.set_title('每日入住间夜数房型构成', fontweight='bold', fontsize=12)
    ax2.set_xlabel('日期')
    ax2.set_ylabel('间夜数')
    ax2.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    ax2.grid(True, alpha=0.3)
    
    # 格式化x轴日期
    ax2.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
    ax2.xaxis.set_major_locator(mdates.DayLocator(interval=2))
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45)
    
    plt.tight_layout()
    return fig

def plot_weekday_distribution(weekday_dist):
    """绘制星期几入住分布图表"""
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 10))
    
    # 1. 星期几总间晚数
    x = range(len(weekday_dist))
    bars1 = ax1.bar(x, weekday_dist['总间晚数'], color='#FF6B6B', alpha=0.8)
    ax1.set_title('星期几入住总间晚数', fontweight='bold', fontsize=12)
    ax1.set_xlabel('星期')
    ax1.set_ylabel('总间晚数')
    ax1.set_xticks(x)
    ax1.set_xticklabels(weekday_dist.index)
    ax1.grid(True, alpha=0.3)
    
    # 添加数值标签
    for bar in bars1:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height + 5,
                f'{int(height)}', ha='center', va='bottom')
    
    # 2. 星期几平均ADR
    bars2 = ax2.bar(x, weekday_dist['平均ADR'], color='#4ECDC4', alpha=0.8)
    ax2.set_title('星期几入住平均ADR', fontweight='bold', fontsize=12)
    ax2.set_xlabel('星期')
    ax2.set_ylabel('平均ADR (¥)')
    ax2.set_xticks(x)
    ax2.set_xticklabels(weekday_dist.index)
    ax2.grid(True, alpha=0.3)
    
    # 添加数值标签
    for bar in bars2:
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height + 10,
                f'¥{int(height)}', ha='center', va='bottom')
    
    # 3. 星期几订单占比
    ax3.pie(weekday_dist['订单占比(%)'], labels=weekday_dist.index, autopct='%1.1f%%',
           colors=plt.cm.Set3(np.linspace(0, 1, 7)), startangle=90)
    ax3.set_title('星期几入住订单占比', fontweight='bold', fontsize=12)
    
    # 4. 星期几间晚占比vs收入占比
    width = 0.35
    bars4_1 = ax4.bar([i - width/2 for i in x], weekday_dist['间晚占比(%)'], width, 
                     label='间晚占比', color='#45B7D1', alpha=0.8)
    bars4_2 = ax4.bar([i + width/2 for i in x], weekday_dist['收入占比(%)'], width, 
                     label='收入占比', color='#96CEB4', alpha=0.8)
    
    ax4.set_title('星期几入住间晚占比vs收入占比', fontweight='bold', fontsize=12)
    ax4.set_xlabel('星期')
    ax4.set_ylabel('占比 (%)')
    ax4.set_xticks(x)
    ax4.set_xticklabels(weekday_dist.index)
    ax4.legend()
    ax4.grid(True, alpha=0.3)
    
    plt.tight_layout()
    return fig

def plot_purchase_to_booking(time_diff_by_range):
    """绘制购买到预约时间差分析图表"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    
    # 1. 各时间区间订单数分布
    x = range(len(time_diff_by_range))
    bars1 = ax1.bar(x, time_diff_by_range['订单数'], color='#FF6B6B', alpha=0.8)
    ax1.set_title('购买到预约时间差区间订单数分布', fontweight='bold', fontsize=12)
    ax1.set_xlabel('时间差区间')
    ax1.set_ylabel('订单数')
    ax1.set_xticks(x)
    ax1.set_xticklabels(time_diff_by_range.index, rotation=45, ha='right')
    ax1.grid(True, alpha=0.3)
    
    # 添加数值标签
    for bar in bars1:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height + 1,
                f'{int(height)}', ha='center', va='bottom')
    
    # 2. 各时间区间平均ADR
    bars2 = ax2.bar(x, time_diff_by_range['平均ADR'], color='#4ECDC4', alpha=0.8)
    ax2.set_title('购买到预约时间差区间平均ADR', fontweight='bold', fontsize=12)
    ax2.set_xlabel('时间差区间')
    ax2.set_ylabel('平均ADR (¥)')
    ax2.set_xticks(x)
    ax2.set_xticklabels(time_diff_by_range.index, rotation=45, ha='right')
    ax2.grid(True, alpha=0.3)
    
    # 添加数值标签
    for bar in bars2:
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height + 10,
                f'¥{int(height)}', ha='center', va='bottom')
    
    plt.tight_layout()
    return fig

# ----------------------
# 5. 主页面内容
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
    
    # 酒店配置
    total_rooms, room_config = configure_hotel_rooms()
    
    if uploaded_file is None:
        st.warning("请在左侧边栏上传Excel数据文件，并配置酒店房间信息")
        st.stop()
    
    # 加载数据
    with st.spinner("正在加载和预处理数据..."):
        df, valid_df = load_and_preprocess_data(uploaded_file)
        st.session_state['valid_df'] = valid_df
    
    # 侧边栏数据概览
    st.sidebar.markdown("### 📊 数据概览")
    st.sidebar.write(f"总记录数: **{len(df):,}**")
    st.sidebar.write(f"有效记录数: **{len(valid_df):,}**")
    
    if len(valid_df) > 0:
        st.sidebar.write(f"数据时间范围:")
        st.sidebar.write(f"- 预约时间: {valid_df['预约时间'].min().strftime('%Y-%m-%d')} 至 {valid_df['预约时间'].max().strftime('%Y-%m-%d')}")
        st.sidebar.write(f"- 入住时间: {valid_df['入住日期'].min().strftime('%Y-%m-%d')} 至 {valid_df['入住日期'].max().strftime('%Y-%m-%d')}")
        
        # 计算整体OCC
        total_room_nights = valid_df['入住天数'].sum()
        date_start = valid_df['入住日期'].min()
        date_end = valid_df['入住日期'].max()
        date_range_days = (date_end - date_start).days + 1
        overall_occ = calculate_occ(total_room_nights, total_rooms, date_range_days)
        st.sidebar.metric("整体OCC", f"{overall_occ}%")
    
    # 侧边栏分析选项
    st.sidebar.markdown("### 📈 分析选项")
    analysis_type = st.sidebar.radio(
        "选择分析类型",
        ["数据概览", "房型深度分析", "预订趋势分析", "入住间夜分析", 
         "入住时间分析", "购买预约时间差", "OCC分析", "详细数据"]
    )
    
    # 主内容区域
    st.title("🏨 酒店预约入住数据分析系统")
    st.markdown("---")
    
    # ----------------------
    # 6. 数据概览
    # ----------------------
    if analysis_type == "数据概览":
        st.header("📊 数据整体概览")
        
        # 关键指标卡片
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("总预约单数", f"{len(valid_df):,}")
        
        with col2:
            total_room_nights = valid_df['入住天数'].sum()
            st.metric("总间夜数", f"{total_room_nights:,}")
        
        with col3:
            avg_adr = valid_df['ADR'].mean() if 'ADR' in valid_df.columns else 0
            st.metric("平均ADR", f"¥{avg_adr:.2f}")
        
        with col4:
            total_revenue = valid_df['订单实收'].sum() if '订单实收' in valid_df.columns else 0
            st.metric("总实收金额", f"¥{total_revenue:,.2f}")
        
        with col5:
            overall_occ = calculate_occ(total_room_nights, total_rooms, date_range_days)
            st.metric("整体OCC", f"{overall_occ}%")
        
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
                f"{valid_df['入住天数'].sum():,}",
                f"{valid_df['入住天数'].mean():.1f} 天/单",
                f"¥{valid_df['ADR'].mean():.2f}",
                f"¥{valid_df['订单实收'].sum():,.2f}",
                f"{overall_occ}%",
                f"{valid_df['预约房型'].value_counts().index[0]}",
                f"{valid_df['入住日期'].min().strftime('%Y-%m-%d')} 至 {valid_df['入住日期'].max().strftime('%Y-%m-%d')}",
                f"{valid_df['是否节假日'].sum() / len(valid_df) * 100:.1f}%",
                f"{valid_df['购买到预约时间差'].mean():.1f} 小时" if '购买到预约时间差' in valid_df.columns else "N/A"
            ]
        }
        stats_df = pd.DataFrame(stats_data)
        st.dataframe(stats_df, use_container_width=True)
    
    # ----------------------
    # 7. 房型深度分析
    # ----------------------
    elif analysis_type == "房型深度分析":
        st.header("🏠 房型深度分析")
        
        # 筛选器
        col1, col2 = st.columns(2)
        with col1:
            room_types = valid_df['预约房型'].unique()
            selected_room = st.multiselect(
                "选择房型（可多选）", 
                room_types, 
                default=room_types,
                help="选择要分析的房型"
            )
        
        with col2:
            date_range = st.date_input(
                "选择日期范围",
                [valid_df['入住日期'].min().date(), valid_df['入住日期'].max().date()],
                help="选择分析的入住日期范围"
            )
        
        # 应用筛选
        filtered_df = valid_df[
            (valid_df['预约房型'].isin(selected_room)) &
            (valid_df['入住日期'].dt.date >= date_range[0]) &
            (valid_df['入住日期'].dt.date <= date_range[1])
        ]
        
        if len(filtered_df) == 0:
            st.warning("没有符合筛选条件的数据，请调整筛选参数")
            st.stop()
        
        # 执行分析
        room_analysis = analyze_room_type_performance(filtered_df, room_config)
        
        # 显示分析结果表格
        st.subheader("1. 房型性能指标汇总")
        st.dataframe(room_analysis, use_container_width=True)
        
        # 可视化
        st.subheader("2. 房型性能可视化分析")
        fig = plot_room_type_performance(room_analysis)
        st.pyplot(fig)
        
        # 关键洞察
        st.subheader("3. 关键洞察")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            best_occ_room = room_analysis['房型OCC(%)'].idxmax()
            st.info(f"""
            **最高OCC房型**  
            房型: {best_occ_room}  
            OCC: {room_analysis.loc[best_occ_room, '房型OCC(%)']}%  
            间晚数: {room_analysis.loc[best_occ_room, '总间晚数']:,}
            """)
        
        with col2:
            best_adr_room = room_analysis['平均ADR'].idxmax()
            st.info(f"""
            **最高ADR房型**  
            房型: {best_adr_room}  
            平均ADR: ¥{room_analysis.loc[best_adr_room, '平均ADR']:.2f}  
            总收入: ¥{room_analysis.loc[best_adr_room, '总收入']:,.2f}
            """)
        
        with col3:
            top_revenue_room = room_analysis['收入占比(%)'].idxmax()
            st.info(f"""
            **收入贡献最大房型**  
            房型: {top_revenue_room}  
            收入占比: {room_analysis.loc[top_revenue_room, '收入占比(%)']}%  
            间晚占比: {room_analysis.loc[top_revenue_room, '间晚占比(%)']}%
            """)
    
    # ----------------------
    # 8. 预订趋势分析
    # ----------------------
    elif analysis_type == "预订趋势分析":
        st.header("📈 预订趋势分析")
        
        # 筛选器
        col1, col2 = st.columns(2)
        with col1:
            time_unit = st.radio(
                "选择时间维度",
                ['day', 'week'],
                format_func=lambda x: '按日' if x == 'day' else '按周',
                horizontal=True
            )
        
        with col2:
            selected_rooms = st.multiselect(
                "选择要显示的房型",
                valid_df['预约房型'].unique(),
                default=valid_df['预约房型'].unique()
            )
        
        # 筛选数据
        filtered_df = valid_df[valid_df['预约房型'].isin(selected_rooms)]
        
        # 执行分析
        trend_df, room_trends, trend_title = analyze_booking_trend(filtered_df, time_unit)
        
        # 显示趋势数据
        st.subheader(f"1. {trend_title}数据")
        st.dataframe(trend_df, use_container_width=True)
        
        # 可视化
        st.subheader(f"2. {trend_title}可视化")
        fig = plot_booking_trend(trend_df, room_trends, trend_title)
        st.pyplot(fig)
        
        # 趋势分析
        st.subheader("3. 趋势分析")
        peak_date = trend_df['入住天数'].idxmax()
        peak_value = trend_df['入住天数'].max()
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"""
            **预订高峰**  
            时间: {peak_date}  
            间晚数: {peak_value:,}  
            订单数: {trend_df.loc[peak_date, '预约房型']:,}
            """)
        
        with col2:
            avg_daily_nights = trend_df['入住天数'].mean()
            st.info(f"""
            **平均预订量**  
            日均/周间晚数: {avg_daily_nights:.1f}  
            日均/周订单数: {trend_df['预约房型'].mean():.1f}  
            趋势波动性: {trend_df['入住天数'].std() / avg_daily_nights:.2f}
            """)
    
    # ----------------------
    # 9. 入住间夜分析
    # ----------------------
    elif analysis_type == "入住间夜分析":
        st.header("🛏️ 入住间夜及房型构成分析")
        
        # 执行分析
        daily_room_nights = analyze_daily_room_nights(valid_df)
        
        # 显示每日间夜数据
        st.subheader("1. 每日入住间夜及房型构成")
        st.dataframe(daily_room_nights, use_container_width=True)
        
        # 可视化
        st.subheader("2. 每日间夜分析可视化")
        fig = plot_daily_room_nights(daily_room_nights)
        st.pyplot(fig)
        
        # 关键统计
        st.subheader("3. 关键统计")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            peak_night_date = daily_room_nights['总计'].idxmax()
            peak_night_value = daily_room_nights['总计'].max()
            is_holiday_peak = daily_room_nights.loc[peak_night_date, '是否节假日']
            st.info(f"""
            **入住高峰日**  
            日期: {peak_night_date}  
            总间夜数: {peak_night_value:,}  
            节假日: {'是' if is_holiday_peak else '否'}
            """)
        
        with col2:
            holiday_nights = daily_room_nights[daily_room_nights['是否节假日']]['总计'].sum()
            total_nights = daily_room_nights['总计'].sum()
            holiday_ratio = (holiday_nights / total_nights * 100) if total_nights > 0 else 0
            st.info(f"""
            **节假日入住占比**  
            节假日总间夜: {holiday_nights:,}  
            总间夜数: {total_nights:,}  
            占比: {holiday_ratio:.1f}%
            """)
        
        with col3:
            # 主要房型构成
            room_columns = [col for col in daily_room_nights.columns if col not in ['总计', '是否节假日']]
            room_totals = daily_room_nights[room_columns].sum()
            main_room = room_totals.idxmax()
            main_room_ratio = (room_totals[main_room] / total_nights * 100) if total_nights > 0 else 0
            st.info(f"""
            **主要房型构成**  
            主要房型: {main_room}  
            间夜数: {room_totals[main_room]:,}  
            占比: {main_room_ratio:.1f}%
            """)
    
    # ----------------------
    # 10. 入住时间分析
    # ----------------------
    elif analysis_type == "入住时间分析":
        st.header("⏰ 入住时间维度分析")
        
        # 执行分析
        weekday_dist = analyze_weekday_distribution(valid_df)
        
        # 显示星期分布数据
        st.subheader("1. 星期几入住分布")
        st.dataframe(weekday_dist, use_container_width=True)
        
        # 可视化
        st.subheader("2. 入住时间可视化分析")
        fig = plot_weekday_distribution(weekday_dist)
        st.pyplot(fig)
        
        # 关键洞察
        st.subheader("3. 关键洞察")
        col1, col2 = st.columns(2)
        
        with col1:
            peak_weekday = weekday_dist['总间晚数'].idxmax()
            st.info(f"""
            **入住高峰星期**  
            星期: {peak_weekday}  
            总间夜数: {weekday_dist.loc[peak_weekday, '总间晚数']:,}  
            占比: {weekday_dist.loc[peak_weekday, '间晚占比(%)']}%
            """)
        
        with col2:
            highest_adr_weekday = weekday_dist['平均ADR'].idxmax()
            st.info(f"""
            **最高ADR星期**  
            星期: {highest_adr_weekday}  
            平均ADR: ¥{weekday_dist.loc[highest_adr_weekday, '平均ADR']:.2f}  
            总收入: ¥{weekday_dist.loc[highest_adr_weekday, '总收入']:,.2f}
            """)
    
    # ----------------------
    # 11. 购买预约时间差
    # ----------------------
    elif analysis_type == "购买-预约时间差":
        st.header("📅 购买到预约时间差分析")
        
        if '购买到预约时间差' not in valid_df.columns:
            st.warning("数据中缺少下单时间或预约时间字段，无法计算时间差")
            st.stop()
        
        # 执行分析
        time_diff_analysis, time_diff_by_range = analyze_purchase_to_booking(valid_df)
        
        # 显示时间差统计
        st.subheader("1. 购买到预约时间差统计")
        st.dataframe(time_diff_analysis, use_container_width=True)
        
        # 显示区间分析
        st.subheader("2. 时间差区间分析")
        st.dataframe(time_diff_by_range, use_container_width=True)
        
        # 可视化
        st.subheader("3. 时间差分析可视化")
        fig = plot_purchase_to_booking(time_diff_by_range)
        st.pyplot(fig)
        
        # 关键洞察
        st.subheader("4. 关键洞察")
        col1, col2 = st.columns(2)
        
        with col1:
            instant_booking_ratio = time_diff_by_range.loc['即时预约(≤0h)', '订单占比(%)']
            st.info(f"""
            **即时预约占比**  
            订单数: {time_diff_by_range.loc['即时预约(≤0h)', '订单数']:,}  
            占比: {instant_booking_ratio}%  
            平均ADR: ¥{time_diff_by_range.loc['即时预约(≤0h)', '平均ADR']:.2f}
            """)
        
        with col2:
            avg_time_diff = time_diff_analysis.loc['mean', '购买到预约时间差']
            median_time_diff = time_diff_analysis.loc['median', '购买到预约时间差']
            st.info(f"""
            **时间差统计**  
            平均时间差: {avg_time_diff:.1f} 小时  
            中位时间差: {median_time_diff:.1f} 小时  
            最长时间差: {time_diff_analysis.loc['max', '购买到预约时间差']:.1f} 小时
            """)
    
    # ----------------------
    # 12. OCC分析
    # ----------------------
    elif analysis_type == "OCC分析":
        st.header("📊 入住率(OCC)深度分析")
        
        # 按时间段计算OCC
        st.subheader("1. 整体OCC趋势")
        
        # 按日期计算每日OCC
        daily_occ = valid_df.groupby(valid_df['入住日期'].dt.date).agg({
            '入住天数': 'sum'
        })
        daily_occ.columns = ['总间晚数']
        daily_occ['当日OCC(%)'] = (daily_occ['总间晚数'] / total_rooms * 100).round(2)
        daily_occ['是否节假日'] = [is_holiday(pd.to_datetime(date)) for date in daily_occ.index]
        
        # 显示每日OCC
        st.dataframe(daily_occ, use_container_width=True)
        
        # 可视化OCC趋势
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(16, 10))
        
        # 1. 每日OCC趋势（节假日标红）
        dates = pd.to_datetime(daily_occ.index)
        holiday_mask = daily_occ['是否节假日']
        
        ax1.plot(dates[~holiday_mask], daily_occ.loc[~holiday_mask, '当日OCC(%)'], 
                'o-', color='#4ECDC4', linewidth=2, label='非节假日', alpha=0.8)
        if holiday_mask.any():
            ax1.plot(dates[holiday_mask], daily_occ.loc[holiday_mask, '当日OCC(%)'], 
                    'o-', color='red', linewidth=2, markersize=8, label='节假日', alpha=0.8)
        
        ax1.axhline(daily_occ['当日OCC(%)'].mean(), color='black', linestyle='--', 
                   label=f'平均OCC: {daily_occ["当日OCC(%)"].mean():.1f}%')
        ax1.set_title('每日OCC趋势（节假日标红）', fontweight='bold', fontsize=12)
        ax1.set_xlabel('日期')
        ax1.set_ylabel('当日OCC (%)')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        
        # 格式化x轴
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
        ax1.xaxis.set_major_locator(mdates.DayLocator(interval=2))
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)
        
        # 2. 各房型OCC对比
        room_analysis = analyze_room_type_performance(valid_df, room_config)
        x = range(len(room_analysis))
        bars = ax2.bar(x, room_analysis['房型OCC(%)'], color=plt.cm.Set3(np.linspace(0, 1, len(room_analysis))), alpha=0.8)
        ax2.axhline(overall_occ, color='red', linestyle='--', label=f'整体OCC: {overall_occ}%')
        
        ax2.set_title('各房型OCC对比', fontweight='bold', fontsize=12)
        ax2.set_xlabel('房型')
        ax2.set_ylabel('房型OCC (%)')
        ax2.set_xticks(x)
        ax2.set_xticklabels(room_analysis.index, rotation=45, ha='right')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        # 添加数值标签
        for bar in bars:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{height:.1f}%', ha='center', va='bottom', fontsize=9)
        
        plt.tight_layout()
        st.pyplot(fig)
        
        # OCC统计
        st.subheader("2. OCC关键统计")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            max_occ_date = daily_occ['当日OCC(%)'].idxmax()
            max_occ_value = daily_occ['当日OCC(%)'].max()
            st.info(f"""
            **最高OCC日期**  
            日期: {max_occ_date}  
            OCC: {max_occ_value:.1f}%  
            间夜数: {daily_occ.loc[max_occ_date, '总间晚数']:,}
            """)
        
        with col2:
            holiday_occ = daily_occ[daily_occ['是否节假日']]['当日OCC(%)'].mean()
            weekday_occ = daily_occ[~daily_occ['是否节假日']]['当日OCC(%)'].mean()
            st.info(f"""
            **节假日vs平日OCC**  
            节假日平均OCC: {holiday_occ:.1f}%  
            平日平均OCC: {weekday_occ:.1f}%  
            差值: {holiday_occ - weekday_occ:.1f}%
            """)
        
        with col3:
            # OCC达标率（假设80%为达标线）
            target_occ = 80
            days_above_target = (daily_occ['当日OCC(%)'] >= target_occ).sum()
            total_days = len(daily_occ)
            target_ratio = (days_above_target / total_days * 100) if total_days > 0 else 0
            st.info(f"""
            **OCC达标率**  
            目标OCC: {target_occ}%  
            达标天数: {days_above_target}/{total_days}  
            达标率: {target_ratio:.1f}%
            """)
    
    # ----------------------
    # 13. 详细数据
    # ----------------------
    elif analysis_type == "详细数据":
        st.header("📋 详细数据查询与导出")
        
        # 保留原有的详细数据功能...
        # 数据筛选器
        st.subheader("数据筛选")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            room_filter = st.multiselect(
                "筛选房型",
                valid_df['预约房型'].unique(),
                default=valid_df['预约房型'].unique()
            )
        
        with col2:
            stay_filter = st.multiselect(
                "筛选入住天数",
                sorted(valid_df['入住天数'].unique()),
                default=sorted(valid_df['入住天数'].unique())
            )
        
        with col3:
            price_range = st.slider(
                "筛选实收金额范围（¥）",
                min_value=float(valid_df['订单实收'].min()),
                max_value=float(valid_df['订单实收'].max()),
                value=(float(valid_df['订单实收'].min()), float(valid_df['订单实收'].max())),
                step=10.0
            )
        
        # 应用筛选
        filtered_df = valid_df[
            (valid_df['预约房型'].isin(room_filter)) &
            (valid_df['入住天数'].isin(stay_filter)) &
            (valid_df['订单实收'] >= price_range[0]) &
            (valid_df['订单实收'] <= price_range[1])
        ]
        
        st.write(f"筛选结果：共 **{len(filtered_df):,}** 条记录")
        
        # 选择显示字段
        st.subheader("选择显示字段")
        all_columns = valid_df.columns.tolist()
        
        # 常用字段推荐
        common_fields = [
            '订单ID', '预约单号', '预约房型', '预约门店名称', '订单实收', 'ADR',
            '入住日期', '离店日期', '入住天数', '下单时间', '预约时间',
            '预约状态', '是否节假日', '购买到预约时间差'
        ]
        common_fields = [f for f in common_fields if f in all_columns]
        
        # 字段选择
        selected_fields = st.multiselect(
            "选择要显示的字段（默认显示常用字段）",
            all_columns,
            default=common_fields
        )
        
        if not selected_fields:
            st.warning("请至少选择一个字段")
            st.stop()
        
        # 显示数据
        st.subheader("详细数据表格")
        display_df = filtered_df[selected_fields].copy()
        
        # 格式化日期时间字段
        datetime_fields = ['下单时间', '支付时间', '预约时间', '入住日期', '离店日期']
        for field in datetime_fields:
            if field in display_df.columns:
                display_df[field] = display_df[field].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        # 分页显示
        page_size = st.selectbox("每页显示数量", [20, 50, 100, 200], index=0)
        total_pages = (len(display_df) + page_size - 1) // page_size
        
        if total_pages > 1:
            page = st.number_input("页码", min_value=1, max_value=total_pages, value=1)
            start_idx = (page - 1) * page_size
            end_idx = start_idx + page_size
            paginated_df = display_df.iloc[start_idx:end_idx]
            st.dataframe(paginated_df, use_container_width=True, height=600)
            st.write(f"显示第 {page}/{total_pages} 页，共 {len(display_df):,} 条记录")
        else:
            st.dataframe(display_df, use_container_width=True, height=600)
        
        # 数据导出
        st.subheader("数据导出")
        export_format = st.radio("选择导出格式", ["CSV", "Excel"])
        
        # 准备导出数据
        export_df = filtered_df.copy()
        
        # 格式化日期时间
        for field in datetime_fields:
            if field in export_df.columns:
                export_df[field] = export_df[field].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        # 生成导出文件
        if export_format == "CSV":
            csv_data = export_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="下载CSV文件",
                data=csv_data,
                file_name=f"酒店预约数据_{datetime.now().strftime('%Y%m%d%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            # 使用ExcelWriter保存多个sheet
            from io import BytesIO
            with pd.ExcelWriter(BytesIO(), engine='openpyxl') as writer:
                # 主要数据
                export_df.to_excel(writer, sheet_name='详细数据', index=False)
                
                # 房型统计
                room_summary = analyze_room_type_performance(filtered_df, room_config)
                room_summary.to_excel(writer, sheet_name='房型统计')
                
                # 每日OCC
                daily_occ = filtered_df.groupby(filtered_df['入住日期'].dt.date).agg({
                    '入住天数': 'sum'
                })
                daily_occ.columns = ['总间晚数']
                daily_occ['当日OCC(%)'] = (daily_occ['总间晚数'] / total_rooms * 100).round(2)
                daily_occ.to_excel(writer, sheet_name='每日OCC')
            
            writer.handles['output'].seek(0)
            st.download_button(
                label="下载Excel文件（含多个工作表）",
                data=writer.handles['output'],
                file_name=f"酒店预约数据_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ----------------------
# 运行应用
# ----------------------
if __name__ == "__main__":
    main()