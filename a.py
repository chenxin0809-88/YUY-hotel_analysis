import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os

st.set_page_config(page_title="酒店预约数据分析（上传版）", layout="wide")
st.title("🏨 酒店预约数据分析平台（支持上传）")
st.markdown("---")

# ---------- 文件上传 ----------
st.sidebar.header("1. 上传数据文件")
uploaded_file = st.sidebar.file_uploader(
    "请上传 Excel 文件（格式必须与示例一致）",
    type=["xlsx", "xls"]
)

# ---------- 加载数据 ----------
@st.cache_data
def load_data(file):
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"读取文件失败：{e}")
        st.stop()

if uploaded_file is not None:
    df = load_data(uploaded_file)
    st.sidebar.success(f"✅ 已加载 {len(df)} 行数据")
else:
    st.sidebar.warning("请上传一个 Excel 文件开始分析")
    st.stop()  # 如果没有上传文件，停止执行后续代码

# ---------- 数据预处理 ----------
# 确保必要的列存在
required_cols = ['入住日期', '离店日期', '预约房型', '预约状态', '预约时间', '订单实收', '商家补贴']
missing = [col for col in required_cols if col not in df.columns]
if missing:
    st.error(f"上传的文件缺少以下必要列：{missing}")
    st.stop()

df['入住日期'] = pd.to_datetime(df['入住日期'], errors='coerce')
df['离店日期'] = pd.to_datetime(df['离店日期'], errors='coerce')
df['预约时间'] = pd.to_datetime(df['预约时间'], errors='coerce')
df = df.dropna(subset=['入住日期', '离店日期', '预约房型', '预约状态'])

df['入住星期'] = df['入住日期'].dt.day_name()
df['入住月份日'] = df['入住日期'].dt.strftime('%m-%d')
df['预约日期'] = df['预约时间'].dt.date
df['预约小时'] = df['预约时间'].dt.hour
df['间晚数'] = (df['离店日期'] - df['入住日期']).dt.days
df['提前天数'] = (df['入住日期'] - df['预约时间']).dt.days
df['订单实收'] = pd.to_numeric(df['订单实收'], errors='coerce')

# ---------- 侧边栏：筛选条件 ----------
st.sidebar.header("2. 筛选条件")
selected_status = st.sidebar.multiselect(
    "预约状态", df['预约状态'].unique(), default=df['预约状态'].unique()
)
selected_room = st.sidebar.multiselect(
    "预约房型", df['预约房型'].unique(), default=df['预约房型'].unique()
)
date_range = st.sidebar.date_input(
    "入住日期范围",
    value=(df['入住日期'].min().date(), df['入住日期'].max().date())
)

mask = (
    df['预约状态'].isin(selected_status) &
    df['预约房型'].isin(selected_room) &
    (df['入住日期'].dt.date >= date_range[0]) &
    (df['入住日期'].dt.date <= date_range[1])
)
filtered_df = df[mask].copy()

st.sidebar.markdown(f"**当前数据行数:** {len(filtered_df)}")

# ---------- 侧边栏：酒店房量配置 ----------
st.sidebar.markdown("---")
st.sidebar.subheader("🏨 酒店房量配置")
all_room_types = sorted(df['预约房型'].unique())
room_inventory = {}
for rt in all_room_types:
    key = f"inv_{rt}"
    default_val = st.session_state.get(key, 0)
    room_inventory[rt] = st.sidebar.number_input(
        f"{rt} 房间数", min_value=0, value=default_val, step=1, key=key
    )
total_rooms = sum(room_inventory.values())
st.sidebar.write(f"**总计房间数：{total_rooms}**")

# ---------- 节假日定义 ----------
holiday_strs = ["2026-02-17"]  # 可在此修改或添加更多节假日
holidays = pd.to_datetime(holiday_strs)

# ---------- 计算每日入住间夜 ----------
daily_nights = filtered_df.groupby('入住日期')['间晚数'].sum().reset_index(name='入住间夜')
daily_nights = daily_nights.set_index('入住日期').asfreq('D', fill_value=0).reset_index()

room_daily = filtered_df.groupby(['入住日期', '预约房型'])['间晚数'].sum().reset_index()
room_daily_pivot = room_daily.pivot(index='入住日期', columns='预约房型', values='间晚数').fillna(0)

# ---------- 主面板 ----------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 房型深度分析", "📅 入住时间分析", "⏰ 预约行为分析",
    "💰 收入与补贴", "📈 入住率与构成"
])

# -------------------- 标签页1：房型深度分析 --------------------
with tab1:
    st.header("房型深度分析")
    
    # 平均间晚数和平均ADR
    room_stats = filtered_df.groupby('预约房型').agg(
        平均间晚数=('间晚数', 'mean'),
        平均ADR=('订单实收', lambda x: x.sum() / (filtered_df.loc[x.index, '间晚数'].sum() if filtered_df.loc[x.index, '间晚数'].sum()>0 else 1))
    ).round(2).reset_index()
    room_stats.columns = ['房型', '平均间晚数', '平均ADR（元/间夜）']
    st.subheader("各房型核心指标")
    st.dataframe(room_stats, use_container_width=True)
    
    # 各房型每日预订趋势
    st.subheader("各房型每日预订间夜趋势")
    daily_room = filtered_df.groupby(['入住日期', '预约房型'])['间晚数'].sum().reset_index()
    fig_room_trend = px.line(daily_room, x='入住日期', y='间晚数', color='预约房型',
                             title="每日各房型入住间夜数", markers=True)
    for h in holidays:
        fig_room_trend.add_shape(
            type="line",
            x0=h, x1=h,
            y0=0, y1=1,
            yref="paper",
            line=dict(color="red", width=2, dash="dash")
        )
        fig_room_trend.add_annotation(
            x=h, y=1, yref="paper",
            text="春节", showarrow=False,
            font=dict(color="red")
        )
    st.plotly_chart(fig_room_trend, use_container_width=True)
    
    # 各房型每周预订趋势
    st.subheader("各房型每周预订间夜趋势")
    filtered_df['周数'] = filtered_df['入住日期'].dt.isocalendar().week
    weekly_room = filtered_df.groupby(['周数', '预约房型'])['间晚数'].sum().reset_index()
    fig_weekly = px.line(weekly_room, x='周数', y='间晚数', color='预约房型',
                         title="每周各房型入住间夜数", markers=True)
    st.plotly_chart(fig_weekly, use_container_width=True)
    
    # 提前预订天数分布（按房型）
    st.subheader("各房型提前预订天数分布")
    fig_advance = px.box(filtered_df[filtered_df['提前天数']>=0], x='预约房型', y='提前天数',
                         title="提前预订天数（按房型）", points="all")
    st.plotly_chart(fig_advance, use_container_width=True)

# -------------------- 标签页2：入住时间分析 --------------------
with tab2:
    st.header("入住时间分析")
    col1, col2 = st.columns(2)
    
    with col1:
        fig_daily = px.line(daily_nights, x='入住日期', y='入住间夜',
                            title="每日入住间夜数", markers=True)
        for h in holidays:
            fig_daily.add_shape(
                type="line",
                x0=h, x1=h,
                y0=0, y1=1,
                yref="paper",
                line=dict(color="red", width=2, dash="dash")
            )
            fig_daily.add_annotation(
                x=h, y=1, yref="paper",
                text="春节", showarrow=False,
                font=dict(color="red")
            )
        st.plotly_chart(fig_daily, use_container_width=True)
    
    with col2:
        weekday_counts = filtered_df['入住星期'].value_counts().reindex(
            ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        ).reset_index()
        weekday_counts.columns = ['星期', '订单数']
        fig_weekday = px.bar(weekday_counts, x='星期', y='订单数',
                             title="各星期入住订单数分布", color='星期')
        st.plotly_chart(fig_weekday, use_container_width=True)
    
    # 每日房型构成堆积图
    st.subheader("每日入住间夜房型构成")
    room_daily_stack = room_daily_pivot.reset_index().melt(id_vars='入住日期', var_name='房型', value_name='间夜数')
    fig_stack = px.bar(room_daily_stack, x='入住日期', y='间夜数', color='房型',
                       title="每日各房型入住间夜堆积图", barmode='stack')
    for h in holidays:
        fig_stack.add_shape(
            type="line",
            x0=h, x1=h,
            y0=0, y1=1,
            yref="paper",
            line=dict(color="red", width=2, dash="dash")
        )
        fig_stack.add_annotation(
            x=h, y=1, yref="paper",
            text="春节", showarrow=False,
            font=dict(color="red")
        )
    st.plotly_chart(fig_stack, use_container_width=True)
    
    with st.expander("查看每日房型构成明细"):
        st.dataframe(room_daily_pivot, use_container_width=True)

# -------------------- 标签页3：预约行为分析 --------------------
with tab3:
    st.header("预约行为分析")
    col1, col2 = st.columns(2)
    with col1:
        hour_counts = filtered_df['预约小时'].value_counts().sort_index().reset_index()
        hour_counts.columns = ['小时', '预约数量']
        fig_hour = px.bar(hour_counts, x='小时', y='预约数量', title="预约小时分布")
        st.plotly_chart(fig_hour, use_container_width=True)
    with col2:
        daily_order = filtered_df.groupby('预约日期').size().reset_index(name='数量')
        fig_order_trend = px.line(daily_order, x='预约日期', y='数量', title="每日预约趋势", markers=True)
        st.plotly_chart(fig_order_trend, use_container_width=True)
    
    st.subheader("提前预订天数分布")
    fig_adv_hist = px.histogram(filtered_df[filtered_df['提前天数']>=0], x='提前天数', nbins=30,
                                title="提前预订天数分布")
    st.plotly_chart(fig_adv_hist, use_container_width=True)

# -------------------- 标签页4：收入与补贴 --------------------
with tab4:
    st.header("收入与补贴分析")
    col1, col2 = st.columns(2)
    with col1:
        fig_income = px.histogram(filtered_df, x='订单实收', nbins=30, title="订单实收分布")
        st.plotly_chart(fig_income, use_container_width=True)
    with col2:
        avg_room_income = filtered_df.groupby('预约房型')['订单实收'].mean().round(2).reset_index()
        avg_room_income.columns = ['房型', '平均订单实收']
        fig_avg_income = px.bar(avg_room_income, x='房型', y='平均订单实收', color='房型',
                                title="各房型平均订单实收", text='平均订单实收')
        fig_avg_income.update_layout(showlegend=False)
        st.plotly_chart(fig_avg_income, use_container_width=True)
    
    st.subheader("补贴与实收关系")
    fig_subsidy = px.scatter(filtered_df, x='商家补贴', y='订单实收', color='预约房型',
                             size='间晚数', title="商家补贴 vs 订单实收")
    st.plotly_chart(fig_subsidy, use_container_width=True)

# -------------------- 标签页5：入住率与构成 --------------------
with tab5:
    st.header("入住率分析（需填写酒店房量）")
    if total_rooms == 0:
        st.warning("请在侧边栏填写各房型房间数以计算入住率。")
    else:
        daily_occ = daily_nights.copy()
        daily_occ['入住率'] = (daily_occ['入住间夜'] / total_rooms * 100).round(1)
        fig_occ = px.line(daily_occ, x='入住日期', y='入住率', title="每日入住率 (%)", markers=True)
        for h in holidays:
            fig_occ.add_shape(
                type="line",
                x0=h, x1=h,
                y0=0, y1=1,
                yref="paper",
                line=dict(color="red", width=2, dash="dash")
            )
            fig_occ.add_annotation(
                x=h, y=1, yref="paper",
                text="春节", showarrow=False,
                font=dict(color="red")
            )
        st.plotly_chart(fig_occ, use_container_width=True)
        
        st.subheader("分房型每日入住率")
        occ_by_room = []
        for rt in all_room_types:
            if rt in room_inventory and room_inventory[rt] > 0:
                if rt in room_daily_pivot.columns:
                    rt_daily = room_daily_pivot[rt].reset_index()
                    rt_daily.columns = ['入住日期', '间夜数']
                    rt_daily['入住率'] = (rt_daily['间夜数'] / room_inventory[rt] * 100).round(1)
                    rt_daily['房型'] = rt
                    occ_by_room.append(rt_daily)
        if occ_by_room:
            occ_df = pd.concat(occ_by_room, ignore_index=True)
            fig_occ_room = px.line(occ_df, x='入住日期', y='入住率', color='房型',
                                   title="各房型每日入住率 (%)", markers=True)
            for h in holidays:
                fig_occ_room.add_shape(
                    type="line",
                    x0=h, x1=h,
                    y0=0, y1=1,
                    yref="paper",
                    line=dict(color="red", width=2, dash="dash")
                )
            st.plotly_chart(fig_occ_room, use_container_width=True)
        else:
            st.info("请填写各房型房间数以查看分房型入住率。")
    
    st.subheader("入住间夜构成明细（每日）")
    daily_detail = filtered_df.groupby(['入住日期', '预约房型'])['间晚数'].sum().reset_index()
    st.dataframe(daily_detail.sort_values('入住日期'), use_container_width=True)

# ---------- 底部 ----------
st.markdown("---")
st.header("原始数据预览")
st.dataframe(filtered_df.head(100))

csv = filtered_df.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 下载筛选后的数据 (CSV)",
    data=csv,
    file_name='filtered_bookings.csv',
    mime='text/csv'
)