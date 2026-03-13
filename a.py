import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re

st.set_page_config(page_title="酒店精细化运营看板", layout="wide")

st.title("🏨 酒店精细化运营与OCC分析系统")
st.markdown("通过精细化拆解订单日期，实现基于具体日期、房型、真实可用房量的深入分析。")

# 1. 侧边栏：文件上传
uploaded_file = st.sidebar.file_uploader("📂 上传预约明细 CSV 文件", type="csv")

if uploaded_file is not None:
    # --- 数据加载与预处理 ---
    df = pd.read_csv(uploaded_file)
    
    # 过滤有效预约（假设状态为"预约成功"，根据你的实际数据调整）
    # 如果要看所有状态，可以注释掉这行，但算OCC通常只看有效单
    if '预约状态' in df.columns:
        df_valid = df[df['预约状态'] == '预约成功'].copy()
    else:
        df_valid = df.copy()

    # 清洗日期格式
    df_valid['入住日期'] = pd.to_datetime(df_valid['入住日期'], errors='coerce')
    df_valid['离店日期'] = pd.to_datetime(df_valid['离店日期'], errors='coerce')
    
    # 从“间晚”字段提取房间数量 (例如 "1间2晚" 提取 "1")
    def extract_rooms(text):
        if pd.isna(text): return 1
        match = re.search(r'(\d+)间', str(text))
        return int(match.group(1)) if match else 1

    df_valid['房间数'] = df_valid['间晚'].apply(extract_rooms)

    # --- 用户自定义区域：房型库存配置 & 套餐简称映射 ---
    st.sidebar.header("⚙️ 运营参数配置")
    st.sidebar.markdown("请在下方表格中补充**各类房型总数量**及**套餐简称**，以计算精准的OCC和简化图表。")
    
    # A. 房型库存配置
    st.sidebar.subheader("1. 房型库存设置")
    unique_rooms = df_valid['预约房型'].dropna().unique()
    room_inventory_df = pd.DataFrame({
        "预约房型": unique_rooms,
        "每日总房量": [10] * len(unique_rooms) # 默认给10间，供用户修改
    })
    # 使用 st.data_editor 允许用户在网页上直接修改房量
    edited_room_inventory = st.sidebar.data_editor(
        room_inventory_df, 
        column_config={"每日总房量": st.column_config.NumberColumn(min_value=1)},
        hide_index=True, key="room_editor"
    )
    total_hotel_capacity = edited_room_inventory['每日总房量'].sum()

    # B. 套餐(商品ID)简称配置
    st.sidebar.subheader("2. 套餐(商品ID)简称设置")
    if '商品ID' in df_valid.columns and '商品名称' in df_valid.columns:
        unique_products = df_valid.drop_duplicates(subset=['商品ID'])[['商品ID', '商品名称']].copy()
        # 默认截取前8个字符作为简称
        unique_products['套餐简称'] = unique_products['商品名称'].apply(lambda x: str(x)[:8] + '...')
        edited_products = st.sidebar.data_editor(
            unique_products,
            disabled=["商品ID", "商品名称"], # 禁止修改原ID和名称
            hide_index=True, key="product_editor"
        )
        # 将简称映射回主数据
        product_map = dict(zip(edited_products['商品ID'], edited_products['套餐简称']))
        df_valid['套餐简称'] = df_valid['商品ID'].map(product_map)
    else:
        df_valid['套餐简称'] = df_valid['商品名称']

    # --- 核心逻辑：将日期段“爆炸”展开为按天计算 ---
    # 例如：入住23号，离店25号，拆解为23号和24号占用房间
    expanded_rows = []
    for _, row in df_valid.iterrows():
        start = row['入住日期']
        end = row['离店日期']
        if pd.notna(start) and pd.notna(end) and start < end:
            # 生成入住期间的每一天
            dates = pd.date_range(start, end - pd.Timedelta(days=1))
            for d in dates:
                new_row = row.copy()
                new_row['占用日期'] = d
                expanded_rows.append(new_row)
    
    if not expanded_rows:
        st.warning("有效入住日期数据不足，无法生成分析图表。")
        st.stop()

    df_daily = pd.DataFrame(expanded_rows)
    df_daily['星期'] = df_daily['占用日期'].dt.day_name().map({
        'Monday': '周一', 'Tuesday': '周二', 'Wednesday': '周三',
        'Thursday': '周四', 'Friday': '周五', 'Saturday': '周六', 'Sunday': '周日'
    })
    
    # --- 界面展示板块 ---

    st.divider()
    
    # 板块 1: 入住 OCC 深度分析 (具体到日期与房量)
    st.header("📈 日常 OCC (入住率) 分析")
    
    # 每日总占用房数
    daily_occ = df_daily.groupby('占用日期')['房间数'].sum().reset_index()
    daily_occ['总可用房量'] = total_hotel_capacity
    daily_occ['OCC(%)'] = (daily_occ['房间数'] / daily_occ['总可用房量'] * 100).round(2)

    fig_occ = px.line(daily_occ, x='占用日期', y='OCC(%)', markers=True, text='OCC(%)',
                      title=f"全酒店每日总入住率 (基于配置的总房量: {total_hotel_capacity}间)")
    fig_occ.update_traces(textposition="top center")
    st.plotly_chart(fig_occ, use_container_width=True)

    # OCC 文字总结
    max_occ_row = daily_occ.loc[daily_occ['OCC(%)'].idxmax()]
    avg_occ = daily_occ['OCC(%)'].mean()
    st.info(f"💡 **OCC 分析总结：** 选定期间内，酒店平均入住率为 **{avg_occ:.1f}%**。入住率最高峰出现在 **{max_occ_row['占用日期'].strftime('%Y-%m-%d')}**，当天OCC达到 **{max_occ_row['OCC(%)']}%**（售出 {max_occ_row['房间数']} 间）。请重点关注低谷日期的营销动作。")

    st.divider()

    # 板块 2: 房型与具体房量分析
    st.header("🛏️ 房型售卖占比分析")
    col_room1, col_room2 = st.columns(2)
    
    room_sales = df_daily.groupby('预约房型')['房间数'].sum().reset_index()
    
    with col_room1:
        fig_room_pie = px.pie(room_sales, names='预约房型', values='房间数', hole=0.4, title="各房型售卖间夜量占比")
        fig_room_pie.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_room_pie, use_container_width=True)

    with col_room2:
        # 结合用户配置的房量看各房型的去化率
        room_merged = pd.merge(room_sales, edited_room_inventory, on='预约房型')
        # 这里的房型去化率简单计算为：总售出间夜 / (分析天数 * 该房型总数量)
        analysis_days = df_daily['占用日期'].nunique() if not df_daily.empty else 1
        room_merged['预期总间夜'] = room_merged['每日总房量'] * analysis_days
        room_merged['综合去化率(%)'] = (room_merged['房间数'] / room_merged['预期总间夜'] * 100).round(1)
        
        fig_room_bar = px.bar(room_merged, x='综合去化率(%)', y='预约房型', orientation='h', text='综合去化率(%)', title="各房型综合去化率")
        fig_room_bar.update_traces(texttemplate='%{text}%', textposition='outside')
        st.plotly_chart(fig_room_bar, use_container_width=True)

    # 房型文字总结
    top_room = room_sales.loc[room_sales['房间数'].idxmax()]
    st.info(f"💡 **房型分析总结：** 卖得最好的房型是 **{top_room['预约房型']}**，共售出 {top_room['房间数']} 个间夜。结合去化率图表，若某房型去化率远低于平均值，建议在接下来的套餐设计中增加该房型的搭售或升级优惠。")

    st.divider()

    # 板块 3: 星期与节假日规律分析
    st.header("📅 入住星期规律分析")
    weekday_order = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
    weekday_sales = df_daily.groupby('星期')['房间数'].sum().reindex(weekday_order).reset_index().dropna()
    
    fig_week = px.bar(weekday_sales, x='星期', y='房间数', text='房间数', title="一周各天入住间夜量分布", color='房间数', color_continuous_scale='Blues')
    fig_week.update_traces(textposition='outside')
    st.plotly_chart(fig_week, use_container_width=True)

    # 星期总结
    top_weekday = weekday_sales.loc[weekday_sales['房间数'].idxmax()]
    bottom_weekday = weekday_sales.loc[weekday_sales['房间数'].idxmin()]
    st.info(f"💡 **入住时机总结：** 订单主要集中在 **{top_weekday['星期']}**（共 {top_weekday['房间数']} 间夜），而 **{bottom_weekday['星期']}** 最为冷清。建议针对低谷星期推出“连住优惠”或“错峰体验活动”。")

    st.divider()

    # 板块 4: 套餐 (商品ID) 分析
    st.header("🎁 套餐(商品ID) 转化分析")
    package_sales = df_valid.groupby('套餐简称').agg(
        订单量=('订单ID', 'count'),
        实付总额=('用户实付金额', 'sum')
    ).reset_index().sort_values(by='订单量', ascending=False)

    col_pkg1, col_pkg2 = st.columns(2)
    with col_pkg1:
        fig_pkg_bar = px.bar(package_sales, x='订单量', y='套餐简称', orientation='h', text='订单量', title="各套餐(简称) 售出单量")
        fig_pkg_bar.update_traces(textposition='outside')
        st.plotly_chart(fig_pkg_bar, use_container_width=True)
    with col_pkg2:
        fig_pkg_rev = px.pie(package_sales, names='套餐简称', values='实付总额', hole=0.3, title="各套餐(简称) 营收占比")
        fig_pkg_rev.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_pkg_rev, use_container_width=True)

    # 套餐总结
    top_pkg = package_sales.iloc[0]
    st.info(f"💡 **套餐分析总结：** 按照商品ID归类，爆款套餐为 **{top_pkg['套餐简称']}**，共售出 {top_pkg['订单量']} 单，是营收的主要贡献点。建议保留其核心卖点（如包含的特定游乐项目/餐饮），并复制到冷门房型的套餐中。")

else:
    st.info("💡 请在左侧侧边栏上传 CSV 文件，然后配置房型库存和套餐简称，即可查看深度分析。")