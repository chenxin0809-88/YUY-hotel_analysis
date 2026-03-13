import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re

# 设置页面
st.set_page_config(page_title="酒店经营精细化看板", layout="wide")

st.title("🏨 酒店预约明细与 OCC 精细化分析系统")
st.markdown("支持上传 Excel/CSV，自定义房量库存，自动拆解每日入住情况。")

# 1. 文件上传与预处理
st.sidebar.header("📁 数据导入")
uploaded_file = st.sidebar.file_uploader("上传预约明细 (Excel 或 CSV)", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 判断格式读取
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # --- 数据清洗 ---
        # 1. 清除列名空格
        df.columns = [str(c).strip() for c in df.columns]
        
        # 2. 检查关键列是否存在
        required_cols = ['预约房型', '入住日期', '离店日期', '预约状态', '商品ID', '商品名称']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"表格缺少必要列: {missing}")
            st.info(f"当前表格列名: {list(df.columns)}")
            st.stop()

        # 3. 日期格式化
        df['入住日期'] = pd.to_datetime(df['入住日期'], errors='coerce')
        df['离店日期'] = pd.to_datetime(df['离店日期'], errors='coerce')
        df = df.dropna(subset=['入住日期', '离店日期', '预约房型'])

        # 4. 提取房间数
        def get_rooms(x):
            if pd.isna(x): return 1
            res = re.search(r'(\d+)间', str(x))
            return int(res.group(1)) if res else 1
        df['房间数'] = df['间晚'].apply(get_rooms) if '间晚' in df.columns else 1

        # --- 侧边栏配置 ---
        st.sidebar.divider()
        st.sidebar.header("⚙️ 运营参数配置")
        
        # 房型库存配置
        st.sidebar.subheader("1. 房型库存设置")
        unique_rooms = df['预约房型'].unique()
        inventory_init = pd.DataFrame({"预约房型": unique_rooms, "每日库存": [10]*len(unique_rooms)})
        edited_inventory = st.sidebar.data_editor(inventory_init, hide_index=True)
        inventory_dict = dict(zip(edited_inventory['预约房型'], edited_inventory['每日库存']))
        total_capacity = edited_inventory['每日库存'].sum()

        # 套餐简称配置 (按商品ID)
        st.sidebar.subheader("2. 套餐简称设置")
        unique_pkgs = df[['商品ID', '商品名称']].drop_duplicates()
        unique_pkgs['简称'] = unique_pkgs['商品名称'].apply(lambda x: str(x)[:10] + '...')
        edited_pkgs = st.sidebar.data_editor(unique_pkgs, hide_index=True, disabled=['商品ID', '商品名称'])
        pkg_map = dict(zip(edited_pkgs['商品ID'], edited_pkgs['简称']))
        df['套餐简称'] = df['商品ID'].map(pkg_map)

        # --- 核心计算：日期爆炸拆解 ---
        # 只统计“预约成功”的订单用于计算OCC
        df_active = df[df['预约状态'] == '预约成功'].copy()
        expanded_data = []
        for _, row in df_active.iterrows():
            days = (row['离店日期'] - row['入住日期']).days
            if days > 0:
                for i in range(days):
                    curr_date = row['入住日期'] + pd.Timedelta(days=i)
                    expanded_data.append({
                        '日期': curr_date,
                        '房型': row['预约房型'],
                        '房间数': row['房间数'],
                        '星期': curr_date.day_name(),
                        '套餐': row['套餐简称']
                    })
        
        if not expanded_data:
            st.warning("暂无有效的预约成功订单。")
            st.stop()
            
        df_daily = pd.DataFrame(expanded_data)
        
        # --- 报表呈现 ---
        
        # 看板核心指标
        st.header("📊 核心经营指标")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        total_rev = df['用户实付金额'].sum()
        total_nights = df_daily['房间数'].sum()
        avg_occ = (total_nights / (df_daily['日期'].nunique() * total_capacity) * 100) if total_capacity > 0 else 0
        
        kpi1.metric("总销售额", f"¥{total_rev:,.2f}")
        kpi2.metric("总预约成功间夜", f"{total_nights} 间夜")
        kpi3.metric("周期内平均OCC", f"{avg_occ:.1f}%")
        kpi4.metric("覆盖日期天数", f"{df_daily['日期'].nunique()} 天")

        st.divider()

        # 图表展示
        tab1, tab2, tab3 = st.tabs(["📉 每日OCC与房型分析", "📅 周期/星期规律", "🎁 套餐转化分析"])

        with tab1:
            st.subheader("每日入住率(OCC)趋势")
            occ_trend = df_daily.groupby('日期')['房间数'].sum().reset_index()
            occ_trend['OCC%'] = (occ_trend['房间数'] / total_capacity * 100).round(1)
            fig_occ = px.line(occ_trend, x='日期', y='OCC%', text='OCC%', markers=True, 
                             title=f"全店每日OCC趋势 (总库存:{total_capacity})")
            fig_occ.update_traces(textposition="top center")
            st.plotly_chart(fig_occ, use_container_width=True)
            
            st.subheader("各房型去化情况")
            room_analysis = df_daily.groupby('房型')['房间数'].sum().reset_index()
            fig_room = px.pie(room_analysis, names='房型', values='房间数', hole=0.4, title="房型售卖占比")
            fig_room.update_traces(textinfo='percent+label')
            st.plotly_chart(fig_room, use_container_width=True)

        with tab2:
            st.subheader("星期分布规律")
            week_map = {'Monday':'周一','Tuesday':'周二','Wednesday':'周三','Thursday':'周四','Friday':'周五','Saturday':'周六','Sunday':'周日'}
            df_daily['星期中文'] = df_daily['星期'].map(week_map)
            week_order = ['周一','周二','周三','周四','周五','周六','周日']
            week_data = df_daily.groupby('星期中文')['房间数'].sum().reindex(week_order).reset_index()
            fig_week = px.bar(week_data, x='星期中文', y='房间数', color='房间数', title="一周内入住热度分布")
            st.plotly_chart(fig_week, use_container_width=True)
            
            st.info(f"💡 **分析提示：** 数据显示 **{week_data.loc[week_data['房间数'].idxmax(), '星期中文']}** 是入住高峰期。建议针对低谷日推出特惠。")

        with tab3:
            st.subheader("套餐(按商品ID归类)售卖排行")
            pkg_analysis = df.groupby('套餐简称').agg({'订单ID':'count', '用户实付金额':'sum'}).reset_index()
            pkg_analysis = pkg_analysis.sort_values('订单ID', ascending=False)
            
            col_a, col_b = st.columns(2)
            with col_a:
                fig_p1 = px.bar(pkg_analysis, x='订单ID', y='套餐简称', orientation='h', title="套餐单量榜")
                st.plotly_chart(fig_p1, use_container_width=True)
            with col_b:
                fig_p2 = px.pie(pkg_analysis, names='套餐简称', values='用户实付金额', title="套餐营收贡献")
                st.plotly_chart(fig_p2, use_container_width=True)

        # 数据明细
        with st.expander("查看清洗后的数据明细"):
            st.dataframe(df)

    except Exception as e:
        st.error(f"处理数据时出错: {e}")
        st.exception(e)

else:
    st.info("👋 请在左侧侧边栏上传您的酒店预约明细文件 (支持 .xlsx 或 .csv)")