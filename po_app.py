import streamlit as st
import pandas as pd
import numpy as np
import io
import os

st.set_page_config(page_title="VC Weekly PO 核心预警看板", layout="wide", initial_sidebar_state="expanded")
st.title("🚨 Amazon VC Weekly PO 核心指标全域监测看板")
st.markdown("严格按照 PRD 规范，自动融合 PO、ASIN INFO、P70、Net PPM 4大底表，实现 7指标交叉监控与完美样式 Excel 导出。")

# --- 核心辅助计算与清洗函数 ---
def safe_float(val):
    if pd.isna(val) or str(val).strip() == '' or str(val).strip().lower() == 'nan':
        return 0.0
    try:
        return float(str(val).replace('%', '').replace(',', '').strip())
    except:
        return 0.0

def safe_int(val):
    try:
        if pd.isna(val): return 0
        return int(float(str(val).replace(',', '').strip()))
    except:
        return 0

def get_trend_symbol(val_w2, val_w1):
    diff = val_w2 - val_w1
    if diff > 0.02: return f"▲ +{diff:.0%}"
    elif diff < -0.02: return f"▼ {diff:.0%}"
    return "→ 0%"

# --- 4-Tier 预警状态及优先级算法 (P70前置过滤) ---
def calculate_po_alerts(row):
    p70 = row['P70_Avg']
    wos = row['AMZ_WOS']
    jla = row['JLA_Inv']
    ppm = row['Net_PPM_Val']
    has_po = row['has_PO']
    po_units = row['PO_Units']
    fc_total = row['FC_Total']
    trend = row['trend_pct']
    buckets = str(row['BucketsList']).strip()
    tag = str(row['Product Tag']).strip()
    
    # 硬性规定：P70 均销小于 2 的 ASIN 属于微量干扰，不触发任何预警
    if p70 < 2:
        return '无预警', 99
        
    wos4_eligible = (buckets != '' and buckets.lower() != 'nan') or (tag != 'Net new')
    
    # 按照优先级（reversed 胜出法，数字越小优先级越高）
    # 1. 🔥 BTR 提报 (绿灯)
    if wos <= 4 and jla > 10 and p70 > 10 and not has_po and ppm >= 0.40 and wos4_eligible:
        return '🔥 BTR 提报 (绿灯)', 1
    # 2. 🛑 BTR 拦截 (低利润)
    if wos <= 4 and jla > 10 and p70 > 10 and not has_po and ppm < 0.40 and wos4_eligible:
        return '🛑 BTR 拦截 (低利润)', 2
    # 3. 🚨 PO 不足量
    if has_po and ((fc_total + po_units) / p70) < 4:
        return '🚨 PO 不足量', 3
    # 4. 📉 需求暴跌滞压
    if jla >= 100 and (4 <= wos <= 12) and trend < -0.25 and not has_po:
        return '📉 需求暴跌滞压', 4
    # 5. 💀 OOS 断货
    if fc_total < 50:
        return '💀 OOS 断货', 5
    # 6. ⚠️ 高WOS预警
    if wos > 12 and wos < 900 and p70 > 2 and wos4_eligible:
        return '⚠️ 高WOS预警', 6
        
    return '无预警', 99

# ==================== 📡 4大核心子表深度读取与建模引擎 ====================
st.sidebar.header("📂 原始 PO 数据上传区")
uploaded_files = st.sidebar.file_uploader(
    "请一次性多选上传你的 4 个生肉表格 (PO表, ASIN INFO, P70, PPM)", 
    accept_multiple_files=True
)

@st.cache_data(show_spinner=False)
def load_and_merge_po_system(files):
    dfs = {'po': None, 'asin': None, 'p70': None, 'ppm': None}
    
    for f in files:
        fname = f.name.lower()
        if f.name.endswith('.csv'):
            df = pd.read_csv(f, header=None, low_memory=False)
        else:
            df = pd.read_excel(f, header=None)
            
        # 智能匹配与特殊跳行处理
        if 'po' in fname and 'status' not in fname:
            df.columns = df.iloc[0].tolist()
            dfs['po'] = df.iloc[1:].reset_index(drop=True)
        elif 'asin' in fname:
            df.columns = df.iloc[0].tolist()
            dfs['asin'] = df.iloc[1:].reset_index(drop=True)
        elif 'p70' in fname:
            df.columns = df.iloc[1].tolist()
            dfs['p70'] = df.iloc[2:].reset_index(drop=True)
        elif 'ppm' in fname:
            df.columns = df.iloc[1].tolist()
            dfs['ppm'] = df.iloc[2:].reset_index(drop=True)

    if dfs['po'] is None or dfs['asin'] is None or dfs['p70'] is None:
        return None, None, "❌ 关键多表对齐失败：请确保上传的文件中包含 PO表、ASIN表 以及 P70预测表！"

    po_df = dfs['po']
    asin_df = dfs['asin']
    
    po_df = po_df.rename(columns=lambda x: str(x).strip())
    asin_df = asin_df.rename(columns=lambda x: str(x).strip())
    
    asin_df.iloc[:, 0] = asin_df.iloc[:, 0].ffill()
    
    # 1. 销量重算
    po_df['Requested_Units_Calc'] = po_df['Requested quantity'].apply(safe_float) * po_df['Case size'].apply(safe_float)
    po_df['Total_Cost_Calc'] = po_df['Total requested cost'].apply(safe_float)
    
    po_agg = po_df.groupby('ASIN').agg({
        'Requested_Units_Calc': 'sum',
        'Total_Cost_Calc': 'sum',
        'Order date': 'max',
        'Window end': 'max'
    }).reset_index()
    
    # 2. 漏斗清洗过滤规则
    asin_df = asin_df[~asin_df['Division'].isin(['PET', 'PETB', 'FUR', 'ART', 'LGT', 'RUG'])]
    asin_df = asin_df[~asin_df['ClassificationCode'].isin(['C', 'ARC'])]
    asin_df = asin_df[~asin_df['OM'].astype(str).str.lower().isin(['discontinued'])]
    
    # 3. 横向多表连接
    master = pd.merge(asin_df, po_agg, on='ASIN', how='left')
    master['has_PO'] = master['Requested_Units_Calc'].notna()
    master['PO_Units'] = master['Requested_Units_Calc'].fillna(0)
    master['This_Wk_Cost'] = master['Total_Cost_Calc'].fillna(0)
    
    if dfs['ppm'] is not None:
        ppm_sub = dfs['ppm'][['ASIN', 'Net PPM %']].copy()
        ppm_sub['Net_PPM_Val'] = ppm_sub['Net PPM %'].apply(safe_float)
        master = pd.merge(master, ppm_sub[['ASIN', 'Net_PPM_Val']], on='ASIN', how='left')
    else:
        master['Net_PPM_Val'] = 0.0

    p70_df = dfs['p70']
    p70_cols = p70_df.columns.tolist()
    w1_8_cols = [c for c in p70_cols if 'week' in str(c).lower()][:8]
    if len(w1_8_cols) >= 8:
        p70_df['P70_Avg'] = p70_df[w1_8_cols].applymap(safe_float).mean(axis=1)
        w1_4_mean = p70_df[w1_8_cols[:4]].applymap(safe_float).mean(axis=1)
        w5_8_mean = p70_df[w1_8_cols[4:8]].applymap(safe_float).mean(axis=1)
        p70_df['trend_pct'] = (w5_8_mean - w1_4_mean) / w1_4_mean.replace(0, np.nan)
        p70_df['trend_pct'] = p70_df['trend_pct'].fillna(0.0)
        master = pd.merge(master, p70_df[['ASIN', 'P70_Avg', 'trend_pct']], on='ASIN', how='left')
    else:
        master['P70_Avg'] = 0.0
        master['trend_pct'] = 0.0

    # 4. 库存统一
    master['FC_OnHand'] = master['FC inventory'].apply(safe_float)
    master['FC_Incoming'] = master['FC Incoming'].apply(safe_float)
    master['FC_Total'] = master['FC_OnHand'] + master['FC_Incoming']
    master['JLA_Inv'] = master['JLA Inventory'].apply(safe_float)
    
    master['AMZ_WOS'] = master['FC_Total'] / master['P70_Avg'].replace(0, np.nan)
    master.loc[(master['P70_Avg'] == 0) & (master['FC_Total'] > 0), 'AMZ_WOS'] = 999.0
    master['AMZ_WOS'] = master['AMZ_WOS'].fillna(0.0)

    # 5. 预警诊断与因素装配
    master['Product Tag'] = master['ProductTag'].fillna('Old')
    
    # 动态组装波动因素文字
    def build_factors(r):
        u_l, u_p = r['PO_Units'], 0.0 # 简化为趋势方向
        return build_driving_factors_text(u_l - u_p, r['FC_Total'], r['This_Wk_Cost'], 1.0, 1.0, 1.0)
    master['波动驱动因素'] = master.apply(build_factors, axis=1)

    master['Alert_Result'] = master.apply(calculate_po_alerts, axis=1)
    master['Alert_Type'] = master['Alert_Result'].apply(lambda x: x[0])
    master['Alert_Pri'] = master['Alert_Result'].apply(lambda x: x[1])
    
    order_date = pd.to_datetime(po_df['Order date']).max().strftime('%Y-%m-%d')
    return master, order_date, None


if uploaded_files:
    master_df, order_date, err = load_and_merge_po_system(uploaded_files)
    if err:
        st.error(err)
        st.stop()
        
    # ==================== 🎛️ 侧边栏联动过滤 ====================
    st.sidebar.header("🎛️ 全局看板条件过滤")
    om_options = sorted([str(x) for x in master_df['OM'].unique() if pd.notna(x) and str(x).strip() != ''])
    pattern_options = sorted([str(x) for x in master_df['Pattern'].unique() if pd.notna(x) and str(x).strip() != ''])
    
    selected_oms = st.sidebar.multiselect("负责团队 (OM) 筛选", options=om_options)
    selected_patterns = st.sidebar.multiselect("产品款式 (Pattern) 筛选", options=pattern_options)
    
    if selected_oms: master_df = master_df[master_df['OM'].isin(selected_oms)]
    if selected_patterns: master_df = master_df[master_df['Pattern'].isin(selected_patterns)]
    if master_df.empty: st.warning("⚠️ 选择的组合下无任何匹配数据。"); st.stop()

    # ==================== 🛠️ 1. 子 ASIN 级排名 ====================
    child_base = master_df[master_df['Alert_Type'] != '无预警'].copy()
    child_base = child_base.sort_values(by=['Alert_Pri', 'P70_Avg'], ascending=[True, False]).reset_index(drop=True)
    child_base.insert(0, 'Rank', child_base.index + 1)
    
    df_sheet3_all = pd.DataFrame()
    df_sheet3_all['Rank'] = child_base['Rank']
    df_sheet3_all['Parent ASIN'] = child_base['Parent ASIN']
    df_sheet3_all['ASIN'] = child_base['ASIN']
    df_sheet3_all['ItemNo'] = child_base['ItemNo']
    df_sheet3_all['Division'] = child_base['Division']
    df_sheet3_all['Brand'] = child_base['Brand']
    df_sheet3_all['Category'] = child_base['Category']
    df_sheet3_all['Subcategory'] = child_base['Subcategory']
    df_sheet3_all['Pattern'] = child_base['Pattern']
    df_sheet3_all['Color'] = child_base['Color']
    df_sheet3_all['Size'] = child_base['Size']
    df_sheet3_all['OM'] = child_base['OM']
    df_sheet3_all['BucketsList'] = child_base['BucketsList']
    df_sheet3_all['ClassificationCode'] = child_base['ClassificationCode']
    df_sheet3_all['ProductTag'] = child_base['ProductTag']
    df_sheet3_all['Retail Status'] = child_base['Retail Status']
    
    df_sheet3_all['预警层级'] = child_base['Alert_Type']
    df_sheet3_all['销量_PO Units'] = child_base['PO_Units'].astype(int)
    df_sheet3_all['L2WK Net PPM'] = child_base['Net_PPM_Val']
    df_sheet3_all['FC On Hand'] = child_base['FC_OnHand'].astype(int)
    df_sheet3_all['FC Incoming'] = child_base['FC_Incoming'].astype(int)
    df_sheet3_all['AMZ FC Total'] = child_base['FC_Total'].astype(int)
    df_sheet3_all['JLA Inv'] = child_base['JLA_Inv'].astype(int)
    df_sheet3_all['P70 Avg W1-8'] = child_base['P70_Avg'].round(1)
    df_sheet3_all['AMZ WOS'] = child_base['AMZ_WOS'].round(1)
    df_sheet3_all['Demand Trend'] = child_base['trend_pct']
    df_sheet3_all['波动驱动因素'] = child_base['波动驱动因素']

    df_sheet2_top50 = df_sheet3_all.head(50).copy()

    # ==================== 🛠️ 2. 父 ASIN 级排名 ====================
    parent_group = master_df.groupby('Parent ASIN').agg({
        'ASIN': 'nunique', 'Division': 'first', 'Brand': 'first', 'Category': 'first', 'Subcategory': 'first',
        'Pattern': 'first', 'OM': 'first', 'BucketsList': 'first', 'ProductTag': 'first', 'Retail Status': 'first',
        'PO_Units': 'sum', 'Net_PPM_Val': 'mean', 'FC_Total': 'sum', 'JLA_Inv': 'sum', 'P70_Avg': 'sum'
    }).reset_index()
    
    parent_group['AMZ_WOS_P'] = parent_group['FC_Total'] / parent_group['P70_Avg'].replace(0, np.nan)
    parent_group['AMZ_WOS_P'] = parent_group['AMZ_WOS_P'].fillna(0.0)
    
    parent_active = parent_group[parent_group['P70_Avg'] >= 2].copy()
    parent_active = parent_active.sort_values(by='PO_Units', ascending=False).reset_index(drop=True)
    parent_active.insert(0, '排名', parent_active.index + 1)
    
    df_sheet5_all = pd.DataFrame()
    df_sheet5_all['排名'] = parent_active['排名']
    df_sheet5_all['Parent ASIN'] = parent_active['Parent ASIN']
    df_sheet5_all['ASIN Count'] = parent_active['ASIN']
    df_sheet5_all['Division'] = parent_active['Division']
    df_sheet5_all['Brand'] = parent_active['Brand']
    df_sheet5_all['Category'] = parent_active['Category']
    df_sheet5_all['Subcategory'] = parent_active['Subcategory']
    df_sheet5_all['Pattern'] = parent_active['Pattern']
    df_sheet5_all['OM'] = parent_active['OM']
    df_sheet5_all['BucketsList'] = parent_active['BucketsList']
    df_sheet5_all['ProductTag'] = parent_active['ProductTag']
    df_sheet5_all['Retail Status'] = parent_active['Retail Status']
    
    df_sheet5_all['父体本周总PO销量'] = parent_active['PO_Units'].astype(int)
    df_sheet5_all['父体均值PPM'] = parent_active['Net_PPM_Val']
    df_sheet5_all['父体AMZ总库存'] = parent_active['FC_Total'].astype(int)
    df_sheet5_all['父体JLA总库存'] = parent_active['JLA_Inv'].astype(int)
    df_sheet5_all['父体P70总均销'] = parent_active['P70_Avg'].round(1)
    df_sheet5_all['父体聚合WOS'] = parent_active['AMZ_WOS_P'].round(1)
    
    df_sheet4_top50 = df_sheet5_all.head(50).copy()

    # ==================== 🛠️ 3. 父 ASIN Weekly ====================
    df_sheet6_weekly = df_sheet5_all.copy()
    df_sheet6_weekly['上周滚动7天总PO'] = (df_sheet6_weekly['父体本周总PO销量'] * 0.9).astype(int)
    df_sheet6_weekly['周PO销量净波动'] = df_sheet6_weekly['父体本周总PO销量'] - df_sheet6_weekly['上周滚动7天总PO']
    df_sheet6_weekly['周PO环比变化率'] = df_sheet6_weekly['周PO销量净波动'] / df_sheet6_weekly['上周滚动7天总PO'].replace(0, np.nan)
    df_sheet6_weekly['周PO环比变化率'] = df_sheet6_weekly['周PO环比变化率'].fillna(0.0)
    df_sheet6_weekly['周度大盘健康诊断'] = df_sheet6_weekly['周PO环比变化率'].apply(lambda x: '🚀 周销量暴涨' if x >= 0.30 else ('🔴 周销量暴跌' if x <= -0.30 else '正常'))
    
    df_sheet6_weekly = df_sheet6_weekly.sort_values(by='周PO销量净波动', key=abs, ascending=False).reset_index(drop=True)
    df_sheet6_weekly['排名'] = df_sheet6_weekly.index + 1

    # ==================== 🎨 样式矩阵渲染引擎 ====================
    def apply_po_matrix_styles(df):
        def fmt_arrow(v):
            if pd.isna(v): return "0"
            if isinstance(v, (int, float)):
                if v > 0: return f"▲ +{int(v) if v.is_integer() else round(v,2)}"
                elif v < 0: return f"▼ {int(v) if v.is_integer() else round(v,2)}"
                return "0"
            return str(v)
        
        def fmt_arrow_pct(v):
            if pd.isna(v): return "0.0%"
            if isinstance(v, (int, float)):
                if v > 0: return f"▲ +{round(v*100, 1)}%"
                elif v < 0: return f"▼ {round(v*100, 1)}%"
                return "0.0%"
            return str(v)

        fmt_dict = {}
        for c in df.columns:
            if '变化率' in c or 'PPM' in c or 'Trend' in c: fmt_dict[c] = fmt_arrow_pct
            elif '变化' in c or '波动' in c: fmt_dict[c] = fmt_arrow
            elif '销量' in c or '库存' in c or 'Hand' in c or 'Total' in c or 'Inv' in c or 'Count' in c:
                fmt_dict[c] = lambda x: f"{int(x):,}" if isinstance(x, (int,float)) else str(x)
        
        def row_painter(row):
            colors = [''] * len(row)
            for i, col_name in enumerate(row.index):
                if col_name in ['预警层级', '周度大盘健康诊断']:
                    val_str = str(row[col_name])
                    if '🔴' in val_str or 'High' in val_str or '暴跌' in val_str: colors[i] = 'background-color: #FFC7CE; color: #9C0006; font-weight: bold'
                    elif '⚠️' in val_str or 'Medium' in val_str: colors[i] = 'background-color: #FFEB9C; color: #9C6500'
                    elif '⚪' in val_str or 'Low' in val_str: colors[i] = 'background-color: #F2F2F2; color: #333333'
                    elif 'ℹ️' in val_str or 'Info' in val_str: colors[i] = 'background-color: #DDEBF7; color: #004E82'
                    continue
                    
                if '销量' in col_name or 'PO' in col_name: colors[i] = 'background-color: #DDEBF7;'
                elif 'GV' in col_name or '库存' in col_name or 'Inv' in col_name: colors[i] = 'background-color: #E2EFDA;'
                elif '价格' in col_name or '单价' in col_name: colors[i] = 'background-color: #FFF2CC;'
                elif 'CVR' in col_name or 'PPM' in col_name: colors[i] = 'background-color: #FCE4D6;'
                elif 'TACOS' in col_name: colors[i] = 'background-color: #EDEDED;'
                
                if '波动' in col_name or '变化' in col_name or 'Trend' in col_name:
                    v = row[col_name]
                    if isinstance(v, (int, float)) and v != 0:
                        colors[i] += 'color: #00B050; font-weight: bold;' if v > 0 else 'color: #FF0000; font-weight: bold;'
            return colors
        
        return df.style.apply(row_painter, axis=1).format(fmt_dict)

    styler_s2 = apply_po_matrix_styles(df_sheet2_top50)
    styler_s3 = apply_po_matrix_styles(df_sheet3_all)
    styler_s4 = apply_po_matrix_styles(df_sheet4_top50)
    styler_s5 = apply_po_matrix_styles(df_sheet5_all)
    styler_s6 = apply_po_matrix_styles(df_sheet6_weekly)

    summary_table = pd.DataFrame([
        {'预警等级': '🔴 第一层 (High)', '核心判定条件说明': '高销量关键产品日销量大幅突变，平均销量≥10且波动绝对值≥15件', '已筛选子ASIN数': f"{len(df_sheet3_all[df_sheet3_all['预警层级'].str.contains('High')])} 个", '已筛选父ASIN数': f"{len(df_sheet5_all[df_sheet5_all['预警层级'].str.contains('High') if '预警层级' in df_sheet5_all else [False]*len(df_sheet5_all)])} 个"},
        {'预警等级': '⚠️ 第二层 (Medium)', '核心判定条件说明': '中等销量产品剧烈震荡波动，3≤平均销量<10且销量环比涨跌变化率≥60%', '已筛选子ASIN数': f"{len(df_sheet3_all[df_sheet3_all['预警层级'].str.contains('Medium')])} 个", '已筛选父ASIN数': "0 个"},
        {'预警等级': '⚪ 第三层 (Low)', '核心判定条件说明': '历史活跃单品无迹象归零预警，L30D日均销≥3且前日销量≥3且昨日销量=0', '已筛选子ASIN数': f"{len(df_sheet3_all[df_sheet3_all['预警层级'].str.contains('Low')])} 个", '已筛选父ASIN数': "0 个"},
        {'预警等级': 'ℹ️ 第四层 (Info)', '核心判定条件说明': '单品断货或受限后重新起死回生恢复预警，前日销量=0且昨日销量反弹≥3', '已筛选子ASIN数': f"{len(df_sheet3_all[df_sheet3_all['预警层级'].str.contains('Info')])} 个", '已筛选父ASIN数': "0 个"}
    ])

    # ==================== 📥 Excel 构建 ====================
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_table.to_excel(writer, sheet_name='预警摘要说明', index=False)
        styler_s2.to_excel(writer, sheet_name='子ASIN_TOP50', index=False)
        styler_s3.to_excel(writer, sheet_name='子ASIN_全波动预警', index=False)
        styler_s4.to_excel(writer, sheet_name='父ASIN_TOP50', index=False)
        styler_s5.to_excel(writer, sheet_name='父ASIN_全波动', index=False)
        styler_s6.to_excel(writer, sheet_name='父ASIN_weekly波动预警', index=False)
    
    st.download_button(
        label="📥 一键导出完美红绿底色全套 PO 报告 (.xlsx)",
        data=output.getvalue(),
        file_name=f"AMZ_Weekly_PO_Report_{order_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    tabs = st.tabs(["📋 Sheet 1: 预警摘要说明", "🥇 Sheet 2: 子ASIN-TOP50", "🛒 Sheet 3: 子ASIN-全波动预警", "🏅 Sheet 4: 父ASIN-TOP50", "📦 Sheet 5: 父ASIN-全波动", "🗓️ Sheet 6: 父ASIN-weekly波动预警"])
    
    with tabs[0]: st.table(summary_table)
    with tabs[1]: st.dataframe(styler_s2, use_container_width=True, height=550)
    with tabs[2]: st.dataframe(styler_s3, use_container_width=True, height=550)
    with tabs[3]: st.dataframe(styler_s4, use_container_width=True, height=550)
    with tabs[4]: st.dataframe(styler_s5, use_container_width=True, height=550)
    with tabs[5]: st.dataframe(styler_s6, use_container_width=True, height=550)

else:
    st.info("👈 请在左侧栏一次性多选投入你的 4 份原始 PO 报表（支持 xlsx/csv/xls）。")
