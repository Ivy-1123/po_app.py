import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from datetime import datetime

st.set_page_config(page_title="VC Weekly PO 智能预警看板", layout="wide", initial_sidebar_state="expanded")
st.title("🚨 Amazon VC Weekly PO 核心指标监测系统")
st.markdown("本看板严格对齐 PO Skill 规范，全自动智能识别 4 大生肉表格，多维联动交叉计算，输出纯正 PO 数据立方体。")

# --- 核心辅助计算与清洗函数 ---
def safe_float(val):
    if pd.isna(val) or str(val).strip() == '' or str(val).strip().lower() == 'nan':
        return 0.0
    try:
        return float(str(val).replace('%', '').replace('$', '').replace(',', '').strip())
    except:
        return 0.0

def safe_int(val):
    try:
        if pd.isna(val): return 0
        return int(float(str(val).replace(',', '').strip()))
    except:
        return 0

# --- 🚨 Action Required 核心预警算法 (严格执行优先级的 reversed 胜出法) ---
def calculate_po_alerts_core(row):
    p70 = row['P70_Avg_Calc']
    wos = row['AMZ_WOS_Calc']
    jla = row['JLA_Inv_Calc']
    ppm = row['Net_PPM_Calc']
    has_po = row['has_PO_Calc']
    po_units = row['PO_Units_Calc']
    fc_total = row['FC_Total_Calc']
    trend = row['trend_pct_Calc']
    buckets = str(row['Final_Buckets']).strip()
    tag = str(row['Product Tag']).strip()
    
    # 铁律：P70 均销 < 2 的 ASIN 属于微量干扰，完全剔除不触发任何预警
    if p70 < 2:
        return '无预警', 99
        
    wos4_eligible = (buckets != '' and buckets.lower() != 'nan') or (tag != 'Net new')
    
    alert_label = '无预警'
    priority = 99
    
    # reversed 倒序循环：低优先级先执行，高优先级后执行并覆盖，确保高优先级(pri=1)绝不动摇
    alert_logic_list = [
        (wos > 12 and wos < 900 and p70 > 2 and wos4_eligible, '⚠️ 高WOS预警', 6),
        (fc_total < 50, '💀 OOS 断货', 5),
        (jla >= 100 and (4 <= wos <= 12) and trend < -0.25 and not has_po, '📉 需求暴跌滞压', 4),
        (has_po and ((fc_total + po_units) / p70) < 4, '🚨 PO 不足量', 3),
        (wos <= 4 and jla > 10 and p70 > 10 and not has_po and ppm < 0.40 and wos4_eligible, '🛑 BTR 拦截 (低利润)', 2),
        (wos <= 4 and jla > 10 and p70 > 10 and not has_po and ppm >= 0.40 and wos4_eligible, '🔥 BTR 提报 (绿灯)', 1)
    ]
    
    for cond, label, pri in alert_logic_list:
        if cond:
            alert_label = label
            priority = pri
            
    return alert_label, priority

# ==================== 📡 智能特征寻标指纹识别引擎 ====================
st.sidebar.header("📂 原始 PO 数据上传区")
uploaded_files = st.sidebar.file_uploader(
    "可同时上传本周4张表(PO、ASIN、P70、PPM)，支持上传上周PO激活WoW分析", 
    accept_multiple_files=True
)

try:
    if uploaded_files:
        @st.cache_data(show_spinner=False)
        def load_and_merge_po_system(files):
            dfs = {'po_this': None, 'po_prev': None, 'asin': None, 'p70': None, 'ppm': None}
            
            for f in files:
                fname = f.name.lower()
                try:
                    if fname.endswith('.csv'): test_df = pd.read_csv(f, nrows=5, header=None, low_memory=False)
                    else: test_df = pd.read_excel(f, nrows=5, header=None)
                except: continue
                    
                row_str_pool = []
                for r_idx in range(min(3, len(test_df))):
                    row_str_pool.extend([str(x).strip().lower() for x in test_df.iloc[r_idx].dropna().tolist()])
                all_headers_combined = " ".join(row_str_pool)

                if 'requested quantity' in all_headers_combined or 'total requested cost' in all_headers_combined:
                    if fname.endswith('.csv'): df = pd.read_csv(f, low_memory=False)
                    else: df = pd.read_excel(f)
                    df.columns = [str(c).strip() for c in df.columns]
                    if 'prev' in fname or 'old' in fname or 'last' in fname or dfs['po_this'] is not None:
                        dfs['po_prev'] = df
                    else:
                        dfs['po_this'] = df
                elif 'classificationcode' in all_headers_combined or 'producttag' in all_headers_combined or 'jla inventory' in all_headers_combined:
                    if fname.endswith('.csv'): df = pd.read_csv(f, low_memory=False)
                    else: df = pd.read_excel(f)
                    df.columns = [str(c).strip() for c in df.columns]
                    dfs['asin'] = df
                elif 'week 1' in all_headers_combined or 'week1' in all_headers_combined or 'forecast' in all_headers_combined:
                    if fname.endswith('.csv'): raw = pd.read_csv(f, header=None, low_memory=False)
                    else: raw = pd.read_excel(f, header=None)
                    raw.iloc[:, 0] = raw.iloc[:, 0].ffill()
                    header_row_idx = 1
                    for i in range(min(5, len(raw))):
                        if any('week' in str(x).lower() for x in raw.iloc[i].tolist()):
                            header_row_idx = i; break
                    df = raw.iloc[header_row_idx+1:].copy()
                    df.columns = [str(x).strip() for x in raw.iloc[header_row_idx].tolist()]
                    dfs['p70'] = df.reset_index(drop=True)
                elif 'net ppm %' in all_headers_combined or 'ppm' in all_headers_combined:
                    if fname.endswith('.csv'): raw = pd.read_csv(f, header=None, low_memory=False)
                    else: raw = pd.read_excel(f, header=None)
                    header_row_idx = 1
                    for i in range(min(5, len(raw))):
                        if any('ppm' in str(x).lower() for x in raw.iloc[i].tolist()):
                            header_row_idx = i; break
                    df = raw.iloc[header_row_idx+1:].copy()
                    df.columns = [str(x).strip() for x in raw.iloc[header_row_idx].tolist()]
                    dfs['ppm'] = df.reset_index(drop=True)

            if dfs['po_this'] is None or dfs['asin'] is None or dfs['p70'] is None:
                return None, None, None, "❌ 关键多表特征对齐失败：请确保上传了包含有效表头特征的 PO表、ASIN INFO 和 P70预测表！"

            po_df = dfs['po_this']
            asin_df = dfs['asin']
            asin_cols = [str(c).strip() for c in asin_df.columns]
            asin_cols_lower = [c.lower() for c in asin_cols]
            
            def get_col_fallback(kws, d_idx):
                for kw in kws:
                    for i, c in enumerate(asin_cols_lower):
                        if kw in c: return asin_cols[i]
                return asin_cols[d_idx] if d_idx < len(asin_cols) else asin_cols[0]

            col_parent = get_col_fallback(['parent'], 0)
            col_asin = get_col_fallback(['asin'], 1)
            col_itemno = get_col_fallback(['itemno', 'item_no'], 2)
            col_division = get_col_fallback(['division', 'div'], 3)
            col_brand = get_col_fallback(['brand'], 4)
            col_category = get_col_fallback(['category'], 5)
            col_subcat = get_col_fallback(['subcategory', 'sub_category'], 6)
            col_pattern = get_col_fallback(['pattern'], 7)
            col_color = get_col_fallback(['color'], 8)
            col_size = get_col_fallback(['size'], 9)
            col_om = get_col_fallback(['om'], 11)
            col_buckets = get_col_fallback(['buckets', 'bucket'], 12)
            col_class = get_col_fallback(['classification', 'classcode', 'code'], 13)
            col_tag = get_col_fallback(['producttag', 'tag'], 15)
            col_status = get_col_fallback(['status', 'retail'], 18)

            asin_df[col_parent] = asin_df[col_parent].ffill()
            
            # 统一洗出基础映射列 (加固安全层)
            asin_df['Final_Parent'] = asin_df[col_parent]
            asin_df['Final_ASIN'] = asin_df[col_asin]
            asin_df['Final_ItemNo'] = asin_df[col_itemno]
            asin_df['Final_Division'] = asin_df[col_division]
            asin_df['Final_Brand'] = asin_df[col_brand]
            asin_df['Final_Category'] = asin_df[col_category]
            asin_df['Final_Subcat'] = asin_df[col_subcat]
            asin_df['Final_Pattern'] = asin_df[col_pattern]
            asin_df['Final_Color'] = asin_df[col_color]
            asin_df['Final_Size'] = asin_df[col_size]
            asin_df['Final_OM'] = asin_df[col_om]
            asin_df['Final_Buckets'] = asin_df[col_buckets]
            asin_df['Final_Class'] = asin_df[col_class]
            asin_df['Final_Status'] = asin_df[col_status]
            asin_df['Product Tag'] = asin_df[col_tag].fillna('Old')

            # 过滤清洗规则
            asin_df = asin_df[~asin_df['Final_Division'].isin(['PET', 'PETB', 'FUR', 'ART', 'LGT', 'RUG'])]
            asin_df = asin_df[~asin_df['Final_Class'].isin(['C', 'ARC'])]
            asin_df = asin_df[~asin_df['Final_OM'].astype(str).str.lower().isin(['discontinued'])]
            
            # PO 基础销量重算与聚合
            po_df['Units_Calc'] = po_df['Requested quantity'].apply(safe_float) * po_df['Case size'].apply(safe_float)
            po_df['Cost_Calc'] = po_df['Total requested cost'].apply(safe_float)
            po_agg = po_df.groupby('ASIN').agg({'Units_Calc': 'sum', 'Cost_Calc': 'sum', 'Order date': 'max', 'Window end': 'max'}).reset_index()
            
            # 多表交叉融合
            master = pd.merge(asin_df, po_agg, left_on='Final_ASIN', right_on='ASIN', how='left')
            master['has_PO_Calc'] = master['Units_Calc'].notna()
            master['PO_Units_Calc'] = master['Units_Calc'].fillna(0)
            master['This_Wk_Cost_Calc'] = master['Cost_Calc'].fillna(0)
            
            # 注入 PPM
            if dfs['ppm'] is not None:
                ppm_col = [c for c in dfs['ppm'].columns if 'ppm' in str(c).lower()][0]
                ppm_sub = dfs['ppm'][['ASIN', ppm_col]].copy()
                ppm_sub['Net_PPM_Calc'] = ppm_sub[ppm_col].apply(safe_float)
                master = pd.merge(master, ppm_sub[['ASIN', 'Net_PPM_Calc']], left_on='Final_ASIN', right_on='ASIN', how='left', suffixes=('', '_ppm'))
            else:
                master['Net_PPM_Calc'] = 0.0

            # 注入 P70 预测
            p70_df = dfs['p70']
            w1_8_cols = [c for c in p70_df.columns if 'week' in str(c).lower()][:8]
            if len(w1_8_cols) >= 8:
                p70_df['P70_Avg_Calc'] = p70_df[w1_8_cols].map(safe_float).mean(axis=1)
                w1_4_mean = p70_df[w1_8_cols[:4]].map(safe_float).mean(axis=1)
                w5_8_mean = p70_df[w1_8_cols[4:8]].map(safe_float).mean(axis=1)
                p70_df['trend_pct_Calc'] = (w5_8_mean - w1_4_mean) / w1_4_mean.replace(0, np.nan)
                p70_df['trend_pct_Calc'] = p70_df['trend_pct_Calc'].fillna(0.0)
                master = pd.merge(master, p70_df[['ASIN', 'P70_Avg_Calc', 'trend_pct_Calc']], left_on='Final_ASIN', right_on='ASIN', how='left', suffixes=('', '_p70'))
            else:
                master['P70_Avg_Calc'] = 0.0
                master['trend_pct_Calc'] = 0.0

            # 库存重组
            inv_col = [c for c in master.columns if 'jla inventory' in str(c).lower() or 'jla_inv' in str(c).lower() or 'jla inventory' in str(c).lower()]
            fc_col = [c for c in master.columns if 'fc inventory' in str(c).lower() or 'fc_onhand' in str(c).lower()]
            fcin_col = [c for c in master.columns if 'fc incoming' in str(c).lower() or 'fc_incoming' in str(c).lower()]

            master['FC_OnHand_Calc'] = master[fc_col[0]].apply(safe_float) if fc_col else 0.0
            master['FC_Incoming_Calc'] = master[fcin_col[0]].apply(safe_float) if fcin_col else 0.0
            master['FC_Total_Calc'] = master['FC_OnHand_Calc'] + master['FC_Incoming_Calc']
            master['JLA_Inv_Calc'] = master[inv_col[0]].apply(safe_float) if inv_col else 0.0
            
            master['AMZ_WOS_Calc'] = master['FC_Total_Calc'] / master['P70_Avg_Calc'].replace(0, np.nan)
            master.loc[(master['P70_Avg_Calc'] == 0) & (master['FC_Total_Calc'] > 0), 'AMZ_WOS_Calc'] = 999.0
            master['AMZ_WOS_Calc'] = master['AMZ_WOS_Calc'].fillna(0.0)

            # 【重要修正】：逻辑顺序正确归位。此时所有 Final_ 基础列已经生成完毕，可以安全调用预警逻辑！
            master['Final_ProductTag'] = master['Product Tag']
            master['Alert_Result'] = master.apply(calculate_po_alerts_core, axis=1)
            master['Alert_Type'] = master['Alert_Result'].apply(lambda x: x[0])
            master['Alert_Pri'] = master['Alert_Result'].apply(lambda x: x[1])
            
            try: odate = pd.to_datetime(po_df['Order date']).max().strftime('%Y-%m-%d')
            except: odate = datetime.now().strftime('%Y-%m-%d')
            
            wow_data = None
            if dfs['po_prev'] is not None:
                p_po = dfs['po_prev']
                p_po['P_Units'] = p_po['Requested quantity'].apply(safe_float) * p_po['Case size'].apply(safe_float)
                p_po['P_Cost'] = p_po['Total requested cost'].apply(safe_float)
                p_agg = p_po.groupby('ASIN').agg({'P_Units': 'sum', 'P_Cost': 'sum'}).reset_index()
                wow_data = p_agg

            return master, odate, wow_data, None

        master_df, order_date, wow_data, err = load_and_merge_po_system(uploaded_files)
        if err:
            st.error(err)
            st.stop()
            
        # 侧边栏过滤
        st.sidebar.header("🎛️ 全局看板条件过滤")
        om_options = sorted([str(x) for x in master_df['Final_OM'].unique() if pd.notna(x)])
        pattern_options = sorted([str(x) for x in master_df['Final_Pattern'].unique() if pd.notna(x)])
        selected_oms = st.sidebar.multiselect("负责团队 (OM) 筛选", options=om_options)
        selected_patterns = st.sidebar.multiselect("产品款式 (Pattern) 筛选", options=pattern_options)
        
        if selected_oms: master_df = master_df[master_df['Final_OM'].isin(selected_oms)]
        if selected_patterns: master_df = master_df[master_df['Final_Pattern'].isin(selected_patterns)]
        if master_df.empty: st.warning("⚠️ 选择的筛选组合下没有匹配到任何数据。"); st.stop()

        # ==================== 🛠️ 纯正 PO 看板四大工作表深度组装 ====================
        
        # --- 1. 📊 Summary 工作表组装 ---
        tot_filtered_cost = master_df['This_Wk_Cost_Calc'].sum()
        # Section A
        sec_a = master_df.groupby('Product Tag').agg({'PO_Units_Calc': 'sum', 'This_Wk_Cost_Calc': 'sum'}).reset_index()
        sec_a.columns = ['Product Tag', 'Requested Units', 'Requested Cost (USD)']
        sec_a['% of Total Cost'] = sec_a['Requested Cost (USD)'] / tot_filtered_cost if tot_filtered_cost > 0 else 0.0
        
        # Section B
        sec_b = master_df.groupby('Final_Division').agg({
            'Final_ASIN': 'nunique', 'PO_Units_Calc': 'sum', 'This_Wk_Cost_Calc': 'sum', 'Alert_Type': lambda x: len([i for i in x if '高WOS' in str(i) or '提报' in str(i)])
        }).reset_index()
        sec_b.columns = ['Division', 'WOS≤4\nASINs', 'Requested Units', 'Requested Cost (USD)', 'BTR⭐']
        sec_b['% of Total Cost'] = sec_b['Requested Cost (USD)'] / tot_filtered_cost if tot_filtered_cost > 0 else 0.0
        sec_b = sec_b[['Division', 'Requested Units', 'Requested Cost (USD)', '% of Total Cost', 'WOS≤4\nASINs', 'BTR⭐']].sort_values(by='Requested Cost (USD)', ascending=False).reset_index(drop=True)

        # --- 2. 📦 Old Detail 工作表组装 ---
        old_base = master_df[master_df['Product Tag'] == 'Old'].copy()
        old_pat = old_base.groupby(['Final_Division', 'Final_Brand', 'Final_Pattern']).agg({
            'Final_ASIN': 'nunique', 'PO_Units_Calc': 'sum', 'This_Wk_Cost_Calc': 'sum', 'JLA_Inv_Calc': 'sum', 'P70_Avg_Calc': 'mean', 'AMZ_WOS_Calc': 'mean'
        }).reset_index()
        old_pat.columns = ['Division', 'Brand', 'Pattern', 'ASIN Count', 'Units', 'Cost (USD)', 'JLA Inv Total', 'P70 Avg W1-8', 'AMZ WOS Avg']
        old_pat = old_pat.sort_values(by='Cost (USD)', ascending=False).reset_index(drop=True)
        
        old_asin_df = pd.DataFrame()
        old_asin_df['Division'] = old_base['Final_Division']
        old_asin_df['Brand'] = old_base['Final_Brand']
        old_asin_df['Pattern'] = old_base['Final_Pattern']
        old_asin_df['Product Category'] = old_base['Final_Category']
        old_asin_df['Subcategory'] = old_base['Final_Subcat']
        old_asin_df['Color'] = old_base['Final_Color']
        old_asin_df['Size'] = old_base['Final_Size']
        old_asin_df['ClassCode'] = old_base['Final_Class']
        old_asin_df['ASIN'] = old_base['Final_ASIN']
        old_asin_df['Units'] = old_base['PO_Units_Calc'].astype(int)
        old_asin_df['Cost (USD)'] = old_base['This_Wk_Cost_Calc']
        old_asin_df['JLA Inv'] = old_base['JLA_Inv_Calc'].astype(int)
        old_asin_df['P70 W1-8'] = old_base['P70_Avg_Calc'].round(1)
        old_asin_df['AMZ WOS'] = old_base['AMZ_WOS_Calc'].round(1)
        old_asin_df['BucketsList'] = old_base['Final_Buckets']
        old_asin_df = old_asin_df.sort_values(['Division', 'Pattern', 'Cost (USD)'], ascending=[True, True, False]).reset_index(drop=True)

        # --- 3. 🆕 New Release Detail 工作表组装 ---
        new_base = master_df[master_df['Product Tag'].isin(['Net new', 'New color/size'])].copy()
        new_pat = new_base.groupby(['Product Tag', 'Final_Division', 'Final_Brand', 'Final_Pattern']).agg({
            'Final_ASIN': 'nunique', 'PO_Units_Calc': 'sum', 'This_Wk_Cost_Calc': 'sum', 'JLA_Inv_Calc': 'sum', 'P70_Avg_Calc': 'mean', 'AMZ_WOS_Calc': 'mean'
        }).reset_index()
        new_pat.columns = ['Product Tag', 'Division', 'Brand', 'Pattern', 'ASIN Count', 'Units', 'Cost (USD)', 'JLA Inv Total', 'P70 Avg W1-8', 'AMZ WOS Avg']
        new_pat['Product Tag'] = new_pat['Product Tag'].apply(lambda x: '🆕 Net New' if x == 'Net new' else '🎨 New color/size')
        new_pat = new_pat.sort_values(['Product Tag', 'Cost (USD)'], ascending=[True, False]).reset_index(drop=True)

        # --- 4. 🚨 Action Required 工作表组装 ---
        ar_base = master_df[master_df['Alert_Type'] != '无预警'].copy()
        ar_base = ar_base.sort_values(['Alert_Pri', 'P70_Avg_Calc'], ascending=[True, False]).reset_index(drop=True)
        
        df_sheet4_ar = pd.DataFrame()
        df_sheet4_ar['Alert Type\n(预警类型)'] = ar_base['Alert_Type']
        df_sheet4_ar['Division'] = ar_base['Final_Division']
        df_sheet4_ar['OM'] = ar_base['Final_OM']
        df_sheet4_ar['Brand'] = ar_base['Final_Brand']
        df_sheet4_ar['Pattern'] = ar_base['Final_Pattern']
        df_sheet4_ar['Color'] = ar_base['Final_Color']
        df_sheet4_ar['Size'] = ar_base['Final_Size']
        df_sheet4_ar['Class\nCode'] = ar_base['Final_Class']
        df_sheet4_ar['ASIN'] = ar_base['Final_ASIN']
        df_sheet4_ar['L2WK Net\nPPM'] = ar_base['Net_PPM_Calc']
        df_sheet4_ar['Has\nPO'] = ar_base['has_PO_Calc'].apply(lambda x: '✅' if x else '—')
        df_sheet4_ar['PO Units\n(This Wk)'] = ar_base['PO_Units_Calc'].astype(int)
        df_sheet4_ar['FC On\nHand'] = ar_base['FC_OnHand_Calc'].astype(int)
        df_sheet4_ar['FC\nIncoming'] = ar_base['FC_Incoming_Calc'].astype(int)
        df_sheet4_ar['AMZ FC\nTotal'] = ar_base['FC_Total_Calc'].astype(int)
        df_sheet4_ar['JLA\nInv'] = ar_base['JLA_Inv_Calc'].astype(int)
        df_sheet4_ar['P70 Avg\nW1-8'] = ar_base['P70_Avg_Calc'].round(1)
        df_sheet4_ar['AMZ\nWOS'] = ar_base['AMZ_WOS_Calc'].round(1)
        df_sheet4_ar['Demand Trend\n(W1-4→W5-8)'] = ar_base['trend_pct_Calc']
        df_sheet4_ar['BucketsList'] = ar_base['Final_Buckets']
        df_sheet4_ar['Product\nTag'] = ar_base['Product Tag']

        # --- 5. 📈 WoW Top20 变化 工作表组装 ---
        df_sheet5_wow = pd.DataFrame()
        if wow_data is not None:
            wow_merge = pd.merge(master_df, wow_data, left_on='Final_ASIN', right_on='ASIN', how='left')
            wow_merge['P_Cost'] = wow_merge['P_Cost'].fillna(0.0)
            wow_merge['Cost_Change_Calc'] = wow_merge['This_Wk_Cost_Calc'] - wow_merge['P_Cost']
            wow_merge['Abs_Change'] = wow_merge['Cost_Change_Calc'].abs()
            
            wow_top20 = wow_merge.sort_values(by='Abs_Change', ascending=False).head(20).reset_index(drop=True)
            df_sheet5_wow['Rank'] = wow_top20.index + 1
            df_sheet5_wow['Division'] = wow_top20['Final_Division']
            df_sheet5_wow['OM'] = wow_top20['Final_OM']
            df_sheet5_wow['Brand'] = wow_top20['Final_Brand']
            df_sheet5_wow['Pattern'] = wow_top20['Final_Pattern']
            df_sheet5_wow['Color'] = wow_top20['Final_Color']
            df_sheet5_wow['Size'] = wow_top20['Final_Size']
            df_sheet5_wow['Class\nCode'] = wow_top20['Final_Class']
            df_sheet5_wow['ASIN'] = wow_top20['Final_ASIN']
            df_sheet5_wow['L2WK Net\nPPM'] = wow_top20['Net_PPM_Calc']
            df_sheet5_wow['BucketsList'] = wow_top20['Final_Buckets']
            df_sheet5_wow['Product\nTag'] = wow_top20['Product Tag']
            df_sheet5_wow['本周\nReq Cost'] = wow_top20['This_Wk_Cost_Calc']
            df_sheet5_wow['上周\nReq Cost'] = wow_top20['P_Cost']
            df_sheet5_wow['Cost\nChange'] = wow_top20['Cost_Change_Calc']
            df_sheet5_wow['变化\n方向'] = wow_top20['Cost_Change_Calc'].apply(lambda x: '▲ 增加' if x > 0 else ('▼ 减少' if x < 0 else '→ 持平'))
            df_sheet5_wow['FC On\nHand'] = wow_top20['FC_OnHand_Calc'].astype(int)
            df_sheet5_wow['FC\nIncoming'] = wow_top20['FC_Incoming_Calc'].astype(int)
            df_sheet5_wow['AMZ FC\nTotal'] = wow_top20['FC_Total_Calc'].astype(int)
            df_sheet5_wow['JLA\nInv'] = wow_top20['JLA_Inv_Calc'].astype(int)
            df_sheet5_wow['P70 Avg\nW1-8'] = wow_top20['P70_Avg_Calc'].round(1)
            df_sheet5_wow['AMZ\nWOS'] = wow_top20['AMZ_WOS_Calc'].round(1)
            df_sheet5_wow['Demand Trend\n(W1-4→W5-8)'] = wow_top20['trend_pct_Calc']

        # ==================== 🎨 渲染高亮主题包装 ====================
        def apply_po_theme_styles(df, is_ar=False, is_wow=False):
            def fmt_pct(v): return f"{round(v*100, 1)}%" if isinstance(v, (int,float)) else str(v)
            def fmt_money(v): return f"${v:,.2f}" if isinstance(v, (int,float)) else str(v)
            
            fmt_dict = {}
            for c in df.columns:
                if 'PPM' in c or 'Cost%' in c or 'Trend' in c or 'WOS' in c: fmt_dict[c] = fmt_pct
                elif 'Cost' in c or 'Change' in c: fmt_dict[c] = fmt_money
                elif 'Units' in c or 'Inv' in c or 'Total' in c or 'Hand' in c or 'Count' in c:
                    fmt_dict[c] = lambda x: f"{int(x):,}" if isinstance(x, (int,float)) else str(x)
            
            def row_painter(row):
                colors = [''] * len(row)
                for i, col_name in enumerate(row.index):
                    if is_ar and col_name == 'Alert Type\n(预警类型)':
                        val = str(row[col_name])
                        if '绿灯' in val: colors[i] = 'background-color: #EAF7EE; color: #1A7A3C; font-weight: bold'
                        elif '拦截' in val: colors[i] = 'background-color: #FDEAEA; color: #B22222; font-weight: bold'
                        elif '不足' in val: colors[i] = 'background-color: #FEF0E6; color: #C05700; font-weight: bold'
                        elif '暴跌' in val: colors[i] = 'background-color: #F3ECFC; color: #5B2C8D; font-weight: bold'
                        elif '断货' in val: colors[i] = 'background-color: #F5E3E3; color: #8B0000; font-weight: bold'
                        elif '高WOS' in val: colors[i] = 'background-color: #EDF7EA; color: #3A7D28; font-weight: bold'
                        continue
                    if is_wow and col_name == '变化\n方向':
                        val = str(row[col_name])
                        if '增加' in val: colors[i] = 'background-color: #E8F8F0; color: #1A7A3C; font-weight: bold'
                        elif '减少' in val: colors[i] = 'background-color: #FDEAEA; color: #B22222; font-weight: bold'
                        continue
                    if 'Units' in col_name or '销量' in col_name: colors[i] = 'background-color: #DDEBF7;'
                    elif 'Inv' in col_name or 'FC' in col_name or '库存' in col_name: colors[i] = 'background-color: #E2EFDA;'
                    elif 'Cost' in col_name or 'Change' in col_name: colors[i] = 'background-color: #FFF2CC;'
                    elif 'PPM' in col_name: colors[i] = 'background-color: #FCE4D6;'
                return colors
            return df.style.apply(row_painter, axis=1).format(fmt_dict)

        styler_s1_a = apply_po_theme_styles(sec_a)
        styler_s1_b = apply_po_theme_styles(sec_b)
        styler_s2_pat = apply_po_theme_styles(old_pat)
        styler_s2_asin = apply_po_theme_styles(old_asin_df)
        styler_s3_pat = apply_po_theme_styles(new_pat)
        styler_s4_ar = apply_po_theme_styles(df_sheet4_ar, is_ar=True)
        styler_s5_wow = apply_po_theme_styles(df_sheet5_wow, is_wow=True) if wow_data is not None else pd.DataFrame()

        h_c = len(df_sheet4_ar[df_sheet4_ar['Alert Type\n(预警类型)'].str.contains('提报|拦截|High', na=False)])
        m_c = len(df_sheet4_ar[df_sheet4_ar['Alert Type\n(预警类型)'].str.contains('不足量|Medium', na=False)])
        l_c = len(df_sheet4_ar[df_sheet4_ar['Alert Type\n(预警类型)'].str.contains('暴跌|Low', na=False)])
        i_c = len(df_sheet4_ar[df_sheet4_ar['Alert Type\n(预警类型)'].str.contains('OOS|断货|高WOS', na=False)])

        summary_table = pd.DataFrame([
            {'预警等级': '🔴 核心决策层 (High / 提报 / 拦截)', '核心判定条件说明': 'WOS≤4 缺货严重，JLA储备充足，且利润率高(绿灯提报) 或 利润过低(拦截红灯)', '已筛选子ASIN触发数': f"{h_c} 个"},
            {'预警等级': '⚠️ 运营干预层 (PO不足量 / Medium)', '核心判定条件说明': '亚马逊已下了PO订单，但是(现有库存+PO订货)仍然填不满 4周的周均销缺口', '已筛选子ASIN触发数': f"{m_c} 个"},
            {'预警等级': '⚪ 供应链风控层 (需求暴跌 / Low)', '核心判定条件说明': 'JLA大货在库，但是AMZ端WOS已满，且P70后四周预测大跌25%以上', '已筛选子ASIN触发数': f"{l_c} 个"},
            {'预警等级': 'ℹ️ 基础链接层 (OOS断货 / 高WOS)', '核心判定条件说明': 'AMZ总库存低于50件进入断货警戒；或WOS超出12周以上面临长期仓储费风险', '已筛选子ASIN触发数': f"{i_c} 个"}
        ])

        # ==================== 📥 Excel 导出打包 ====================
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_table.to_excel(writer, sheet_name='📊 Summary说明', index=False)
            sec_a.to_excel(writer, sheet_name='📊 Summary_大盘总览', index=False)
            sec_b.to_excel(writer, sheet_name='📊 Summary_部门排行', index=False)
            old_pat.to_excel(writer, sheet_name='📦 Old_Pattern_Summary', index=False)
            old_asin_df.to_excel(writer, sheet_name='📦 Old_ASIN_Detail', index=False)
            new_pat.to_excel(writer, sheet_name='🆕 New_Release_Detail', index=False)
            df_sheet4_ar.to_excel(writer, sheet_name='🚨 Action Required', index=False)
            if wow_data is not None: df_sheet5_wow.to_excel(writer, sheet_name='📈 WoW Top20 变化', index=False)
            
        st.download_button(
            label="📥 一键导出严格规范 Excel 预警报告 (.xlsx)",
            data=output.getvalue(),
            file_name=f"VC_Weekly_PO_Report_{order_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        # ==================== 📊 页面新装 4-5 大 Tab 专属业务视图 ====================
        tabs = st.tabs(["📊 Summary", "📦 Old Detail", "🆕 New Release Detail", "🚨 Action Required", "📈 WoW Top20 变化"])
        
        with tabs[0]:
            st.subheader("📋 预警触发摘要统计说明")
            st.table(summary_table)
            st.subheader("📦 Section A: 本周 PO 总览 (含 Pets/Hardline)")
            st.dataframe(styler_s1_a, use_container_width=True)
            st.subheader("📋 Section B: Division 汇总 (按 Cost 降序)")
            st.dataframe(styler_s1_b, use_container_width=True)
            
        with tabs[1]:
            st.subheader("🪵 上半部分: Pattern 汇总")
            st.dataframe(styler_s2_pat, use_container_width=True, height=350)
            st.subheader("🔍 下半部分: ASIN 明细清单")
            st.dataframe(styler_s2_asin, use_container_width=True, height=450)
            
        with tabs[2]:
            st.subheader("🆕 新品大类 Pattern 与 ASIN 交叉监测")
            st.dataframe(styler_s3_pat, use_container_width=True, height=550)
            
        with tabs[3]:
            st.caption("💡 警报徽章说明: ①BTR提报(绿灯)利润足 | ②BTR拦截(红灯)利润低 | ③PO不足量补货少 | ④需求暴跌滞压防死货 | ⑤OOS断货 | ⑥高WOS压仓")
            st.dataframe(styler_s4_ar, use_container_width=True, height=550)
            
        with tabs[4]:
            if wow_data is not None:
                st.dataframe(styler_s5_wow, use_container_width=True, height=550)
            else:
                st.info("💡 **WoW 环比功能未激活**：检测到您本次仅投入了本周的 4 张基础表格。如果您想查看与上周对比的 WoW Top20 巨幅异动表，只需在左侧栏重新多选文件，**把上周的原始 PO 导出表一并扔进来**，系统就会自动为你解锁本 Tab 视图！")

except Exception as e:
    st.error(f"💥 发生运算错误：{e}")
