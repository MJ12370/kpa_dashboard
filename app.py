import streamlit as st
import pandas as pd
import numpy as np
import os
import glob
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="제지산업 수출입통계 대시보드", layout="wide")

# 1. 파일 경로 탐색
if os.path.exists(r"D:\kita"):
    DATA_DIR = r"D:\kita"
else:
    DATA_DIR = "."

excel_files = glob.glob(os.path.join(DATA_DIR, "제지산업_수출입통계_통합_*.xlsx"))

if not excel_files:
    st.error("데이터 엑셀 파일을 찾을 수 없습니다. 폴더 내 파일명을 확인해 주세요.")
    st.stop()

latest_file = max(excel_files, key=os.path.getctime)

@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path, sheet_name="통합_전체데이터")
    
    # 수출입 실적(톤) 단위 통일
    if '수출실적(톤)' not in df.columns and '수출중량(kg)' in df.columns:
        df['수출실적(톤)'] = df['수출중량(kg)'] / 1000.0
    if '수입실적(톤)' not in df.columns and '수입중량(kg)' in df.columns:
        df['수입실적(톤)'] = df['수입중량(kg)'] / 1000.0
        
    # 기준년월에서 숫자만 정밀 추출
    raw_date = df['기준년월'].astype(str).str.replace(r'\.0$', '', regex=True)
    clean_digits = raw_date.str.replace(r'[^0-9]', '', regex=True)
    
    df['기준연도'] = clean_digits.str.slice(0, 4)
    month_series = pd.to_numeric(clean_digits.str.slice(4, 6), errors='coerce').fillna(0).astype(int)
    df['월'] = month_series
    df['기준년월'] = df['기준연도'] + "." + df['월'].apply(lambda x: f"{x:02d}")
    
    if '중분류' not in df.columns:
        df['중분류'] = df['품목명']
        
    return df

df = load_data(latest_file)

# 2. 상단 필터 UI
raw_cats = df['대분류'].dropna().unique().tolist()
raw_cats = [c for c in raw_cats if c not in ['폐신문', '폐신문지']]
priority_order = ['폐지', '골판지원지', '펄프']
available_cats = [c for c in priority_order if c in raw_cats] + [c for c in raw_cats if c not in priority_order]

trade_modes = ["수출", "수입", "수출+수입"]

c_cat, c_trade, c_sub, c_period = st.columns([1.1, 1.4, 1.4, 1.1])

with c_cat:
    selected_cat = st.radio("📂 **대분류 품목**", available_cats, horizontal=True)

with c_trade:
    trade_type = st.radio("🔄 **수출 / 수입 구분**", trade_modes, horizontal=True)

waste_items = ['폐골판지', '폐신문지', '고급폐지', '기타폐지']

with c_sub:
    if trade_type == "수출+수입 동시조회":
        if selected_cat == '폐지':
            sub_options = ["전체 합계"] + waste_items
            chosen_sub_item = st.radio("📑 **세부 품목 선택**", sub_options, horizontal=True)
        else:
            other_sub_items = sorted(df[df['대분류'] == selected_cat]['중분류'].dropna().unique().tolist())
            sub_options = ["전체 합계"] + other_sub_items
            chosen_sub_item = st.radio("📑 **세부 품목 선택**", sub_options, horizontal=True)
    else:
        st.write("")
        chosen_sub_item = None

with c_period:
    period_type = st.radio("📅 **집계 주기**", ["연간 합계 (YoY)", "월별 실적 (MoM)"], horizontal=True)

# 3. 사이드바 및 데이터 필터링
filtered_df = df[df['대분류'] == selected_cat].copy()
st.sidebar.markdown("### 🌐 국가별 상세 조회")

# -------------------------------------------------------------
# CASE A: 수출+수입 동시조회 모드 (순수출/총교역량 제거)
# -------------------------------------------------------------
if trade_type == "수출+수입 동시조회":
    if chosen_sub_item and chosen_sub_item != "전체 합계":
        filtered_df = filtered_df[filtered_df['중분류'] == chosen_sub_item].copy()
        display_sub_name = chosen_sub_item
    else:
        display_sub_name = f"{selected_cat} 전체"

    all_countries = sorted(filtered_df['국가명'].dropna().unique().tolist())
    top10_countries = (
        filtered_df.groupby('국가명')[['수출실적(톤)', '수입실적(톤)']]
        .sum().sum(axis=1)
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    other_countries = [c for c in all_countries if c not in top10_countries]

    country_options = ["전체 합산"] + top10_countries + ["기타 (직접 선택)"]
    selected_option = st.sidebar.radio(f"🏆 **{display_sub_name} 교역 상위 10개국**", country_options, index=0)

    if selected_option == "전체 합산":
        country_title_label = "[전체 국가]"
    elif selected_option == "기타 (직접 선택)":
        chosen_other = st.sidebar.selectbox("🔍 기타 국가를 선택하세요:", other_countries)
        filtered_df = filtered_df[filtered_df['국가명'] == chosen_other]
        country_title_label = f"[{chosen_other}]"
    else:
        filtered_df = filtered_df[filtered_df['국가명'] == selected_option]
        country_title_label = f"[{selected_option}]"

    if filtered_df.empty:
        st.warning("선택된 조건에 해당하는 데이터가 없습니다.")
        st.stop()

    if "연간" in period_type:
        valid_years = sorted([y for y in filtered_df['기준연도'].unique() if len(str(y)) == 4 and str(y).isdigit()])
        last_year = valid_years[-1]
        last_year_months = sorted([m for m in filtered_df[filtered_df['기준연도'] == last_year]['월'].unique() if 1 <= m <= 12])
        is_partial = (len(last_year_months) < 12 and len(last_year_months) > 0)
        
        if is_partial:
            partial_label = f"{last_year}.{last_year_months[0]}-{last_year_months[-1]}"
        else:
            partial_label = last_year

        pivot_full = filtered_df.pivot_table(
            index='기준연도',
            values=['수출실적(톤)', '수입실적(톤)'],
            aggfunc='sum'
        ).fillna(0)

        if is_partial:
            df_same = filtered_df[filtered_df['월'].isin(last_year_months)]
            pivot_partial = df_same.pivot_table(
                index='기준연도',
                values=['수출실적(톤)', '수입실적(톤)'],
                aggfunc='sum'
            ).fillna(0)
        else:
            pivot_partial = pivot_full.copy()

        pivot_diff = pivot_full.diff()
        pivot_pct = pivot_full.pct_change() * 100.0

        if is_partial and len(valid_years) >= 2:
            prev_year = valid_years[-2]
            curr_ytd = pivot_partial.loc[last_year]
            prev_ytd = pivot_partial.loc[prev_year] if prev_year in pivot_partial.index else pd.Series(0, index=pivot_partial.columns)
            diff_ytd = curr_ytd - prev_ytd
            pct_ytd = (diff_ytd / prev_ytd.replace(0, np.nan)) * 100.0

            pivot_full = pivot_full.rename(index={last_year: partial_label})
            pivot_diff = pivot_diff.rename(index={last_year: partial_label})
            pivot_pct = pivot_pct.rename(index={last_year: partial_label})

            pivot_diff.loc[partial_label] = diff_ytd
            pivot_pct.loc[partial_label] = pct_ytd

        pivot_pct = pivot_pct.replace([np.inf, -np.inf], np.nan)
        date_index_col = '연도'
        pivot_base = pivot_full

    else:
        pivot_base = filtered_df.pivot_table(
            index='기준년월',
            values=['수출실적(톤)', '수입실적(톤)'],
            aggfunc='sum'
        ).fillna(0)
        pivot_diff = pivot_base.diff()
        pivot_pct = pivot_base.pct_change() * 100.0
        pivot_pct = pivot_pct.replace([np.inf, -np.inf], np.nan)
        date_index_col = '기준년월'

    # 2중 헤더: 수출 3열, 수입 3열만 구성
    final_data = {
        ('수출', '실적'): pivot_base['수출실적(톤)'],
        ('수출', '증감량'): pivot_diff['수출실적(톤)'],
        ('수출', '증감률'): pivot_pct['수출실적(톤)'],
        ('수입', '실적'): pivot_base['수입실적(톤)'],
        ('수입', '증감량'): pivot_diff['수입실적(톤)'],
        ('수입', '증감률'): pivot_pct['수입실적(톤)']
    }
    pivot_final = pd.DataFrame(final_data, index=pivot_base.index)

    st.markdown(f"### 📋 {display_sub_name} {country_title_label} 수출/수입 실적 ({period_type})")
    st.caption("(단위 : 톤)")

    def format_values(val, col_type):
        if pd.isna(val):
            return "-"
        if col_type == '증감률':
            return f"{val:,.1f}"
        return f"{int(round(val)):,}"

    def apply_styles(val):
        if pd.isna(val):
            return 'color: #888888;'
        if isinstance(val, (int, float)) and val < 0:
            return 'color: #d9534f; font-weight: bold;'
        return 'color: #212529;'

    format_dict = {col: (lambda v, t=col[1]: format_values(v, t)) for col in pivot_final.columns}
    styled_table = pivot_final.style.format(format_dict).map(apply_styles)
    st.dataframe(styled_table, use_container_width=True, height=450)

    # 수출 vs 수입 추이 차트
    st.write("---")
    st.markdown(f"### 📊 {display_sub_name} {country_title_label} 수출 vs 수입 추이 차트")

    if "월별" in period_type:
        chart_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01')
    else:
        chart_x = pivot_base.index.astype(str)

    fig_dual_line = go.Figure()
    fig_dual_line.add_trace(go.Scatter(x=chart_x, y=pivot_base['수출실적(톤)'], name='수출실적(톤)', mode='lines+markers', line=dict(color='#2E7D32', width=2.5)))
    fig_dual_line.add_trace(go.Scatter(x=chart_x, y=pivot_base['수입실적(톤)'], name='수입실적(톤)', mode='lines+markers', line=dict(color='#C62828', width=2.5)))

    if "월별" in period_type:
        fig_dual_line.update_xaxes(dtick="M3", tickformat="%Y-%m")

    fig_dual_line.update_layout(
        margin=dict(l=20, r=20, t=30, b=20),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="실적(톤)", tickformat=",.0f")
    )
    st.plotly_chart(fig_dual_line, use_container_width=True)

# -------------------------------------------------------------
# CASE B: 단일 수출 또는 수입 모드
# -------------------------------------------------------------
else:
    target_col = '수출실적(톤)' if trade_type == '수출' else '수입실적(톤)'

    if selected_cat == '폐지':
        all_waste_countries = sorted(filtered_df['국가명'].dropna().unique().tolist())
        top10_countries = (
            filtered_df.groupby('국가명')[target_col]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .index.tolist()
        )
        other_countries = [c for c in all_waste_countries if c not in top10_countries]
        country_options = ["전체 합산"] + top10_countries + ["기타 (직접 선택)"]
        selected_option = st.sidebar.radio(f"🏆 **폐지 {trade_type} 국가 선택**", country_options, index=0)
        
        if selected_option == "전체 합산":
            country_title_label = "[전체 국가]"
        elif selected_option == "기타 (직접 선택)":
            chosen_other = st.sidebar.selectbox("🔍 기타 국가를 선택하세요:", other_countries)
            filtered_df = filtered_df[filtered_df['국가명'] == chosen_other]
            country_title_label = f"[{chosen_other}]"
        else:
            filtered_df = filtered_df[filtered_df['국가명'] == selected_option]
            country_title_label = f"[{selected_option}]"
    else:
        all_countries = sorted(filtered_df['국가명'].dropna().unique().tolist())
        selected_countries = st.sidebar.multiselect("국가 필터 (선택 안 하면 전체 합산)", all_countries)
        if selected_countries:
            filtered_df = filtered_df[filtered_df['국가명'].isin(selected_countries)]
            country_title_label = f"[{', '.join(selected_countries)}]"
        else:
            country_title_label = "[전체 국가]"

    if filtered_df.empty:
        st.warning("선택된 조건에 해당하는 데이터가 없습니다.")
        st.stop()

    if "연간" in period_type:
        valid_years = sorted([y for y in filtered_df['기준연도'].unique() if len(str(y)) == 4 and str(y).isdigit()])
        last_year = valid_years[-1]
        last_year_months = sorted([m for m in filtered_df[filtered_df['기준연도'] == last_year]['월'].unique() if 1 <= m <= 12])
        is_partial = (len(last_year_months) < 12 and len(last_year_months) > 0)
        
        if is_partial:
            start_m = last_year_months[0]
            end_m = last_year_months[-1]
            partial_label = f"{last_year}.{start_m}-{end_m}"
        else:
            partial_label = last_year

        pivot_full = filtered_df.pivot_table(
            index='기준연도',
            columns='중분류',
            values=target_col,
            aggfunc='sum'
        ).fillna(0)

        if is_partial:
            df_same_period = filtered_df[filtered_df['월'].isin(last_year_months)]
            pivot_partial = df_same_period.pivot_table(
                index='기준연도',
                columns='중분류',
                values=target_col,
                aggfunc='sum'
            ).fillna(0)
        else:
            pivot_partial = pivot_full.copy()

        all_cols = sorted(pivot_full.columns.tolist())
        if selected_cat == '폐지':
            custom_waste_order = ['폐골판지', '폐신문지', '고급폐지', '기타폐지']
            ordered_cols = [c for c in custom_waste_order if c in all_cols]
            remaining_cols = [c for c in all_cols if c not in custom_waste_order]
            target_cols = ordered_cols + remaining_cols
        else:
            target_cols = all_cols

        pivot_full = pivot_full.reindex(columns=target_cols, fill_value=0)
        pivot_partial = pivot_partial.reindex(columns=target_cols, fill_value=0)

        pivot_full['합계'] = pivot_full.sum(axis=1)
        pivot_partial['합계'] = pivot_partial.sum(axis=1)

        pivot_diff = pivot_full.diff()
        pivot_pct = pivot_full.pct_change() * 100.0

        if is_partial and len(valid_years) >= 2:
            prev_year = valid_years[-2]
            prev_ytd_series = pivot_partial.loc[prev_year] if prev_year in pivot_partial.index else pd.Series(0, index=pivot_full.columns)
            curr_ytd_series = pivot_partial.loc[last_year]

            diff_ytd = curr_ytd_series - prev_ytd_series
            pct_ytd = (diff_ytd / prev_ytd_series.replace(0, np.nan)) * 100.0

            pivot_full = pivot_full.rename(index={last_year: partial_label})
            pivot_diff = pivot_diff.rename(index={last_year: partial_label})
            pivot_pct = pivot_pct.rename(index={last_year: partial_label})

            pivot_diff.loc[partial_label] = diff_ytd
            pivot_pct.loc[partial_label] = pct_ytd

        pivot_pct = pivot_pct.replace([np.inf, -np.inf], np.nan)
        pivot_base = pivot_full
        date_index_col = '연도'

    else:
        pivot_base = filtered_df.pivot_table(
            index='기준년월',
            columns='중분류',
            values=target_col,
            aggfunc='sum'
        ).fillna(0)

        if selected_cat == '폐지':
            custom_waste_order = ['폐골판지', '폐신문지', '고급폐지', '기타폐지']
            ordered_cols = [c for c in custom_waste_order if c in pivot_base.columns]
            remaining_cols = [c for c in pivot_base.columns if c not in custom_waste_order]
            pivot_base = pivot_base[ordered_cols + remaining_cols]

        pivot_base['합계'] = pivot_base.sum(axis=1)
        pivot_diff = pivot_base.diff()
        pivot_pct = pivot_base.pct_change() * 100.0
        pivot_pct = pivot_pct.replace([np.inf, -np.inf], np.nan)
        date_index_col = '기준년월'

    final_data = {}
    for col in pivot_base.columns:
        final_data[(col, col if col != '합계' else '합계')] = pivot_base[col]
        final_data[(col, '증감량')] = pivot_diff[col]
        final_data[(col, '증감률')] = pivot_pct[col]

    pivot_final = pd.DataFrame(final_data, index=pivot_base.index)

    st.markdown(f"### 📋 {selected_cat} {country_title_label} {trade_type}실적 ({period_type})")
    st.caption("(단위 : 톤)")

    def format_values(val, col_type):
        if pd.isna(val):
            return "-"
        if col_type == '증감률':
            return f"{val:,.1f}"
        return f"{int(round(val)):,}"

    def apply_styles(val):
        if pd.isna(val):
            return 'color: #888888;'
        if isinstance(val, (int, float)) and val < 0:
            return 'color: #d9534f; font-weight: bold;'
        return 'color: #212529;'

    format_dict = {col: (lambda v, t=col[1]: format_values(v, t)) for col in pivot_final.columns}
    styled_table = pivot_final.style.format(format_dict).map(apply_styles)
    st.dataframe(styled_table, use_container_width=True, height=450)

    st.write("---")
    st.markdown(f"### 📊 {selected_cat} {country_title_label} {trade_type} 추이 차트")

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.markdown("##### 📈 세부 품목별 실적 (톤)")
        item_cols = [c for c in pivot_base.columns if c != '합계']
        chart_line_df = pivot_base[item_cols].reset_index()
        chart_line_df.rename(columns={chart_line_df.columns[0]: date_index_col}, inplace=True)
        
        if "월별" in period_type:
            chart_line_df['날짜'] = pd.to_datetime(chart_line_df[date_index_col].astype(str).str.replace('.', '-') + '-01')
            x_col = '날짜'
        else:
            x_col = date_index_col

        chart_line_df = chart_line_df.melt(
            id_vars=x_col, 
            value_vars=item_cols,
            var_name='품목', 
            value_name='실적(톤)'
        )
        
        fig_line = px.line(chart_line_df, x=x_col, y='실적(톤)', color='품목', markers=True)
        if "월별" in period_type:
            fig_line.update_xaxes(dtick="M3", tickformat="%Y-%m")

        fig_line.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            yaxis=dict(title="실적(톤)", tickformat=",.0f")
        )
        st.plotly_chart(fig_line, use_container_width=True)

    with chart_col2:
        st.markdown("##### 🏛️ 합계 실적 및 증감량 (톤)")
        fig_bar = go.Figure()
        
        if "월별" in period_type:
            bar_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01')
        else:
            bar_x = pivot_base.index.astype(str)

        fig_bar.add_trace(go.Bar(x=bar_x, y=pivot_base['합계'], name='총 실적(톤)', marker_color='#4A90E2'))
        fig_bar.add_trace(go.Scatter(x=bar_x, y=pivot_diff['합계'], name='증감량(톤)', mode='lines+markers', line=dict(color='#E2594A', width=2), yaxis='y2'))
        
        if "월별" in period_type:
            fig_bar.update_xaxes(dtick="M3", tickformat="%Y-%m")

        fig_bar.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            yaxis=dict(title="총 실적(톤)", tickformat=",.0f"),
            yaxis2=dict(title="증감량(톤)", overlaying='y', side='right', showgrid=False, tickformat=",.0f")
        )
        st.plotly_chart(fig_bar, use_container_width=True)