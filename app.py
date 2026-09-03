import streamlit as st
import pandas as pd
import numpy as np
import os
import glob
import io
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components

st.set_page_config(page_title="제지산업 수출입 통관실적 대시보드", layout="wide")

# 인쇄 및 공통 스타일
st.markdown("""
<style>
@media print {
    [data-testid="stSidebar"], 
    header, 
    [data-testid="stRadio"], 
    [data-testid="stSelectbox"], 
    .stButton, 
    .stDownloadButton,
    iframe,
    button {
        display: none !important;
    }
    .chart-box, 
    .chart-box *, 
    .stPlotlyChart, 
    hr {
        display: none !important;
    }
    .block-container {
        padding: 0 !important;
        margin: 0 !important;
    }
}
.custom-table-container {
    width: 100%;
    overflow-x: auto;
    margin-bottom: 0.5rem;
    border: 1.5px solid #1E3A8A;
    border-radius: 4px;
}
.custom-table {
    width: 100%;
    border-collapse: collapse;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    font-size: 13px;
    white-space: nowrap;
}
.custom-table th {
    background-color: #1E3A8A !important;
    color: #FFFFFF !important;
    font-weight: 700;
    padding: 8px 8px;
    border: 1px solid #1E40AF;
    text-align: center !important;
}
.custom-table td {
    padding: 6px 14px 6px 10px;
    border: 1px solid #94A3B8;
    font-weight: 600;
    color: #0F172A;
}
.custom-table td.col-label {
    text-align: center !important;
    padding: 6px 8px !important;
}
.custom-table td.col-num {
    text-align: right !important;
}
.custom-table tr:nth-child(even) td {
    background-color: #F8FAFC;
}
.custom-table tr:hover td {
    background-color: #EFF6FF;
}
.row-subtotal td {
    background-color: #E2E8F0 !important;
    font-weight: 700 !important;
    border-top: 1.5px solid #64748B !important;
    border-bottom: 1.5px solid #64748B !important;
}
.row-total td {
    background-color: #FEF3C7 !important;
    font-weight: 800 !important;
    color: #92400E !important;
    border-top: 2px solid #D97706 !important;
    border-bottom: 2px solid #D97706 !important;
}
.row-grand-total td {
    background-color: #DBEAFE !important;
    font-weight: 800 !important;
    color: #1E3A8A !important;
    border-top: 2px solid #1E3A8A !important;
    border-bottom: 2px solid #1E3A8A !important;
}
.val-negative {
    color: #DC2626 !important;
    font-weight: 700;
}
.val-zero {
    color: #94A3B8 !important;
}
</style>
""", unsafe_allow_html=True)

DATA_DIR = r"D:\kita" if os.path.exists(r"D:\kita") else "."

# 최상단 대시보드 메뉴
st.sidebar.markdown("## 📊 수출입 통관실적")
main_menu = st.sidebar.radio("품목 구분", ["원료", "종이판지"], index=0)
st.sidebar.write("---")

# ==========================================
# 1. 원료 수출입 통관실적 대시보드
# ==========================================
if main_menu == "원료":
    raw_files = glob.glob(os.path.join(DATA_DIR, "제지산업_수출입통계_통합_*.xlsx"))
    if not raw_files:
        st.error("원료 데이터 엑셀 파일을 찾을 수 없습니다. 폴더 내 파일명을 확인해 주세요.")
        st.stop()
    latest_file = max(raw_files, key=os.path.getctime)

    @st.cache_data
    def load_data(file_path):
        df = pd.read_excel(file_path, sheet_name="통합_전체데이터")
        if '수출실적(톤)' not in df.columns and '수출중량(kg)' in df.columns:
            df['수출실적(톤)'] = df['수출중량(kg)'] / 1000.0
        if '수입실적(톤)' not in df.columns and '수입중량(kg)' in df.columns:
            df['수입실적(톤)'] = df['수입중량(kg)'] / 1000.0
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

    st.title("원료 수출입 통관실적")
    st.write("---")

    waste_order = ['전체', '폐골판지', '폐신문지', '고급폐지', '기타폐지']
    pulp_order = ['전체', 'GP', 'DP', 'UKP', 'BKP', 'BCTMP', '면린터펄프', 'DIP', '기타']
    liner_order = ['전체', '라이너', '골심지']

    st.sidebar.markdown("### 📂 품목 선택")
    selected_cat = st.sidebar.radio("대분류", ['폐지', '펄프', '골판지원지'])
    cat_df = df[df['대분류'] == selected_cat]

    if selected_cat == '폐지':
        actual_items = cat_df['중분류'].dropna().unique().tolist()
        sub_choices = [item for item in waste_order if item == '전체' or item in actual_items]
    elif selected_cat == '펄프':
        actual_items = cat_df['중분류'].dropna().unique().tolist()
        ordered_items = [item for item in pulp_order if item == '전체' or item in actual_items]
        rem_items = [item for item in actual_items if item not in pulp_order]
        sub_choices = ordered_items + rem_items
    else:
        actual_items = cat_df['중분류'].dropna().unique().tolist()
        ordered_items = [item for item in liner_order if item == '전체' or item in actual_items]
        rem_items = [item for item in actual_items if item not in liner_order]
        sub_choices = ordered_items + rem_items

    chosen_sub_item = st.sidebar.radio(f"세부 품목 ({selected_cat})", sub_choices)

    st.sidebar.write("---")
    st.sidebar.markdown("### 🔄 무역 구분")
    trade_type = st.sidebar.radio("수출 / 수입 선택", ["수출", "수입", "수출+수입"])

    st.sidebar.write("---")
    st.sidebar.markdown("### 🌐 국가 선택")
    country_calc_df = cat_df[cat_df['중분류'] == chosen_sub_item] if chosen_sub_item != "전체" else cat_df
    all_countries = sorted(country_calc_df['국가명'].dropna().unique().tolist())
    top10_countries = (
        country_calc_df.groupby('국가명')[['수출실적(톤)', '수입실적(톤)']]
        .sum().sum(axis=1).sort_values(ascending=False).head(10).index.tolist()
    )
    other_countries = [c for c in all_countries if c not in top10_countries]

    c_options = ["전체 국가"] + top10_countries + ["기타 (직접 선택)"]
    country_pick = st.sidebar.radio("상위 교역국", c_options)

    if country_pick == "전체 국가":
        selected_country = "전체 국가"
    elif country_pick == "기타 (직접 선택)":
        selected_country = st.sidebar.selectbox("기타 국가 선택", other_countries)
    else:
        selected_country = country_pick

    all_years = sorted([y for y in df['기준연도'].unique() if len(str(y)) == 4 and str(y).isdigit()])
    all_dates = sorted(df['기준년월'].unique().tolist())

    col_period_type, col_start, col_end = st.columns([1.2, 1.4, 1.4])
    with col_period_type:
        period_mode = st.radio("📅 **조회기간**", ["연간", "월간"], horizontal=True)

    if period_mode == "연간":
        period_type = "연간 합계 (YoY)"
        with col_start:
            start_year = st.selectbox("시작 연도", all_years, index=0)
        with col_end:
            valid_end_years = [y for y in all_years if y >= start_year]
            end_year = st.selectbox("종료 연도", valid_end_years, index=len(valid_end_years)-1)
        selected_desc = f"{start_year}년 ~ {end_year}년 (연간)"
    else:
        period_type = "월별 실적 (MoM)"
        with col_start:
            start_date = st.selectbox("시작 연월", all_dates, index=max(0, len(all_dates)-24))
        with col_end:
            valid_end_dates = [d for d in all_dates if d >= start_date]
            end_date = st.selectbox("종료 연월", valid_end_dates, index=len(valid_end_dates)-1)
        selected_desc = f"{start_date} ~ {end_date} (월별)"

    filtered_df = cat_df.copy()
    display_item_title = chosen_sub_item if chosen_sub_item != "전체" else f"{selected_cat}"
    if chosen_sub_item != "전체":
        filtered_df = filtered_df[filtered_df['중분류'] == chosen_sub_item]

    if selected_country != "전체 국가":
        filtered_df = filtered_df[filtered_df['국가명'] == selected_country]
        title_country_str = f" ({selected_country})"
    else:
        title_country_str = ""

    if period_mode == "연간":
        filtered_df = filtered_df[(filtered_df['기준연도'] >= start_year) & (filtered_df['기준연도'] <= end_year)]
    else:
        filtered_df = filtered_df[(filtered_df['기준년월'] >= start_date) & (filtered_df['기준년월'] <= end_date)]

    if filtered_df.empty:
        st.warning("선택된 기간 및 조건에 해당하는 데이터가 없습니다.")
        st.stop()

    valid_years = sorted([y for y in filtered_df['기준연도'].unique() if len(str(y)) == 4 and str(y).isdigit()])
    last_year = valid_years[-1]
    last_year_months = sorted([m for m in filtered_df[filtered_df['기준연도'] == last_year]['월'].unique() if 1 <= m <= 12])
    is_partial = (len(last_year_months) < 12 and len(last_year_months) > 0 and period_mode == "연간")
    partial_label = f"{last_year}.{last_year_months[0]}-{last_year_months[-1]}" if is_partial else last_year

    if trade_type == "수출+수입" or chosen_sub_item != "전체":
        if period_mode == "연간":
            pivot_full = filtered_df.pivot_table(index='기준연도', values=['수출실적(톤)', '수입실적(톤)'], aggfunc='sum').fillna(0)
            pivot_diff = pivot_full.diff()
            pivot_pct = pivot_full.pct_change() * 100.0
            if is_partial and len(valid_years) >= 2:
                prev_year_val = valid_years[-2]
                df_same = filtered_df[filtered_df['월'].isin(last_year_months)]
                pivot_partial = df_same.pivot_table(index='기준연도', values=['수출실적(톤)', '수입실적(톤)'], aggfunc='sum').fillna(0)
                curr_ytd = pivot_partial.loc[last_year]
                prev_ytd = pivot_partial.loc[prev_year_val] if prev_year_val in pivot_partial.index else pd.Series(0, index=pivot_partial.columns)
                diff_ytd = curr_ytd - prev_ytd
                pct_ytd = (diff_ytd / prev_ytd.replace(0, np.nan)) * 100.0
                pivot_full = pivot_full.rename(index={last_year: partial_label})
                pivot_diff = pivot_diff.rename(index={last_year: partial_label})
                pivot_pct = pivot_pct.rename(index={last_year: partial_label})
                pivot_diff.loc[partial_label] = diff_ytd
                pivot_pct.loc[partial_label] = pct_ytd
            pivot_base = pivot_full
            date_index_col = '연도'
        else:
            pivot_base = filtered_df.pivot_table(index='기준년월', values=['수출실적(톤)', '수입실적(톤)'], aggfunc='sum').fillna(0)
            pivot_diff = pivot_base.diff()
            pivot_pct = pivot_base.pct_change() * 100.0
            date_index_col = '기준년월'

        if trade_type == "수출":
            final_data = {('수출', '중량(톤)'): pivot_base['수출실적(톤)'], ('수출', '증감량'): pivot_diff['수출실적(톤)'], ('수출', '증감률(%)'): pivot_pct['수출실적(톤)']}
        elif trade_type == "수입":
            final_data = {('수입', '중량(톤)'): pivot_base['수입실적(톤)'], ('수입', '증감량'): pivot_diff['수입실적(톤)'], ('수입', '증감률(%)'): pivot_pct['수입실적(톤)']}
        else:
            final_data = {
                ('수출', '중량(톤)'): pivot_base['수출실적(톤)'], ('수출', '증감량'): pivot_diff['수출실적(톤)'], ('수출', '증감률(%)'): pivot_pct['수출실적(톤)'],
                ('수입', '중량(톤)'): pivot_base['수입실적(톤)'], ('수입', '증감량'): pivot_diff['수입실적(톤)'], ('수입', '증감률(%)'): pivot_pct['수입실적(톤)']
            }
        pivot_final = pd.DataFrame(final_data, index=pivot_base.index)
    else:
        target_col = '수출실적(톤)' if trade_type == '수출' else '수입실적(톤)'
        if period_mode == "연간":
            pivot_full = filtered_df.pivot_table(index='기준연도', columns='중분류', values=target_col, aggfunc='sum').fillna(0)
            all_cols = pivot_full.columns.tolist()
            sort_ref = [c for c in waste_order if c != '전체'] if selected_cat == '폐지' else ([c for c in pulp_order if c != '전체'] if selected_cat == '펄프' else [c for c in liner_order if c != '전체'])
            ordered_cols = [c for c in sort_ref if c in all_cols] + [c for c in all_cols if c not in sort_ref]
            pivot_full = pivot_full.reindex(columns=ordered_cols, fill_value=0)
            pivot_full['합계'] = pivot_full.sum(axis=1)
            pivot_diff = pivot_full.diff()
            pivot_pct = pivot_full.pct_change() * 100.0
            if is_partial and len(valid_years) >= 2:
                prev_year_val = valid_years[-2]
                df_same = filtered_df[filtered_df['월'].isin(last_year_months)]
                pivot_partial = df_same.pivot_table(index='기준연도', columns='중분류', values=target_col, aggfunc='sum').fillna(0)
                pivot_partial = pivot_partial.reindex(columns=ordered_cols, fill_value=0)
                pivot_partial['합계'] = pivot_partial.sum(axis=1)
                curr_ytd = pivot_partial.loc[last_year]
                prev_ytd = pivot_partial.loc[prev_year_val] if prev_year_val in pivot_partial.index else pd.Series(0, index=pivot_full.columns)
                diff_ytd = curr_ytd - prev_ytd
                pct_ytd = (diff_ytd / prev_ytd.replace(0, np.nan)) * 100.0
                pivot_full = pivot_full.rename(index={last_year: partial_label})
                pivot_diff = pivot_diff.rename(index={last_year: partial_label})
                pivot_pct = pivot_pct.rename(index={last_year: partial_label})
                pivot_diff.loc[partial_label] = diff_ytd
                pivot_pct.loc[partial_label] = pct_ytd
            pivot_base = pivot_full
            date_index_col = '연도'
        else:
            pivot_base = filtered_df.pivot_table(index='기준년월', columns='중분류', values=target_col, aggfunc='sum').fillna(0)
            all_cols = pivot_base.columns.tolist()
            sort_ref = [c for c in waste_order if c != '전체'] if selected_cat == '폐지' else ([c for c in pulp_order if c != '전체'] if selected_cat == '펄프' else [c for c in liner_order if c != '전체'])
            ordered_cols = [c for c in sort_ref if c in all_cols] + [c for c in all_cols if c not in sort_ref]
            pivot_base = pivot_base.reindex(columns=ordered_cols, fill_value=0)
            pivot_base['합계'] = pivot_base.sum(axis=1)
            pivot_diff = pivot_base.diff()
            pivot_pct = pivot_base.pct_change() * 100.0
            date_index_col = '기준년월'

        final_data = {}
        for col in pivot_base.columns:
            final_data[(col, '중량(톤)')] = pivot_base[col]
            final_data[(col, '증감량')] = pivot_diff[col]
            final_data[(col, '증감률(%)')] = pivot_pct[col]
        pivot_final = pd.DataFrame(final_data, index=pivot_base.index)

    st.markdown(f"### 📋 {display_item_title} {trade_type} 실적{title_country_str}")
    col_info, col_btn_excel, col_btn_print = st.columns([3.2, 1.1, 0.9])
    with col_info:
        st.caption(f"(단위 : 톤) | 조회범위: {selected_desc}")

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        pivot_final.to_excel(writer, sheet_name="통관실적")
    excel_data = excel_buffer.getvalue()

    with col_btn_excel:
        st.download_button(
            label="📥 엑셀 다운로드",
            data=excel_data,
            file_name=f"{display_item_title}_{trade_type}_통관실적.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_btn_print:
        components.html("""
            <style>
                body { margin: 0; padding: 0; }
                .print-btn {
                    width: 100%; height: 38px; background-color: #FFFFFF; color: #1E3A8A;
                    border: 1.5px solid #94A3B8; border-radius: 8px; font-weight: 600; font-size: 14px;
                    cursor: pointer; display: flex; align-items: center; justify-content: center;
                }
                .print-btn:hover { background-color: #EFF6FF; border-color: #1E3A8A; }
            </style>
            <button class="print-btn" onclick="window.parent.print()">🖨️ 표 인쇄</button>
        """, height=40)

    top_headers = []
    for col in pivot_final.columns:
        if col[0] not in top_headers:
            top_headers.append(col[0])

    html = ['<div class="custom-table-container"><table class="custom-table">']
    html.append('<thead><tr>')
    index_header_name = "기준연도" if period_mode == "연간" else "기준년월"
    html.append(f'<th rowspan="2" style="vertical-align: middle;">{index_header_name}</th>')
    for top_h in top_headers:
        sub_count = sum(1 for c in pivot_final.columns if c[0] == top_h)
        html.append(f'<th colspan="{sub_count}">{top_h}</th>')
    html.append('</tr><tr>')
    for col in pivot_final.columns:
        html.append(f'<th>{col[1]}</th>')
    html.append('</tr></thead><tbody>')

    for idx, row in pivot_final.iterrows():
        html.append('<tr>')
        html.append(f'<td class="col-label" style="font-weight: 700; background-color: #F1F5F9;">{idx}</td>')
        for col in pivot_final.columns:
            val = row[col]
            sub_type = col[1]
            if pd.isna(val):
                display_val = "-"
                css_class = "val-zero"
            elif "증감률" in sub_type:
                display_val = f"{val:,.1f}%"
                css_class = "val-negative" if val < 0 else ""
            else:
                display_val = f"{int(round(val)):,}"
                css_class = "val-negative" if val < 0 else ""
            html.append(f'<td class="col-num {css_class}">{display_val}</td>')
        html.append('</tr>')

    html.append('</tbody></table></div>')
    st.markdown("".join(html), unsafe_allow_html=True)
    st.caption("※ 자료출처 : 통계청")

    # 하단 차트
    chart_html = [
        '<div class="chart-box">',
        '<hr style="margin: 25px 0; border: none; border-top: 1px solid #94A3B8;">',
        f'<h3 style="margin-bottom: 20px;">📊 {display_item_title} 추이 차트{title_country_str}</h3>'
    ]

    if trade_type == "수출+수입":
        chart_html.append('</div>')
        st.markdown("".join(chart_html), unsafe_allow_html=True)
        chart_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01') if period_mode == "월간" else pivot_base.index.astype(str)
        fig_dual = go.Figure()
        fig_dual.add_trace(go.Scatter(x=chart_x, y=pivot_base['수출실적(톤)'], name='수출', mode='lines+markers', line=dict(color='#2E7D32', width=2.5)))
        fig_dual.add_trace(go.Scatter(x=chart_x, y=pivot_base['수입실적(톤)'], name='수입', mode='lines+markers', line=dict(color='#C62828', width=2.5)))
        st.plotly_chart(fig_dual, use_container_width=True)
    elif chosen_sub_item != "전체":
        chart_html.append('</div>')
        st.markdown("".join(chart_html), unsafe_allow_html=True)
        chart_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01') if period_mode == "월간" else pivot_base.index.astype(str)
        val_col = '수출실적(톤)' if trade_type == '수출' else '수입실적(톤)'
        fig_single = go.Figure()
        fig_single.add_trace(go.Scatter(x=chart_x, y=pivot_base[val_col], name=f'{trade_type}실적', mode='lines+markers', line=dict(color='#2E7D32' if trade_type=='수출' else '#C62828', width=2.5)))
        st.plotly_chart(fig_single, use_container_width=True)
    else:
        chart_html.append('</div>')
        st.markdown("".join(chart_html), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="chart-box"><h5 style="margin-bottom: 10px;">📈 세부 품목별 실적 (톤)</h5></div>', unsafe_allow_html=True)
            item_cols = [c for c in pivot_base.columns if c != '합계']
            chart_line_df = pivot_base[item_cols].reset_index()
            chart_line_df.rename(columns={chart_line_df.columns[0]: date_index_col}, inplace=True)
            x_col = '날짜' if period_mode == "월간" else date_index_col
            if period_mode == "월간":
                chart_line_df['날짜'] = pd.to_datetime(chart_line_df[date_index_col].astype(str).str.replace('.', '-') + '-01')
            chart_line_df = chart_line_df.melt(id_vars=x_col, value_vars=item_cols, var_name='품목', value_name='실적(톤)')
            fig_line = px.line(chart_line_df, x=x_col, y='실적(톤)', color='품목', markers=True)
            st.plotly_chart(fig_line, use_container_width=True)
        with c2:
            st.markdown('<div class="chart-box"><h5 style="margin-bottom: 10px;">🏛️ 합계 실적 및 증감량 (톤)</h5></div>', unsafe_allow_html=True)
            bar_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01') if period_mode == "월간" else pivot_base.index.astype(str)
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(x=bar_x, y=pivot_base['합계'], name='총 실적', marker_color='#4A90E2'))
            fig_bar.add_trace(go.Scatter(x=bar_x, y=pivot_diff['합계'], name='증감량', mode='lines+markers', line=dict(color='#E2594A', width=2), yaxis='y2'))
            fig_bar.update_layout(yaxis2=dict(title="증감량", overlaying='y', side='right', showgrid=False))
            st.plotly_chart(fig_bar, use_container_width=True)

# ==========================================
# 2. 종이판지 수출입 통관실적 대시보드
# ==========================================
else:
    paper_files = glob.glob(os.path.join(DATA_DIR, "*종이판지*수출입통계*.xlsx")) + glob.glob(os.path.join(DATA_DIR, "*지류*수출입통계*.xlsx"))
    if not paper_files:
        st.error("종이판지 수출입 통계 엑셀 파일을 찾을 수 없습니다. (파일명: 종이판지_수출입통계_update.xlsx)")
        st.stop()
    latest_paper_file = max(paper_files, key=os.path.getctime)

    @st.cache_data
    def load_paper_data(path):
        xls = pd.ExcelFile(path)
        sheet_target = "종이판지" if "종이판지" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(path, sheet_name=sheet_target)
        df['연도'] = df['연도'].astype(int)
        df['중량(톤)'] = pd.to_numeric(df['중량(톤)'], errors='coerce').fillna(0)
        return df

    df_p = load_paper_data(latest_paper_file)

    st.title("종이판지 수출입 통관실적")
    st.write("---")

    # 사이드바 품목 선택
    st.sidebar.markdown("### 📂 품목 선택")
    paper_cat_list = [
        "전체 품목 (종합표)",
        "신문용지", "백상지", "비도공 기타", "아트지", "도공 기타", "박엽인쇄용지",
        "감열기록지", "복사용지", "전산용지", "팬시지", "권련지", "위생용지", "중포대용크라프트지",
        "도공 카톤", "도공 SC", "도공 아이보리", "비도공 카톤", "비도공 TM",
        "라이너", "골심지", "밀크카톤", "컵원지", "판지 기타",
        "골판상자", "지대", "감열기록지(제품)", "카본지"
    ]
    selected_paper_item = st.sidebar.selectbox("지종 선택", paper_cat_list, index=0)

    st.sidebar.write("---")
    st.sidebar.markdown("### 🔄 무역 구분")
    paper_trade = st.sidebar.radio("구분", ["수출", "수입"])

    # 상단 기간 설정
    col_view_type, col_p1, col_p2 = st.columns([1.2, 1.4, 1.4])
    with col_view_type:
        view_type = st.radio("📅 **조회기간**", ["연간", "월간"], horizontal=True)

    df_trade = df_p[df_p['수출/수입'] == paper_trade].copy()

    # ----------------------------------------------------
    # Case A. 단일 지종 선택 시 (원료 표처럼 증감량/증감률 표시)
    # ----------------------------------------------------
    if selected_paper_item != "전체 품목 (종합표)":
        target_item_df = df_trade[df_trade['지종'] == selected_paper_item].copy()

        all_years = sorted(target_item_df['연도'].unique().tolist())
        df_monthly_sub = target_item_df[target_item_df['월'] != '누계'].copy()
        df_monthly_sub['월'] = df_monthly_sub['월'].astype(int)
        df_monthly_sub['년월'] = df_monthly_sub['연도'].astype(str) + "." + df_monthly_sub['월'].apply(lambda x: f"{x:02d}")
        all_ym = sorted(df_monthly_sub['년월'].unique().tolist())

        if view_type == "연간":
            with col_p1:
                s_year = st.selectbox("시작 연도", all_years, index=0)
            with col_p2:
                valid_e_years = [y for y in all_years if y >= s_year]
                e_year = st.selectbox("종료 연도", valid_e_years, index=len(valid_e_years)-1)
            desc_text = f"{s_year}년 ~ {e_year}년 (연간)"

            # 연간 수치 계산
            year_records = {}
            for y in range(s_year, e_year + 1):
                sub_y = target_item_df[target_item_df['연도'] == y]
                if '누계' in sub_y['월'].values:
                    val = sub_y[sub_y['월'] == '누계']['중량(톤)'].sum()
                else:
                    val = sub_y['중량(톤)'].sum()
                year_records[y] = val

            p_years = sorted(year_records.keys())
            last_y = p_years[-1]
            last_y_months = sorted(target_item_df[(target_item_df['연도'] == last_y) & (target_item_df['월'] != '누계')]['월'].unique().tolist())
            is_partial_p = (len(last_y_months) < 12 and len(last_y_months) > 0)
            partial_lbl = f"{last_y}.{last_y_months[0]}-{last_y_months[-1]}" if is_partial_p else str(last_y)

            rows_calc = []
            for i, y in enumerate(p_years):
                curr_v = year_records[y]
                if i == 0:
                    diff_v, rate_v = np.nan, np.nan
                else:
                    prev_v = year_records[p_years[i-1]]
                    # 최신연도가 부분 누계인 경우 전년 동기 누계와 비교
                    if y == last_y and is_partial_p:
                        prev_y_same_months = target_item_df[(target_item_df['연도'] == p_years[i-1]) & (target_item_df['월'].isin(last_y_months))]['중량(톤)'].sum()
                        if prev_y_same_months > 0:
                            diff_v = curr_v - prev_y_same_months
                            rate_v = (diff_v / prev_y_same_months) * 100.0
                        else:
                            diff_v, rate_v = curr_v - prev_v, ((curr_v - prev_v) / prev_v * 100.0) if prev_v > 0 else np.nan
                    else:
                        diff_v = curr_v - prev_v
                        rate_v = (diff_v / prev_v * 100.0) if prev_v > 0 else np.nan

                row_name = partial_lbl if (y == last_y and is_partial_p) else str(y)
                rows_calc.append({'기준': row_name, '중량(톤)': curr_v, '증감량': diff_v, '증감률(%)': rate_v})

            pivot_single = pd.DataFrame(rows_calc).set_index('기준')
            idx_name = "기준연도"

        else: # 월간
            with col_p1:
                s_ym = st.selectbox("시작 연월", all_ym, index=max(0, len(all_ym)-24))
            with col_p2:
                valid_e_ym = [ym for ym in all_ym if ym >= s_ym]
                e_ym = st.selectbox("종료 연월", valid_e_ym, index=len(valid_e_ym)-1)
            desc_text = f"{s_ym} ~ {e_ym} (월별)"

            target_ym_list = [ym for ym in all_ym if s_ym <= ym <= e_ym]
            df_target_m = df_monthly_sub[df_monthly_sub['년월'].isin(target_ym_list)].sort_values('년월')
            pivot_m = df_target_m.groupby('년월')['중량(톤)'].sum().reindex(target_ym_list).fillna(0)
            
            diff_m = pivot_m.diff()
            rate_m = pivot_m.pct_change() * 100.0
            pivot_single = pd.DataFrame({'중량(톤)': pivot_m, '증감량': diff_m, '증감률(%)': rate_m})
            idx_name = "기준년월"

        st.markdown(f"### 📋 {selected_paper_item} {paper_trade} 실적")
        col_info, col_btn_excel, col_btn_print = st.columns([3.2, 1.1, 0.9])
        with col_info:
            st.caption(f"(단위 : 톤) | 조회 범위: {desc_text}")

        # 엑셀 다운로드
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            pivot_single.to_excel(writer, sheet_name=selected_paper_item)
        excel_data = excel_buffer.getvalue()

        with col_btn_excel:
            st.download_button(
                label="📥 엑셀 다운로드",
                data=excel_data,
                file_name=f"{selected_paper_item}_{paper_trade}실적_{desc_text}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_btn_print:
            components.html("""
                <style>
                    body { margin: 0; padding: 0; }
                    .print-btn {
                        width: 100%; height: 38px; background-color: #FFFFFF; color: #1E3A8A;
                        border: 1.5px solid #94A3B8; border-radius: 8px; font-weight: 600; font-size: 14px;
                        cursor: pointer; display: flex; align-items: center; justify-content: center;
                    }
                    .print-btn:hover { background-color: #EFF6FF; border-color: #1E3A8A; }
                </style>
                <button class="print-btn" onclick="window.parent.print()">🖨️ 표 인쇄</button>
            """, height=40)

        # HTML 표 렌더링 (원료표 형식)
        html = ['<div class="custom-table-container"><table class="custom-table">']
        html.append('<thead><tr>')
        html.append(f'<th rowspan="2" style="vertical-align: middle; width: 140px;">{idx_name}</th>')
        html.append(f'<th colspan="3">{selected_paper_item}</th>')
        html.append('</tr><tr>')
        html.append('<th>중량(톤)</th><th>증감량</th><th>증감률(%)</th>')
        html.append('</tr></thead><tbody>')

        for idx_row, row in pivot_single.iterrows():
            html.append('<tr>')
            html.append(f'<td class="col-label" style="font-weight: 700; background-color: #F1F5F9;">{idx_row}</td>')
            
            val_w = row['중량(톤)']
            disp_w = f"{int(round(val_w)):,}" if pd.notna(val_w) else "-"
            html.append(f'<td class="col-num">{disp_w}</td>')

            val_d = row['증감량']
            if pd.isna(val_d):
                disp_d, cls_d = "-", "val-zero"
            else:
                disp_d = f"{int(round(val_d)):,}"
                cls_d = "val-negative" if val_d < 0 else ""
            html.append(f'<td class="col-num {cls_d}">{disp_d}</td>')

            val_r = row['증감률(%)']
            if pd.isna(val_r):
                disp_r, cls_r = "-", "val-zero"
            else:
                disp_r = f"{val_r:,.1f}%"
                cls_r = "val-negative" if val_r < 0 else ""
            html.append(f'<td class="col-num {cls_r}">{disp_r}</td>')
            html.append('</tr>')

        html.append('</tbody></table></div>')
        st.markdown("".join(html), unsafe_allow_html=True)
        st.caption("※ 자료출처 : 관세청 통관통계")

        # 단일 품목 추이 차트
        st.write("---")
        st.markdown(f'<div class="chart-box"><h3>📊 {selected_paper_item} {paper_trade} 추이 차트</h3></div>', unsafe_allow_html=True)
        
        chart_x = pivot_single.index.astype(str)
        fig_single = go.Figure()
        fig_single.add_trace(go.Scatter(
            x=chart_x, y=pivot_single['중량(톤)'],
            name=f'{selected_paper_item} 실적(톤)',
            mode='lines+markers',
            line=dict(color='#1E3A8A', width=2.5)
        ))
        fig_single.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            yaxis=dict(title="실적(톤)", tickformat=",.0f")
        )
        st.plotly_chart(fig_single, use_container_width=True)

    # ----------------------------------------------------
    # Case B. 전체 품목 (종합표) 모드
    # ----------------------------------------------------
    else:
        if view_type == "연간":
            all_years = sorted(df_trade['연도'].unique().tolist())
            with col_p1:
                s_year = st.selectbox("시작 연도", all_years, index=0)
            with col_p2:
                valid_e_years = [y for y in all_years if y >= s_year]
                e_year = st.selectbox("종료 연도", valid_e_years, index=len(valid_e_years)-1)
            
            target_columns = [y for y in all_years if s_year <= y <= e_year]
            col_headers = [f"{y}년" for y in target_columns]
            desc_text = f"{s_year}년 ~ {e_year}년 (연간)"
        else:
            df_monthly = df_trade[df_trade['월'] != '누계'].copy()
            df_monthly['월'] = df_monthly['월'].astype(int)
            df_monthly['년월'] = df_monthly['연도'].astype(str) + "." + df_monthly['월'].apply(lambda x: f"{x:02d}")
            all_ym = sorted(df_monthly['년월'].unique().tolist())

            with col_p1:
                s_ym = st.selectbox("시작 연월", all_ym, index=max(0, len(all_ym)-12))
            with col_p2:
                valid_e_ym = [ym for ym in all_ym if ym >= s_ym]
                e_ym = st.selectbox("종료 연월", valid_e_ym, index=len(valid_e_ym)-1)

            target_columns = [ym for ym in all_ym if s_ym <= ym <= e_ym]
            col_headers = target_columns
            desc_text = f"{s_ym} ~ {e_ym} (월별)"

        # 품목 계층 마스터 테이블
        items_tree = [
            ([('종이', 18, 1, 'font-weight: 700; vertical-align: middle; width: 60px;'), ('신문용지', 1, 2, 'text-align: center; width: 150px;')], '신문용지', ''),
            ([('비도공<br>인쇄용지', 3, 1, 'vertical-align: middle; width: 85px;'), ('백상지', 1, 1, 'width: 65px;')], '백상지', ''),
            ([('기타', 1, 1, '')], '비도공 기타', ''),
            ([('계', 1, 1, '')], '__CALC_비도공계__', 'row-subtotal'),
            ([('도공<br>인쇄용지', 3, 1, 'vertical-align: middle; width: 85px;'), ('아트지류', 1, 1, '')], '아트지', ''),
            ([('기타', 1, 1, '')], '도공 기타', ''),
            ([('계', 1, 1, '')], '__CALC_도공계__', 'row-subtotal'),
            ([('박엽인쇄용지', 1, 2, 'text-align: center;')], '박엽인쇄용지', ''),
            ([('정보<br>인쇄용지', 4, 1, 'vertical-align: middle; width: 85px;'), ('감열기록지', 1, 1, '')], '감열기록지', ''),
            ([('복사용지', 1, 1, '')], '복사용지', ''),
            ([('전산용지', 1, 1, '')], '전산용지', ''),
            ([('계', 1, 1, '')], '__CALC_정보계__', 'row-subtotal'),
            ([('인쇄용지소계', 1, 2, 'text-align: center;')], '__CALC_인쇄용지소계__', 'row-subtotal'),
            ([('기타<br>특수지', 3, 1, 'vertical-align: middle; width: 85px;'), ('팬시지', 1, 1, '')], '팬시지', ''),
            ([('권련지', 1, 1, '')], '권련지', ''),
            ([('계', 1, 1, '')], '__CALC_특수지계__', 'row-subtotal'),
            ([('위생용지', 1, 2, 'text-align: center;')], '위생용지', ''),
            ([('중포대용크라프트지', 1, 2, 'text-align: center;')], '중포대용크라프트지', ''),
            ([('종이합계', 1, 3, 'text-align: center; font-weight: 800;')], '__CALC_종이합계__', 'row-total'),

            ([('판지', 15, 1, 'font-weight: 700; vertical-align: middle; width: 60px;'), ('백판지(도공)', 4, 1, 'vertical-align: middle; width: 85px;'), ('카톤용', 1, 1, 'width: 65px;')], '도공 카톤', ''),
            ([('SC', 1, 1, '')], '도공 SC', ''),
            ([('아이보리', 1, 1, '')], '도공 아이보리', ''),
            ([('계', 1, 1, '')], '__CALC_백판지도공계__', 'row-subtotal'),
            ([('백판지(비도공)', 3, 1, 'vertical-align: middle; width: 85px;'), ('카톤용', 1, 1, '')], '비도공 카톤', ''),
            ([('TM', 1, 1, '')], '비도공 TM', ''),
            ([('계', 1, 1, '')], '__CALC_백판지비도공계__', 'row-subtotal'),
            ([('백판지 계', 1, 2, 'text-align: center;')], '__CALC_백판지계__', 'row-subtotal'),
            ([('골판지원지', 3, 1, 'vertical-align: middle; width: 85px;'), ('라이너', 1, 1, '')], '라이너', ''),
            ([('골심지', 1, 1, '')], '골심지', ''),
            ([('계', 1, 1, '')], '__CALC_골판지원지계__', 'row-subtotal'),
            ([('기타판지', 4, 1, 'vertical-align: middle; width: 85px;'), ('밀크카톤등', 1, 1, '')], '밀크카톤', ''),
            ([('컵원지,접시등', 1, 1, '')], '컵원지', ''),
            ([('기타', 1, 1, '')], '판지 기타', ''),
            ([('계', 1, 1, '')], '__CALC_기타판지계__', 'row-subtotal'),
            ([('판지합계', 1, 3, 'text-align: center; font-weight: 800;')], '__CALC_판지합계__', 'row-total'),

            ([('종이판지합계', 1, 3, 'text-align: center; font-weight: 800;')], '__CALC_종이판지합계__', 'row-grand-total'),

            ([('종이<br>제품', 4, 1, 'font-weight: 700; vertical-align: middle; width: 60px;'), ('골판상자', 1, 2, 'text-align: center;')], '골판상자', ''),
            ([('지대', 1, 2, 'text-align: center;')], '지대', ''),
            ([('감열기록지', 1, 2, 'text-align: center;')], '감열기록지(제품)', ''),
            ([('카본지 또는 유사한 복사지', 1, 2, 'text-align: center;')], '카본지', ''),
            ([('지제품합계', 1, 3, 'text-align: center; font-weight: 800;')], '__CALC_지제품합계__', 'row-total'),

            ([('총합계', 1, 3, 'text-align: center; font-weight: 800;')], '__CALC_총합계__', 'row-grand-total')
        ]

        val_map = {item: {col: 0.0 for col in target_columns} for item in df_trade['지종'].unique()}

        if view_type == "연간":
            for col_y in target_columns:
                sub = df_trade[df_trade['연도'] == col_y]
                if '누계' in sub['월'].values:
                    sub_agg = sub[sub['월'] == '누계'].groupby('지종')['중량(톤)'].sum()
                else:
                    sub_agg = sub.groupby('지종')['중량(톤)'].sum()
                for it, val in sub_agg.items():
                    if it in val_map:
                        val_map[it][col_y] = val
        else:
            for col_ym in target_columns:
                y, m = map(int, col_ym.split('.'))
                sub = df_trade[(df_trade['연도'] == y) & (df_trade['월'] == m)]
                sub_agg = sub.groupby('지종')['중량(톤)'].sum()
                for it, val in sub_agg.items():
                    if it in val_map:
                        val_map[it][col_ym] = val

        def get_sum_dict(items):
            res = {c: 0.0 for c in target_columns}
            for it in items:
                if it in val_map:
                    for c in target_columns:
                        res[c] += val_map[it][c]
            return res

        val_map['__CALC_비도공계__'] = get_sum_dict(['백상지', '비도공 기타'])
        val_map['__CALC_도공계__'] = get_sum_dict(['아트지', '도공 기타'])
        val_map['__CALC_정보계__'] = get_sum_dict(['감열기록지', '복사용지', '전산용지'])
        val_map['__CALC_특수지계__'] = get_sum_dict(['팬시지', '권련지'])
        val_map['__CALC_인쇄용지소계__'] = get_sum_dict(['백상지', '비도공 기타', '아트지', '도공 기타', '박엽인쇄용지', '감열기록지', '복사용지', '전산용지'])
        val_map['__CALC_종이합계__'] = get_sum_dict(['신문용지', '백상지', '비도공 기타', '아트지', '도공 기타', '박엽인쇄용지', '감열기록지', '복사용지', '전산용지', '팬시지', '권련지', '위생용지', '중포대용크라프트지'])

        val_map['__CALC_백판지도공계__'] = get_sum_dict(['도공 카톤', '도공 SC', '도공 아이보리'])
        val_map['__CALC_백판지비도공계__'] = get_sum_dict(['비도공 카톤', '비도공 TM'])
        val_map['__CALC_백판지계__'] = get_sum_dict(['도공 카톤', '도공 SC', '도공 아이보리', '비도공 카톤', '비도공 TM'])
        val_map['__CALC_골판지원지계__'] = get_sum_dict(['라이너', '골심지'])
        val_map['__CALC_기타판지계__'] = get_sum_dict(['밀크카톤', '컵원지', '판지 기타'])
        val_map['__CALC_판지합계__'] = get_sum_dict(['도공 카톤', '도공 SC', '도공 아이보리', '비도공 카톤', '비도공 TM', '라이너', '골심지', '밀크카톤', '컵원지', '판지 기타'])

        val_map['__CALC_종이판지합계__'] = {c: val_map['__CALC_종이합계__'][c] + val_map['__CALC_판지합계__'][c] for c in target_columns}
        val_map['__CALC_지제품합계__'] = get_sum_dict(['골판상자', '지대', '감열기록지(제품)', '카본지'])
        val_map['__CALC_총합계__'] = {c: val_map['__CALC_종이판지합계__'][c] + val_map['__CALC_지제품합계__'][c] for c in target_columns}

        st.markdown(f"### 📋 종이·판지 {paper_trade} 실적 (종합표)")
        col_info, col_btn_excel, col_btn_print = st.columns([3.2, 1.1, 0.9])
        with col_info:
            st.caption(f"(단위 : 톤) | 조회 범위: {desc_text}")

        excel_rows = []
        for cells, code, _ in items_tree:
            row_dict = {'구분': code}
            for c, h in zip(target_columns, col_headers):
                row_dict[h] = round(val_map.get(code, {}).get(c, 0))
            excel_rows.append(row_dict)
        paper_excel_df = pd.DataFrame(excel_rows)

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            paper_excel_df.to_excel(writer, sheet_name="종이판지실적", index=False)
        excel_data = excel_buffer.getvalue()

        with col_btn_excel:
            st.download_button(
                label="📥 엑셀 다운로드",
                data=excel_data,
                file_name=f"종이판지_{paper_trade}실적_{desc_text}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_btn_print:
            components.html("""
                <style>
                    body { margin: 0; padding: 0; }
                    .print-btn {
                        width: 100%; height: 38px; background-color: #FFFFFF; color: #1E3A8A;
                        border: 1.5px solid #94A3B8; border-radius: 8px; font-weight: 600; font-size: 14px;
                        cursor: pointer; display: flex; align-items: center; justify-content: center;
                    }
                    .print-btn:hover { background-color: #EFF6FF; border-color: #1E3A8A; }
                </style>
                <button class="print-btn" onclick="window.parent.print()">🖨️ 표 인쇄</button>
            """, height=40)

        html = ['<div class="custom-table-container"><table class="custom-table">']
        html.append('<thead><tr>')
        html.append('<th colspan="3" style="vertical-align: middle; width: 220px;">구 분</th>')
        for h in col_headers:
            html.append(f'<th>{h}</th>')
        html.append('</tr></thead><tbody>')

        for cells, code, row_cls in items_tree:
            html.append(f'<tr class="{row_cls}">')
            for text, rspan, cspan, style in cells:
                attr_r = f' rowspan="{rspan}"' if rspan > 1 else ''
                attr_c = f' colspan="{cspan}"' if cspan > 1 else ''
                attr_s = f' style="{style}"' if style else ''
                html.append(f'<td class="col-label"{attr_r}{attr_c}{attr_s}>{text}</td>')

            for c in target_columns:
                val = val_map.get(code, {}).get(c, 0.0)
                disp = f"{int(round(val)):,}" if val > 0 else "-"
                html.append(f'<td class="col-num">{disp}</td>')
            html.append('</tr>')

        html.append('</tbody></table></div>')
        st.markdown("".join(html), unsafe_allow_html=True)
        st.caption("※ 자료출처 : 관세청 통관통계")