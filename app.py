import streamlit as st
import pandas as pd
import numpy as np
import os
import glob
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="원료 수출입 통관실적", layout="wide")

# 인쇄 시 불필요 영역 숨김 스타일
st.markdown("""
<style>
@media print {
    [data-testid="stSidebar"], header, .stPlotlyChart, [data-testid="stRadio"], [data-testid="stSelectbox"] {
        display: none !important;
    }
    .block-container {
        padding: 0 !important;
    }
}
</style>
""", unsafe_allow_html=True)

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

# 화면 상단 대제목
st.title("원료 수출입 통관실적")
st.write("---")

# 품목별 순서 사전 정의
waste_order = ['전체', '폐골판지', '폐신문지', '고급폐지', '기타폐지']
pulp_order = ['전체', 'GP', 'DP', 'UKP', 'BKP', 'BCTMP', '면린터펄프', 'DIP', '기타']
liner_order = ['전체', '라이너', '골심지']

# 2. 사이드바 계층 UI 구성 (대분류 -> 세부 품목 -> 수출/수입 -> 국가)
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

# 선택된 세부 품목 기준 국가 목록 및 상위국 집계
if chosen_sub_item != "전체":
    country_calc_df = cat_df[cat_df['중분류'] == chosen_sub_item]
else:
    country_calc_df = cat_df

all_countries = sorted(country_calc_df['국가명'].dropna().unique().tolist())
top10_countries = (
    country_calc_df.groupby('국가명')[['수출실적(톤)', '수입실적(톤)']]
    .sum().sum(axis=1)
    .sort_values(ascending=False)
    .head(10)
    .index.tolist()
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

# 3. 상단 설정 영역: [기간 선택 (집계주기+기간 드롭다운)] 단독 배치
all_years = sorted([y for y in df['기준연도'].unique() if len(str(y)) == 4 and str(y).isdigit()])
min_year = int(all_years[0])
max_year = int(all_years[-1])

period_options = [
    "연간 전체 (2010년~현재)",
    "최근 5개년 연간 (YoY)",
    "최근 3개년 연간 (YoY)",
    "최근 12개월 (월별 MoM)",
    "최근 24개월 (월별 MoM)",
    "최근 36개월 (월별 MoM)",
    "월별 전체 실적 (MoM)"
]
selected_period_opt = st.selectbox("📅 **조회 기간 및 집계 방식**", period_options, index=0)

# 4. 데이터 필터링 로직
filtered_df = cat_df.copy()

if chosen_sub_item != "전체":
    filtered_df = filtered_df[filtered_df['중분류'] == chosen_sub_item]
    display_item_title = chosen_sub_item
else:
    display_item_title = f"{selected_cat} 전체"

if selected_country != "전체 국가":
    filtered_df = filtered_df[filtered_df['국가명'] == selected_country]
    title_country_str = f"[{selected_country}]"
else:
    title_country_str = "[전체 국가]"

# 기간 및 주기 파싱
if "연간" in selected_period_opt:
    period_type = "연간 합계 (YoY)"
    if "최근 5개년" in selected_period_opt:
        start_y = str(max(min_year, max_year - 4))
        filtered_df = filtered_df[filtered_df['기준연도'] >= start_y]
    elif "최근 3개년" in selected_period_opt:
        start_y = str(max(min_year, max_year - 2))
        filtered_df = filtered_df[filtered_df['기준연도'] >= start_y]
else:
    period_type = "월별 실적 (MoM)"
    all_dates = sorted(df['기준년월'].unique().tolist())
    if "최근 12개월" in selected_period_opt:
        target_dates = all_dates[-12:]
        filtered_df = filtered_df[filtered_df['기준년월'].isin(target_dates)]
    elif "최근 24개월" in selected_period_opt:
        target_dates = all_dates[-24:]
        filtered_df = filtered_df[filtered_df['기준년월'].isin(target_dates)]
    elif "최근 36개월" in selected_period_opt:
        target_dates = all_dates[-36:]
        filtered_df = filtered_df[filtered_df['기준년월'].isin(target_dates)]

if filtered_df.empty:
    st.warning("선택된 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# 5. 피벗 및 연산
valid_years = sorted([y for y in filtered_df['기준연도'].unique() if len(str(y)) == 4 and str(y).isdigit()])
last_year = valid_years[-1]
last_year_months = sorted([m for m in filtered_df[filtered_df['기준연도'] == last_year]['월'].unique() if 1 <= m <= 12])
is_partial = (len(last_year_months) < 12 and len(last_year_months) > 0 and "연간" in period_type)

partial_label = f"{last_year}.{last_year_months[0]}-{last_year_months[-1]}" if is_partial else last_year

if trade_type == "수출+수입" or chosen_sub_item != "전체":
    if "연간" in period_type:
        pivot_full = filtered_df.pivot_table(index='기준연도', values=['수출실적(톤)', '수입실적(톤)'], aggfunc='sum').fillna(0)
        if is_partial:
            df_same = filtered_df[filtered_df['월'].isin(last_year_months)]
            pivot_partial = df_same.pivot_table(index='기준연도', values=['수출실적(톤)', '수입실적(톤)'], aggfunc='sum').fillna(0)
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

        pivot_base = pivot_full
        date_index_col = '연도'
    else:
        pivot_base = filtered_df.pivot_table(index='기준년월', values=['수출실적(톤)', '수입실적(톤)'], aggfunc='sum').fillna(0)
        pivot_diff = pivot_base.diff()
        pivot_pct = pivot_base.pct_change() * 100.0
        date_index_col = '기준년월'

    if trade_type == "수출":
        final_data = {
            ('수출', '실적'): pivot_base['수출실적(톤)'],
            ('수출', '증감량'): pivot_diff['수출실적(톤)'],
            ('수출', '증감률'): pivot_pct['수출실적(톤)']
        }
    elif trade_type == "수입":
        final_data = {
            ('수입', '실적'): pivot_base['수입실적(톤)'],
            ('수입', '증감량'): pivot_diff['수입실적(톤)'],
            ('수입', '증감률'): pivot_pct['수입실적(톤)']
        }
    else:
        final_data = {
            ('수출', '실적'): pivot_base['수출실적(톤)'],
            ('수출', '증감량'): pivot_diff['수출실적(톤)'],
            ('수출', '증감률'): pivot_pct['수출실적(톤)'],
            ('수입', '실적'): pivot_base['수입실적(톤)'],
            ('수입', '증감량'): pivot_diff['수입실적(톤)'],
            ('수입', '증감률'): pivot_pct['수입실적(톤)']
        }
    pivot_final = pd.DataFrame(final_data, index=pivot_base.index)

else:
    target_col = '수출실적(톤)' if trade_type == '수출' else '수입실적(톤)'
    
    if "연간" in period_type:
        pivot_full = filtered_df.pivot_table(index='기준연도', columns='중분류', values=target_col, aggfunc='sum').fillna(0)
        if is_partial:
            df_same = filtered_df[filtered_df['월'].isin(last_year_months)]
            pivot_partial = df_same.pivot_table(index='기준연도', columns='중분류', values=target_col, aggfunc='sum').fillna(0)
        else:
            pivot_partial = pivot_full.copy()

        all_cols = pivot_full.columns.tolist()
        if selected_cat == '폐지':
            sort_ref = [c for c in waste_order if c != '전체']
        elif selected_cat == '펄프':
            sort_ref = [c for c in pulp_order if c != '전체']
        else:
            sort_ref = [c for c in liner_order if c != '전체']

        ordered_cols = [c for c in sort_ref if c in all_cols] + [c for c in all_cols if c not in sort_ref]
        pivot_full = pivot_full.reindex(columns=ordered_cols, fill_value=0)
        pivot_partial = pivot_partial.reindex(columns=ordered_cols, fill_value=0)
        pivot_full['합계'] = pivot_full.sum(axis=1)
        pivot_partial['합계'] = pivot_partial.sum(axis=1)

        pivot_diff = pivot_full.diff()
        pivot_pct = pivot_full.pct_change() * 100.0

        if is_partial and len(valid_years) >= 2:
            prev_year = valid_years[-2]
            curr_ytd = pivot_partial.loc[last_year]
            prev_ytd = pivot_partial.loc[prev_year] if prev_year in pivot_partial.index else pd.Series(0, index=pivot_full.columns)
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
        if selected_cat == '폐지':
            sort_ref = [c for c in waste_order if c != '전체']
        elif selected_cat == '펄프':
            sort_ref = [c for c in pulp_order if c != '전체']
        else:
            sort_ref = [c for c in liner_order if c != '전체']

        ordered_cols = [c for c in sort_ref if c in all_cols] + [c for c in all_cols if c not in sort_ref]
        pivot_base = pivot_base.reindex(columns=ordered_cols, fill_value=0)
        pivot_base['합계'] = pivot_base.sum(axis=1)

        pivot_diff = pivot_base.diff()
        pivot_pct = pivot_base.pct_change() * 100.0
        date_index_col = '기준년월'

    final_data = {}
    for col in pivot_base.columns:
        final_data[(col, col if col != '합계' else '합계')] = pivot_base[col]
        final_data[(col, '증감량')] = pivot_diff[col]
        final_data[(col, '증감률')] = pivot_pct[col]

    pivot_final = pd.DataFrame(final_data, index=pivot_base.index)

# 6. 표 출력
st.markdown(f"### 📋 {display_item_title} {title_country_str} {trade_type} 실적")
st.caption(f"(단위 : 톤) | 구분: {selected_period_opt}")

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
st.caption("※ 자료출처 : 통계청")

# 7. 차트 출력
st.write("---")
st.markdown(f"### 📊 {display_item_title} {title_country_str} 추이 차트")

if trade_type == "수출+수입":
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

elif chosen_sub_item != "전체":
    if "월별" in period_type:
        chart_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01')
    else:
        chart_x = pivot_base.index.astype(str)

    val_col = '수출실적(톤)' if trade_type == '수출' else '수입실적(톤)'
    color_code = '#2E7D32' if trade_type == '수출' else '#C62828'

    fig_single = go.Figure()
    fig_single.add_trace(go.Scatter(x=chart_x, y=pivot_base[val_col], name=f'{trade_type}실적(톤)', mode='lines+markers', line=dict(color=color_code, width=2.5)))

    if "월별" in period_type:
        fig_single.update_xaxes(dtick="M3", tickformat="%Y-%m")

    fig_single.update_layout(
        margin=dict(l=20, r=20, t=30, b=20),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="실적(톤)", tickformat=",.0f")
    )
    st.plotly_chart(fig_single, use_container_width=True)

else:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### 📈 세부 품목별 실적 (톤)")
        item_cols = [c for c in pivot_base.columns if c != '합계']
        chart_line_df = pivot_base[item_cols].reset_index()
        chart_line_df.rename(columns={chart_line_df.columns[0]: date_index_col}, inplace=True)

        if "월별" in period_type:
            chart_line_df['날짜'] = pd.to_datetime(chart_line_df[date_index_col].astype(str).str.replace('.', '-') + '-01')
            x_col = '날짜'
        else:
            x_col = date_index_col

        chart_line_df = chart_line_df.melt(id_vars=x_col, value_vars=item_cols, var_name='품목', value_name='실적(톤)')
        fig_line = px.line(chart_line_df, x=x_col, y='실적(톤)', color='품목', markers=True)

        if "월별" in period_type:
            fig_line.update_xaxes(dtick="M3", tickformat="%Y-%m")

        fig_line.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        fig_line.update_yaxes(tickformat=",.0f")
        st.plotly_chart(fig_line, use_container_width=True)

    with c2:
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