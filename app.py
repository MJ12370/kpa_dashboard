import streamlit as st
import pandas as pd
import numpy as np
import os
import glob
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="원료 수출입 통관실적", layout="wide")

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
        
    df['기준년월'] = df['기준년월'].astype(str).str.strip()
    df['기준연도'] = df['기준년월'].str.slice(0, 4)
    df['월'] = df['기준년월'].str.slice(4, 6)
    
    if '중분류' not in df.columns:
        df['중분류'] = df['품목명']
        
    return df

df = load_data(latest_file)

# 2. 상단 필터 UI (대분류 품목 -> 수출/수입 -> 집계주기)
col_cat, col_trade, col_period = st.columns([1.5, 1, 1.2])

with col_cat:
    raw_cats = df['대분류'].dropna().unique().tolist()
    raw_cats = [c for c in raw_cats if c not in ['폐신문', '폐신문지']]
    
    priority_order = ['폐지', '골판지원지', '펄프']
    available_cats = [c for c in priority_order if c in raw_cats] + [c for c in raw_cats if c not in priority_order]

    selected_cat = st.radio(
        "📂 **대분류 품목**",
        available_cats,
        horizontal=True
    )

with col_trade:
    trade_type = st.radio("🔄 **수출 / 수입**", ["수출", "수입"], horizontal=True)

with col_period:
    period_type = st.radio("📅 **집계 주기**", ["연간 합계 (YoY)", "월별 실적 (MoM)"], horizontal=True)

target_col = '수출실적(톤)' if trade_type == '수출' else '수입실적(톤)'

# 3. 사이드바: 국가 필터 및 폐지 상위 10개국 선택
filtered_df = df[df['대분류'] == selected_cat].copy()

st.sidebar.markdown("### 🌐 국가별 상세 조회")

if selected_cat == '폐지':
    top10_countries = (
        filtered_df.groupby('국가명')[target_col]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    country_options = ["전체 합산"] + top10_countries
    
    selected_country = st.sidebar.radio(
        f"🏆 **폐지 {trade_type} 상위 10개국**",
        country_options,
        index=0
    )
    
    if selected_country != "전체 합산":
        filtered_df = filtered_df[filtered_df['국가명'] == selected_country]
        country_title_label = f"[{selected_country}]"
    else:
        country_title_label = "[전체 국가]"
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

# 4. 피벗 및 동일 기간(YTD) 증감 계산
if "연간" in period_type:
    # 전체 데이터 기준 마지막 연도와 해당 연도의 포함 월 파악
    years = sorted(df['기준연도'].unique())
    last_year = years[-1]
    last_year_months = sorted(df[df['기준연도'] == last_year]['월'].unique())
    is_partial = len(last_year_months) < 12
    
    if is_partial:
        start_m = int(last_year_months[0])
        end_m = int(last_year_months[-1])
        partial_label = f"{last_year}.{start_m}-{end_m}"
    else:
        partial_label = last_year

    # 기본 연도별 피벗 테이블 (전체 월 기준 합계)
    pivot_full = filtered_df.pivot_table(
        index='기준연도',
        columns='중분류',
        values=target_col,
        aggfunc='sum'
    ).fillna(0)

    # 마지막 연도와 동일한 월 범위만 필터링한 피벗 테이블
    if is_partial and len(years) >= 2:
        df_same_period = filtered_df[filtered_df['월'].isin(last_year_months)]
        pivot_partial = df_same_period.pivot_table(
            index='기준연도',
            columns='중분류',
            values=target_col,
            aggfunc='sum'
        ).fillna(0)
    else:
        pivot_partial = pivot_full

    # 폐지 선택 시 품목 정렬
    if selected_cat == '폐지':
        custom_waste_order = ['폐골판지', '폐신문지', '고급폐지', '기타폐지']
        ordered_cols = [c for c in custom_waste_order if c in pivot_full.columns]
        remaining_cols = [c for c in pivot_full.columns if c not in custom_waste_order]
        target_item_cols = ordered_cols + remaining_cols
        pivot_full = pivot_full[target_item_cols]
        pivot_partial = pivot_partial[target_item_cols]

    pivot_full['합계'] = pivot_full.sum(axis=1)
    pivot_partial['합계'] = pivot_partial.sum(axis=1)

    # 증감량 및 증감률 계산
    pivot_diff = pivot_full.diff()
    pivot_pct = pivot_full.pct_change() * 100.0

    # 마지막 미완료 연도는 직전 연도의 '동일 기간(YTD)'과 비교하여 증감값 덮어쓰기
    if is_partial and len(years) >= 2:
        prev_year = years[-2]
        if prev_year in pivot_partial.index and last_year in pivot_partial.index:
            ytd_diff = pivot_partial.loc[last_year] - pivot_partial.loc[prev_year]
            ytd_pct = (ytd_diff / pivot_partial.loc[prev_year].replace(0, np.nan)) * 100.0
            
            pivot_diff.loc[last_year] = ytd_diff
            pivot_pct.loc[last_year] = ytd_pct

    pivot_pct = pivot_pct.replace([np.inf, -np.inf], np.nan)

    # 행 라벨 변경 (예: 2026 -> 2026.1-7)
    if is_partial and last_year in pivot_full.index:
        pivot_full = pivot_full.rename(index={last_year: partial_label})
        pivot_diff = pivot_diff.rename(index={last_year: partial_label})
        pivot_pct = pivot_pct.rename(index={last_year: partial_label})

    pivot_base = pivot_full
    date_index_col = '연도'

else:
    # 월별 실적 (MoM)
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

# 5. 2중 헤더 테이블 구성
final_data = {}
for col in pivot_base.columns:
    final_data[(col, col if col != '합계' else '합계')] = pivot_base[col]
    final_data[(col, '증감량')] = pivot_diff[col]
    final_data[(col, '증감률')] = pivot_pct[col]

pivot_final = pd.DataFrame(final_data, index=pivot_base.index)
pivot_final = pivot_final.sort_index(ascending=True)

# 6. 메인 표 출력
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

format_dict = {}
for col in pivot_final.columns:
    format_dict[col] = (lambda v, t=col[1]: format_values(v, t))

styled_table = (
    pivot_final.style
    .format(format_dict)
    .map(apply_styles)
)

st.dataframe(styled_table, use_container_width=True, height=450)

# 7. 차트 섹션
st.write("---")
st.markdown(f"### 📊 {selected_cat} {country_title_label} {trade_type} 추이 차트")

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.markdown("##### 📈 세부 품목별 실적 (톤)")
    item_cols = [c for c in pivot_base.columns if c != '합계']
    chart_line_df = pivot_base[item_cols].reset_index()
    # 인덱스 컬럼명 명시
    chart_line_df.rename(columns={chart_line_df.columns[0]: date_index_col}, inplace=True)
    chart_line_df = chart_line_df.melt(
        id_vars=date_index_col, 
        var_name='품목', 
        value_name='실적(톤)'
    )
    
    fig_line = px.line(
        chart_line_df, 
        x=date_index_col, 
        y='실적(톤)', 
        color='품목',
        markers=True
    )
    fig_line.update_layout(
        margin=dict(l=20, r=20, t=30, b=20),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    fig_line.update_yaxes(tickformat=",.0f")
    st.plotly_chart(fig_line, use_container_width=True)

with chart_col2:
    st.markdown("##### 🏛️ 합계 실적 및 증감량 (톤)")
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=pivot_base.index.astype(str), y=pivot_base['합계'], name='총 실적(톤)', marker_color='#4A90E2'
    ))
    fig_bar.add_trace(go.Scatter(
        x=pivot_base.index.astype(str), y=pivot_diff['합계'], name='증감량(톤)', mode='lines+markers', line=dict(color='#E2594A', width=2), yaxis='y2'
    ))
    fig_bar.update_layout(
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="총 실적(톤)", tickformat=",.0f"),
        yaxis2=dict(title="증감량(톤)", overlaying='y', side='right', showgrid=False, tickformat=",.0f")
    )
    st.plotly_chart(fig_bar, use_container_width=True)