import streamlit as st
import pandas as pd
import numpy as np
import os
import glob
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="제지산업 수출입통계 대시보드", layout="wide")

# 1. 파일 경로 자동 탐색
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
    
    if '중분류' not in df.columns:
        df['중분류'] = df['품목명']
        
    return df

df = load_data(latest_file)

# 2. 상단 필터 UI (품목 -> 수출/수입 -> 집계주기 순서)
col_cat, col_trade, col_period = st.columns([1.5, 1, 1.2])

with col_cat:
    raw_cats = df['대분류'].dropna().unique().tolist()
    # 대분류에서 '폐신문' 또는 '폐신문지'가 있다면 제외
    raw_cats = [c for c in raw_cats if c not in ['폐신문', '폐신문지']]
    
    priority_order = ['폐지', '골판지원지', '펄프']
    available_cats = [c for c in priority_order if c in raw_cats] + [c for c in raw_cats if c not in priority_order]

    selected_cat = st.radio(
        "📂 **대분류 품목**",
        available_cats,
        format_func=lambda x: "골판지" if x == "골판지원지" else x,
        horizontal=True
    )

with col_trade:
    trade_type = st.radio("🔄 **수출 / 수입**", ["수출", "수입"], horizontal=True)

with col_period:
    period_type = st.radio("📅 **집계 주기**", ["연간 합계 (YoY)", "월별 실적 (MoM)"], horizontal=True)

# 사이드바 국가 필터
filtered_df = df[df['대분류'] == selected_cat].copy()
all_countries = sorted(filtered_df['국가명'].dropna().unique().tolist())
selected_countries = st.sidebar.multiselect("🌐 국가 필터 (선택 안 하면 전체 합산)", all_countries)

if selected_countries:
    filtered_df = filtered_df[filtered_df['국가명'].isin(selected_countries)]

if filtered_df.empty:
    st.warning("선택된 조건에 해당하는 데이터가 없습니다.")
    st.stop()

target_col = '수출실적(톤)' if trade_type == '수출' else '수입실적(톤)'
date_index_col = '기준연도' if "연간" in period_type else '기준년월'

# 3. 피벗 및 증감 계산
pivot_base = filtered_df.pivot_table(
    index=date_index_col,
    columns='중분류',
    values=target_col,
    aggfunc='sum'
).fillna(0)

# 폐지 선택 시 사용자 지정 순서 적용 (폐골판지 -> 폐신문지 -> 고급폐지 -> 기타폐지)
if selected_cat == '폐지':
    custom_waste_order = ['폐골판지', '폐신문지', '고급폐지', '기타폐지']
    ordered_cols = [c for c in custom_waste_order if c in pivot_base.columns]
    remaining_cols = [c for c in pivot_base.columns if c not in custom_waste_order]
    pivot_base = pivot_base[ordered_cols + remaining_cols]

pivot_base['합계'] = pivot_base.sum(axis=1)
pivot_diff = pivot_base.diff()
pivot_pct = pivot_base.pct_change() * 100.0
pivot_pct = pivot_pct.replace([np.inf, -np.inf], np.nan)

# 4. 2중 헤더 테이블 구성
final_data = {}
for col in pivot_base.columns:
    final_data[(col, col if col != '합계' else '합계')] = pivot_base[col]
    final_data[(col, '증감량')] = pivot_diff[col]
    final_data[(col, '증감률')] = pivot_pct[col]

pivot_final = pd.DataFrame(final_data, index=pivot_base.index)
pivot_final = pivot_final.sort_index(ascending=True)

# 5. 서식 적용 및 테이블 출력
display_cat_name = "골판지" if selected_cat == "골판지원지" else selected_cat
st.markdown(f"### 📋 {display_cat_name} {trade_type}실적 ({period_type})")
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

st.dataframe(styled_table, use_container_width=True, height=380)

# 6. 시각화 그래프 섹션
st.write("---")
st.markdown(f"### 📊 {display_cat_name} {trade_type} 실적 추이")

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.markdown("##### 📈 품목별 실적 (톤)")
    item_cols = [c for c in pivot_base.columns if c != '합계']
    chart_line_df = pivot_base[item_cols].reset_index().melt(
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
    st.markdown("##### 🏛️ 전체 합계 실적 및 증감량 (톤)")
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=pivot_base.index, y=pivot_base['합계'], name='총 실적(톤)', marker_color='#4A90E2'
    ))
    fig_bar.add_trace(go.Scatter(
        x=pivot_diff.index, y=pivot_diff['합계'], name='증감량(톤)', mode='lines+markers', line=dict(color='#E2594A', width=2), yaxis='y2'
    ))
    fig_bar.update_layout(
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(title="총 실적(톤)", tickformat=",.0f"),
        yaxis2=dict(title="증감량(톤)", overlaying='y', side='right', showgrid=False, tickformat=",.0f")
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# 7. 폐지 선택 시 국가별 Top 10 그래프 추가
if selected_cat == '폐지':
    st.write("---")
    st.markdown(f"### 🌍 폐지 {trade_type} 실적 상위 10개국 (전체 기간 합산 기준)")
    
    # 국가 필터 적용 전 원본 데이터에서 폐지 데이터만 추출하여 전체 국가 중 Top 10 계산
    top10_base_df = df[df['대분류'] == '폐지']
    country_sum_df = top10_base_df.groupby('국가명')[target_col].sum().reset_index()
    country_sum_df = country_sum_df.sort_values(by=target_col, ascending=False).head(10)
    
    fig_top10 = px.bar(
        country_sum_df,
        x='국가명',
        y=target_col,
        text=target_col,
        color='국가명',
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    # 막대 위에 수치 표시하고 y축 눈금은 숨겨서 깔끔하게 배치
    fig_top10.update_traces(texttemplate='%{text:,.0f}', textposition='outside', showlegend=False)
    fig_top10.update_layout(
        margin=dict(l=20, r=20, t=30, b=20),
        yaxis=dict(showticklabels=False, title=""),
        xaxis=dict(title="")
    )
    st.plotly_chart(fig_top10, use_container_width=True)