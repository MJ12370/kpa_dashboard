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
    border: 1px solid #1E3A8A;
    border-radius: 4px;
}
.custom-table {
    width: 100%;
    border-collapse: collapse;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    font-size: 13px;
    text-align: center;
}
.custom-table th {
    background-color: #1E3A8A !important;
    color: #FFFFFF !important;
    font-weight: 700;
    padding: 7px 5px;
    border: 1px solid #2563EB;
    text-align: center !important;
}
.custom-table td {
    padding: 6px 5px;
    border: 1px solid #E5E7EB;
    font-weight: 600;
    color: #111827;
    text-align: center !important;
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
}
.row-total td {
    background-color: #FEF3C7 !important;
    font-weight: 800 !important;
    color: #92400E !important;
}
.row-grand-total td {
    background-color: #DBEAFE !important;
    font-weight: 800 !important;
    color: #1E3A8A !important;
}
.val-negative {
    color: #DC2626 !important;
    font-weight: 700;
}
.val-zero {
    color: #9CA3AF !important;
}
</style>
""", unsafe_allow_html=True)

DATA_DIR = r"D:\kita" if os.path.exists(r"D:\kita") else "."

# 최상단 대시보드 모드 선택
st.sidebar.markdown("## 📊 대시보드 메뉴")
main_menu = st.sidebar.radio("통계 구분 선택", ["원료 수출입 통관실적", "종이판지 수출입 통관실적"])
st.sidebar.write("---")

# ==========================================
# 1. 종이판지 수출입 통관실적 대시보드
# ==========================================
if main_menu == "종이판지 수출입 통관실적":
    paper_files = glob.glob(os.path.join(DATA_DIR, "*종이판지*수출입통계*.xlsx"))
    if not paper_files:
        st.error("종이판지 수출입 통계 엑셀 파일을 찾을 수 없습니다. (D:\\kita\\종이판지_수출입통계_update.xlsx)")
        st.stop()
    latest_paper_file = max(paper_files, key=os.path.getctime)

    @st.cache_data
    def load_paper_data(path):
        df = pd.read_excel(path, sheet_name="종이판지")
        df = df[df['월'] != '누계'].copy()
        df['연도'] = df['연도'].astype(int)
        df['월'] = df['월'].astype(int)
        df['중량(톤)'] = pd.to_numeric(df['중량(톤)'], errors='coerce').fillna(0)
        return df

    df_p = load_paper_data(latest_paper_file)

    st.title("종이판지 수출입 통관실적")
    st.write("---")

    # 사이드바 컨트롤
    st.sidebar.markdown("### 🔄 무역 구분")
    paper_trade = st.sidebar.radio("구분", ["수출", "수입"])

    # 상단 기준년월 선택
    df_p['년월'] = df_p['연도'].astype(str) + "." + df_p['월'].apply(lambda x: f"{x:02d}")
    available_ym = sorted(df_p['년월'].unique(), reverse=True)

    col_ym, col_space = st.columns([1.5, 3.5])
    with col_ym:
        target_ym = st.selectbox("📅 **기준 년월 선택**", available_ym, index=0)

    cur_year, cur_month = map(int, target_ym.split("."))
    prev_year = cur_year - 1

    # 품목 계층 마스터 테이블 정의
    items_tree = [
        ('종이', '신문용지', '-', '신문용지', ''),
        ('종이', '비도공\n인쇄용지', '백상지', '백상지', ''),
        ('종이', '비도공\n인쇄용지', '기타', '비도공 기타', ''),
        ('종이', '비도공\n인쇄용지', '계', '__CALC_비도공계__', 'row-subtotal'),
        ('종이', '도공\n인쇄용지', '아트지류', '아트지', ''),
        ('종이', '도공\n인쇄용지', '기타', '도공 기타', ''),
        ('종이', '도공\n인쇄용지', '계', '__CALC_도공계__', 'row-subtotal'),
        ('종이', '박엽인쇄용지', '-', '박엽인쇄용지', ''),
        ('종이', '정보\n인쇄용지', '감열기록지', '감열기록지', ''),
        ('종이', '정보\n인쇄용지', '복사용지', '복사용지', ''),
        ('종이', '정보\n인쇄용지', '전산용지', '전산용지', ''),
        ('종이', '정보\n인쇄용지', '계', '__CALC_정보계__', 'row-subtotal'),
        ('종이', '인쇄용지소계', '-', '__CALC_인쇄용지소계__', 'row-subtotal'),
        ('종이', '기타\n특수지', '팬시지', '팬시지', ''),
        ('종이', '기타\n특수지', '권련지', '권련지', ''),
        ('종이', '기타\n특수지', '계', '__CALC_특수지계__', 'row-subtotal'),
        ('종이', '위생용지', '-', '위생용지', ''),
        ('종이', '중포대용크라프트지', '-', '중포대용크라프트지', ''),
        ('종이', '종이합계', '-', '__CALC_종이합계__', 'row-total'),
        
        ('판지', '백판지(도공)', '카톤용', '도공 카톤', ''),
        ('판지', '백판지(도공)', 'SC', '도공 SC', ''),
        ('판지', '백판지(도공)', '아이보리', '도공 아이보리', ''),
        ('판지', '백판지(도공)', '계', '__CALC_백판지도공계__', 'row-subtotal'),
        ('판지', '백판지(비도공)', '카톤용', '비도공 카톤', ''),
        ('판지', '백판지(비도공)', 'TM', '비도공 TM', ''),
        ('판지', '백판지(비도공)', '계', '__CALC_백판지비도공계__', 'row-subtotal'),
        ('판지', '백판지 계', '-', '__CALC_백판지계__', 'row-subtotal'),
        ('판지', '골판지원지', '라이너', '라이너', ''),
        ('판지', '골판지원지', '골심지', '골심지', ''),
        ('판지', '골판지원지', '계', '__CALC_골판지원지계__', 'row-subtotal'),
        ('판지', '기타판지', '밀크카톤등', '밀크카톤', ''),
        ('판지', '기타판지', '컵원지 등', '컵원지', ''),
        ('판지', '기타판지', '판지 기타', '판지 기타', ''),
        ('판지', '기타판지', '계', '__CALC_기타판지계__', 'row-subtotal'),
        ('판지', '판지합계', '-', '__CALC_판지합계__', 'row-total'),
        
        ('종이판지합계', '-', '-', '__CALC_종이판지합계__', 'row-grand-total'),
        
        ('종이제품', '골판상자', '-', '골판상자', ''),
        ('종이제품', '지대', '-', '지대', ''),
        ('종이제품', '감열기록지(제품)', '-', '감열기록지(제품)', ''),
        ('종이제품', '카본지/복사지', '-', '카본지', ''),
        ('종이제품', '지제품합계', '-', '__CALC_지제품합계__', 'row-total'),
        
        ('총합계', '-', '-', '__CALC_총합계__', 'row-grand-total')
    ]

    filtered = df_p[df_p['수출/수입'] == paper_trade]

    def get_val(item_code, year, month=None, is_ytd=False):
        if is_ytd:
            sub = filtered[(filtered['연도'] == year) & (filtered['월'] <= month) & (filtered['지종'] == item_code)]
        else:
            sub = filtered[(filtered['연도'] == year) & (filtered['월'] == month) & (filtered['지종'] == item_code)]
        return sub['중량(톤)'].sum()

    records = {}
    unique_items = df_p['지종'].unique()
    for item in unique_items:
        cur_m = get_val(item, cur_year, cur_month, False)
        prev_m = get_val(item, prev_year, cur_month, False)
        cur_ytd = get_val(item, cur_year, cur_month, True)
        prev_ytd = get_val(item, prev_year, cur_month, True)
        records[item] = {
            'cur_m': cur_m, 'prev_m': prev_m,
            'cur_ytd': cur_ytd, 'prev_ytd': prev_ytd
        }

    def sum_items(item_list):
        res = {'cur_m': 0, 'prev_m': 0, 'cur_ytd': 0, 'prev_ytd': 0}
        for it in item_list:
            if it in records:
                for k in res:
                    res[k] += records[it][k]
        return res

    records['__CALC_비도공계__'] = sum_items(['백상지', '비도공 기타'])
    records['__CALC_도공계__'] = sum_items(['아트지', '도공 기타'])
    records['__CALC_정보계__'] = sum_items(['감열기록지', '복사용지', '전산용지'])
    records['__CALC_특수지계__'] = sum_items(['팬시지', '권련지'])
    records['__CALC_인쇄용지소계__'] = sum_items(['백상지', '비도공 기타', '아트지', '도공 기타', '박엽인쇄용지', '감열기록지', '복사용지', '전산용지'])
    records['__CALC_종이합계__'] = sum_items(['신문용지', '백상지', '비도공 기타', '아트지', '도공 기타', '박엽인쇄용지', '감열기록지', '복사용지', '전산용지', '팬시지', '권련지', '위생용지', '중포대용크라프트지'])

    records['__CALC_백판지도공계__'] = sum_items(['도공 카톤', '도공 SC', '도공 아이보리'])
    records['__CALC_백판지비도공계__'] = sum_items(['비도공 카톤', '비도공 TM'])
    records['__CALC_백판지계__'] = sum_items(['도공 카톤', '도공 SC', '도공 아이보리', '비도공 카톤', '비도공 TM'])
    records['__CALC_골판지원지계__'] = sum_items(['라이너', '골심지'])
    records['__CALC_기타판지계__'] = sum_items(['밀크카톤', '컵원지', '판지 기타'])
    records['__CALC_판지합계__'] = sum_items(['도공 카톤', '도공 SC', '도공 아이보리', '비도공 카톤', '비도공 TM', '라이너', '골심지', '밀크카톤', '컵원지', '판지 기타'])

    records['__CALC_종이판지합계__'] = sum_items(['신문용지', '백상지', '비도공 기타', '아트지', '도공 기타', '박엽인쇄용지', '감열기록지', '복사용지', '전산용지', '팬시지', '권련지', '위생용지', '중포대용크라프트지', '도공 카톤', '도공 SC', '도공 아이보리', '비도공 카톤', '비도공 TM', '라이너', '골심지', '밀크카톤', '컵원지', '판지 기타'])
    records['__CALC_지제품합계__'] = sum_items(['골판상자', '지대', '감열기록지(제품)', '카본지'])
    records['__CALC_총합계__'] = {k: records['__CALC_종이판지합계__'][k] + records['__CALC_지제품합계__'][k] for k in records['__CALC_종이판지합계__']}

    st.markdown(f"### 📋 종이·판지 {paper_trade} 실적 ({target_ym})")
    col_info, col_btn_excel, col_btn_print = st.columns([3.2, 1.1, 0.9])
    with col_info:
        st.caption(f"(단위 : 톤) | 기준: {cur_year}년 {cur_month}월 (당월 및 1~{cur_month}월 누계)")

    excel_rows = []
    for g1, g2, g3, code, _ in items_tree:
        d = records.get(code, {'cur_m':0, 'prev_m':0, 'cur_ytd':0, 'prev_ytd':0})
        diff_m = d['cur_m'] - d['prev_m']
        rate_m = (diff_m / d['prev_m'] * 100) if d['prev_m'] > 0 else 0
        diff_y = d['cur_ytd'] - d['prev_ytd']
        rate_y = (diff_y / d['prev_ytd'] * 100) if d['prev_ytd'] > 0 else 0

        excel_rows.append({
            '대분류': g1, '중분류': g2.replace('\n', ' '), '소분류': g3,
            f'{prev_year}.{cur_month:02d}': round(d['prev_m']),
            f'{cur_year}.{cur_month:02d}': round(d['cur_m']),
            '당월 증감량': round(diff_m), '당월 증감률(%)': round(rate_m, 1),
            f'{prev_year} 누계': round(d['prev_ytd']),
            f'{cur_year} 누계': round(d['cur_ytd']),
            '누계 증감량': round(diff_y), '누계 증감률(%)': round(rate_y, 1)
        })
    paper_excel_df = pd.DataFrame(excel_rows)

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        paper_excel_df.to_excel(writer, sheet_name="종이판지실적", index=False)
    excel_data = excel_buffer.getvalue()

    with col_btn_excel:
        st.download_button(
            label="📥 엑셀 다운로드",
            data=excel_data,
            file_name=f"종이판지_{paper_trade}실적_{target_ym}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_btn_print:
        components.html("""
            <style>
                body { margin: 0; padding: 0; }
                .print-btn {
                    width: 100%; height: 38px; background-color: #FFFFFF; color: #1E3A8A;
                    border: 1px solid #D1D5DB; border-radius: 8px; font-weight: 600; font-size: 14px;
                    cursor: pointer; display: flex; align-items: center; justify-content: center;
                }
                .print-btn:hover { background-color: #EFF6FF; border-color: #1E3A8A; }
            </style>
            <button class="print-btn" onclick="window.parent.print()">🖨️ 표 인쇄</button>
        """, height=40)

    html = ['<div class="custom-table-container"><table class="custom-table">']
    html.append('<thead>')
    html.append('<tr>')
    html.append('<th colspan="3" rowspan="2" style="vertical-align: middle;">구 분</th>')
    html.append(f'<th colspan="4">당 월 ({cur_month}월)</th>')
    html.append(f'<th colspan="4">누 계 (1~{cur_month}월)</th>')
    html.append('</tr>')
    html.append('<tr>')
    html.append(f'<th>{prev_year}.{cur_month:02d}</th><th>{cur_year}.{cur_month:02d}</th><th>증감량</th><th>증감률</th>')
    html.append(f'<th>{prev_year} 누계</th><th>{cur_year} 누계</th><th>증감량</th><th>증감률</th>')
    html.append('</tr></thead><tbody>')

    for g1, g2, g3, code, row_cls in items_tree:
        d = records.get(code, {'cur_m':0, 'prev_m':0, 'cur_ytd':0, 'prev_ytd':0})
        diff_m = d['cur_m'] - d['prev_m']
        rate_m = (diff_m / d['prev_m'] * 100) if d['prev_m'] > 0 else 0
        diff_y = d['cur_ytd'] - d['prev_ytd']
        rate_y = (diff_y / d['prev_ytd'] * 100) if d['prev_ytd'] > 0 else 0

        html.append(f'<tr class="{row_cls}">')
        if g2 == '-':
            html.append(f'<td colspan="3" style="text-align: center; font-weight: 700;">{g1}</td>')
        elif g3 == '-':
            html.append(f'<td>{g1}</td><td colspan="2" style="text-align: center;">{g2}</td>')
        else:
            html.append(f'<td>{g1}</td><td>{g2.replace(chr(10), "<br>")}</td><td>{g3}</td>')

        html.append(f'<td>{int(round(d["prev_m"])):,}</td>')
        html.append(f'<td>{int(round(d["cur_m"])):,}</td>')
        cls_diff_m = "val-negative" if diff_m < 0 else ""
        html.append(f'<td class="{cls_diff_m}">{int(round(diff_m)):,}</td>')
        cls_rate_m = "val-negative" if rate_m < 0 else ""
        html.append(f'<td class="{cls_rate_m}">{rate_m:,.1f}%</td>')

        html.append(f'<td>{int(round(d["prev_ytd"])):,}</td>')
        html.append(f'<td>{int(round(d["cur_ytd"])):,}</td>')
        cls_diff_y = "val-negative" if diff_y < 0 else ""
        html.append(f'<td class="{cls_diff_y}">{int(round(diff_y)):,}</td>')
        cls_rate_y = "val-negative" if rate_y < 0 else ""
        html.append(f'<td class="{cls_rate_y}">{rate_y:,.1f}%</td>')
        html.append('</tr>')

    html.append('</tbody></table></div>')
    st.markdown("".join(html), unsafe_allow_html=True)
    st.caption("※ 자료출처 : 관세청 통관통계")

# ==========================================
# 2. 원료 수출입 통관실적 대시보드
# ==========================================
else:
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
        period_mode = st.radio("📅 **조회 단위**", ["연간 합계 (YoY)", "월별 실적 (MoM)"], horizontal=True)

    if "연간" in period_mode:
        with col_start:
            start_year = st.selectbox("시작 연도", all_years, index=0)
        with col_end:
            valid_end_years = [y for y in all_years if y >= start_year]
            end_year = st.selectbox("종료 연도", valid_end_years, index=len(valid_end_years)-1)
        selected_desc = f"{start_year}년 ~ {end_year}년 (연간)"
    else:
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

    if "연간" in period_mode:
        period_type = "연간 합계 (YoY)"
        filtered_df = filtered_df[(filtered_df['기준연도'] >= start_year) & (filtered_df['기준연도'] <= end_year)]
    else:
        period_type = "월별 실적 (MoM)"
        filtered_df = filtered_df[(filtered_df['기준년월'] >= start_date) & (filtered_df['기준년월'] <= end_date)]

    if filtered_df.empty:
        st.warning("선택된 기간 및 조건에 해당하는 데이터가 없습니다.")
        st.stop()

    valid_years = sorted([y for y in filtered_df['기준연도'].unique() if len(str(y)) == 4 and str(y).isdigit()])
    last_year = valid_years[-1]
    last_year_months = sorted([m for m in filtered_df[filtered_df['기준연도'] == last_year]['월'].unique() if 1 <= m <= 12])
    is_partial = (len(last_year_months) < 12 and len(last_year_months) > 0 and "연간" in period_type)
    partial_label = f"{last_year}.{last_year_months[0]}-{last_year_months[-1]}" if is_partial else last_year

    if trade_type == "수출+수입" or chosen_sub_item != "전체":
        if "연간" in period_type:
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
        if "연간" in period_type:
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
                    border: 1px solid #D1D5DB; border-radius: 8px; font-weight: 600; font-size: 14px;
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
    index_header_name = "기준연도" if "연간" in period_type else "기준년월"
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
        html.append(f'<td style="font-weight: 700; background-color: #F1F5F9;">{idx}</td>')
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
            html.append(f'<td class="{css_class}">{display_val}</td>')
        html.append('</tr>')

    html.append('</tbody></table></div>')
    st.markdown("".join(html), unsafe_allow_html=True)
    st.caption("※ 자료출처 : 통계청")

    # 하단 차트
    chart_html = [
        '<div class="chart-box">',
        '<hr style="margin: 25px 0; border: none; border-top: 1px solid #E5E7EB;">',
        f'<h3 style="margin-bottom: 20px;">📊 {display_item_title} 추이 차트{title_country_str}</h3>'
    ]

    if trade_type == "수출+수입":
        chart_html.append('</div>')
        st.markdown("".join(chart_html), unsafe_allow_html=True)
        chart_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01') if "월별" in period_type else pivot_base.index.astype(str)
        fig_dual = go.Figure()
        fig_dual.add_trace(go.Scatter(x=chart_x, y=pivot_base['수출실적(톤)'], name='수출', mode='lines+markers', line=dict(color='#2E7D32', width=2.5)))
        fig_dual.add_trace(go.Scatter(x=chart_x, y=pivot_base['수입실적(톤)'], name='수입', mode='lines+markers', line=dict(color='#C62828', width=2.5)))
        st.plotly_chart(fig_dual, use_container_width=True)
    elif chosen_sub_item != "전체":
        chart_html.append('</div>')
        st.markdown("".join(chart_html), unsafe_allow_html=True)
        chart_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01') if "월별" in period_type else pivot_base.index.astype(str)
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
            x_col = '날짜' if "월별" in period_type else date_index_col
            if "월별" in period_type:
                chart_line_df['날짜'] = pd.to_datetime(chart_line_df[date_index_col].astype(str).str.replace('.', '-') + '-01')
            chart_line_df = chart_line_df.melt(id_vars=x_col, value_vars=item_cols, var_name='품목', value_name='실적(톤)')
            fig_line = px.line(chart_line_df, x=x_col, y='실적(톤)', color='품목', markers=True)
            st.plotly_chart(fig_line, use_container_width=True)
        with c2:
            st.markdown('<div class="chart-box"><h5 style="margin-bottom: 10px;">🏛️ 합계 실적 및 증감량 (톤)</h5></div>', unsafe_allow_html=True)
            bar_x = pd.to_datetime(pivot_base.index.astype(str).str.replace('.', '-') + '-01') if "월별" in period_type else pivot_base.index.astype(str)
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(x=bar_x, y=pivot_base['합계'], name='총 실적', marker_color='#4A90E2'))
            fig_bar.add_trace(go.Scatter(x=bar_x, y=pivot_diff['합계'], name='증감량', mode='lines+markers', line=dict(color='#E2594A', width=2), yaxis='y2'))
            fig_bar.update_layout(yaxis2=dict(title="증감량", overlaying='y', side='right', showgrid=False))
            st.plotly_chart(fig_bar, use_container_width=True)