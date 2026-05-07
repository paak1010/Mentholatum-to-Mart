import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime

# ==========================================
# ⚙️ 페이지 설정 및 전역 변수
# ==========================================
st.set_page_config(page_title="멘소래담 통합 수주 자동화", page_icon="🏢", layout="wide")
today_str = datetime.today().strftime("%Y%m%d")

# 최종 통일 양식 컬럼
FINAL_COLUMNS = [
    '구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 
    'ME코드', '상품명', '수량', '단가', 'Total Amount'
]

# ==========================================
# 헬퍼 함수 (엑셀 변환 및 포맷팅)
# ==========================================
def to_excel_unified(df, sheet_name="통합_수주업로드"):
    numeric_cols = ['수량', '단가', 'Total Amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        num_format = workbook.add_format({'num_format': '#,##0'})
        center_format = workbook.add_format({'align': 'center'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#F0F2F6', 'border': 1, 'align': 'center'})
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        for col_idx, col_name in enumerate(df.columns):
            if col_name in numeric_cols:
                worksheet.set_column(col_idx, col_idx, 12, num_format)
            elif col_name in ['구분', '수주날짜', '납품일자', '발주코드', '배송코드']:
                worksheet.set_column(col_idx, col_idx, 14, center_format)
            elif col_name in ['상품명', '배송처']:
                worksheet.set_column(col_idx, col_idx, 30)
            else:
                worksheet.set_column(col_idx, col_idx, 15)
    return output.getvalue()

# ==========================================
# 🔴 [로직 1] TESCO 처리 함수
# ==========================================
def run_tesco_logic(uploaded_file):
    FULL_PRODUCT_MAP = {
        8809020342310: 'ME90521CLA', 8809020342211: 'ME90521CLL', 8809020342419: 'ME90521CLS',
        # ... (중략 - 기존 코드의 매핑 데이터 그대로 사용)
        8809020344338: 'ME00621FHH', 8809020344321: 'ME90621MAM'
    }
    RAW_STORE_MAP = {
        '0903목천물류서비스센터SORTATION': 81020901, '0903목천물류서비스센터FLOW': 81020902,
        '0903목천물류서비스센터STOCK': 81020903, '0982안성ADC물류센터STOCK': 81020982,
        # ... (기존 로직 동일)
    }
    NORM_STORE_MAP = {re.sub(r'^\d+', '', k).replace(" ", "").upper(): v for k, v in RAW_STORE_MAP.items()}

    if uploaded_file.name.endswith('.csv'):
        content = uploaded_file.getvalue()
        try: text = content.decode('utf-8-sig')
        except: text = content.decode('cp949')
        all_rows = list(csv.reader(io.StringIO(text)))
    else:
        df_temp = pd.read_excel(uploaded_file, header=None)
        all_rows = df_temp.fillna('').astype(str).values.tolist()

    parsed_data = []
    col_map = {}
    for row in all_rows:
        row_strs = [str(x).strip() for x in row]
        if '상품코드' in row_strs and ('발주금액' in row_strs or '낱개수량' in row_strs):
            col_map = {k: row_strs.index(v) for k, v in {
                '상품명':'상품명', '상품코드':'상품코드', '입고타입':'입고타입', 
                '수량':'낱개수량', '단가':'낱개당 단가', '금액':'발주금액', 
                '납품처':'납품처', '납품일자':'납품일자'}.items() if v in row_strs}
            continue
        if not col_map or '상품코드' not in col_map: continue
        
        try:
            b_idx = col_map['상품코드']
            barcode = int(re.sub(r'[^\d]', '', row_strs[b_idx]))
            if barcode in FULL_PRODUCT_MAP:
                parsed_data.append({
                    '상품명': row_strs[col_map['상품명']], '바코드': barcode, 
                    '입고타입': row_strs[col_map['입고타입']], '수량': float(re.sub(r'[^\d.]', '', row_strs[col_map['수량']])),
                    '단가': float(re.sub(r'[^\d.]', '', row_strs[col_map['단가']])), '금액': float(re.sub(r'[^\d.]', '', row_strs[col_map['금액']])),
                    '납품처': row_strs[col_map['납품처']], '납품일자': row_strs[col_map['납품일자']]
                })
        except: pass

    df = pd.DataFrame(parsed_data)
    if df.empty: return pd.DataFrame()
    
    df['ME코드'] = df['바코드'].map(FULL_PRODUCT_MAP)
    def get_store_code(row):
        s = re.sub(r'^\d+', '', str(row['납품처']).replace(' ', '').upper())
        t = 'FLOW' if 'FLOW' in str(row['입고타입']).upper() else ('SORTATION' if 'MIX' in str(row['입고타입']).upper() else '')
        key = s + t
        return next((v for k, v in NORM_STORE_MAP.items() if k in key or key in k), 81040913)

    df['배송코드'] = df.apply(get_store_code, axis=1)
    df['발주코드'] = 81020000
    df['수주날짜'] = today_str
    df['납품일자'] = pd.to_datetime(df['납품일자'], errors='coerce').dt.strftime('%Y%m%d')
    df['발주처'] = 'Tesco'
    df['구분'] = "0"
    df.rename(columns={'납품처': '배송처', '금액': 'Total Amount'}, inplace=True)
    return df[FINAL_COLUMNS]

# ==========================================
# 🟡 [로직 2] 이마트 처리 함수
# ==========================================
def run_emart_logic(uploaded_file, prod_df):
    if uploaded_file.name.endswith('.csv'):
        try: raw_df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
        except: raw_df = pd.read_csv(uploaded_file, encoding='cp949')
    else:
        raw_df = pd.read_excel(uploaded_file)

    raw_df = raw_df.dropna(subset=['점포코드'])
    raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)
    raw_df['센터코드'] = raw_df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False).str.strip()
    raw_df['수량'] = pd.to_numeric(raw_df.get('수량', 0), errors='coerce').fillna(0)
    
    date_col = next((c for c in ['센터입하일자', '센터입하일', '점입점일자'] if c in raw_df.columns), '')
    raw_df['납품일자'] = raw_df[date_col].astype(str).str.replace(r'[^0-9]', '', regex=True)

    emart_map_dict = {
        'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
        'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
        'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
    }
    
    delivery_name_map = {'81010902': '이마트 시화물류센터', '81010905': '이마트 여주물류센터', '81033036': '이마트 트레이더스 평택물류', #... (중략)
    }

    def process_row(row):
        code, center = row['점포코드'], str(row['센터코드'])
        cust = 'E-mart' if (1000 <= code <= 1999 or code >= 9000) else ('E-mart(TRD)' if 2000 <= code <= 2999 else 'E-mart(노브랜드)')
        m_code = emart_map_dict.get(cust, {}).get(center, center)
        return pd.Series([cust, m_code])

    raw_df[['발주처', '배송코드']] = raw_df.apply(process_row, axis=1)
    raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).strip()
    
    merged = pd.merge(raw_df, prod_df, left_on='상품코드', right_on='바코드', how='left')
    merged['ME코드'] = merged['상품코드(기획)'].fillna(merged['상품코드'])
    merged['상품명'] = merged['상품명(기획)'].fillna(merged.get('상품명', ''))
    merged['배송처'] = merged['배송코드'].map(delivery_name_map).fillna(merged['배송코드'])
    merged['수주날짜'] = today_str
    merged['발주코드'] = '81010000'
    merged['구분'] = "0"
    merged.rename(columns={'발주원가': '단가', '발주금액': 'Total Amount'}, inplace=True)
    
    return merged[FINAL_COLUMNS]

# ==========================================
# 🟢 [로직 3] 롯데마트 처리 함수
# ==========================================
def run_lotte_logic(uploaded_file):
    CENTER_MAP = {'오산센터': '81030907', '김해센터': '81030908'}
    if uploaded_file.name.endswith('.csv'): df_edi = pd.read_csv(uploaded_file, header=None)
    else: df_edi = pd.read_excel(uploaded_file, header=None)

    parsed_list, curr_center, curr_doc_no, curr_date = [], "", "", ""
    for _, row in df_edi.dropna(how='all').iterrows():
        r = [str(x).strip() for x in row.tolist()]
        if r[0] == 'ORDERS':
            curr_doc_no = r[1].replace('.0', '')
            curr_center = re.sub(r'상온센타|상온센터|센타', '센터', r[5]).replace('센터센터', '센터')
            curr_date = re.sub(r'[^0-9]', '', r[7])
            continue
        if r[1].startswith('880'):
            qty = int(float(r[6])) * (int(float(r[5])) or 1)
            parsed_list.append({
                '발주코드': curr_doc_no, '배송처': curr_center, '납품일자': curr_date,
                'ME코드': r[1].replace('.0', ''), '상품명': r[2], '수량': qty, '단가': float(r[7]), 'Total Amount': qty * float(r[7])
            })
    
    df = pd.DataFrame(parsed_list)
    if df.empty: return df
    df['배송코드'] = df['배송처'].map(lambda x: next((v for k, v in CENTER_MAP.items() if k in x), '81030000'))
    df['수주날짜'] = today_str
    df['발주처'] = '롯데마트'
    df['구분'] = "0"
    return df[FINAL_COLUMNS]

# ==========================================
# 🚀 메인 앱 실행 부분
# ==========================================
with st.sidebar:
    st.header("⚙️ 마스터 데이터")
    # 이마트 마스터 파일 로드 (서버 내 파일 필요)
    @st.cache_data
    def load_master():
        files = ["NEW 이마트 서식파일_20260420납품.xlsx", "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx", "NEW 노브랜드_20260409납품.xlsx"]
        appended = []
        for f in files:
            if os.path.exists(f):
                d = pd.read_excel(f, sheet_name=0)
                d.columns = d.columns.str.strip()
                appended.append(d)
        return pd.concat(appended).drop_duplicates(subset=['바코드']) if appended else None
    
    emart_master = load_master()
    if emart_master is not None: st.success("✅ 이마트 마스터 로드 완료")
    else: st.warning("⚠️ 이마트 마스터 파일 없음")

st.title("📦 마트 통합 수주 자동 변환 대시보드")
st.markdown("> 파일을 한꺼번에 업로드하세요. Tesco, 이마트, 롯데마트를 자동 인식하여 합쳐줍니다.")

uploaded_files = st.file_uploader("📂 발주서 파일 업로드 (xlsx, csv)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)

if uploaded_files:
    results = []
    with st.spinner("🔄 데이터 통합 변환 중..."):
        for f in uploaded_files:
            # 1. 샘플 읽어서 마트 판별
            f.seek(0)
            sample = str(pd.read_excel(f, nrows=10).values) if not f.name.endswith('.csv') else str(f.getvalue()[:1000])
            f.seek(0)

            if 'ORDERS' in sample: 
                df = run_lotte_logic(f)
                mart = "롯데마트"
            elif '점포코드' in sample or '센터입하' in sample:
                df = run_emart_logic(f, emart_master) if emart_master is not None else pd.DataFrame()
                mart = "이마트"
            else:
                df = run_tesco_logic(f)
                mart = "Tesco"
            
            if not df.empty:
                results.append(df)
                st.write(f"✔️ {f.name} -> {mart} 인식 완료 ({len(df)}건)")

    if results:
        final_df = pd.concat(results, ignore_index=True)
        st.divider()
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 총 건수", f"{len(final_df):,} 건")
        c2.metric("🔢 총 수량", f"{final_df['수량'].sum():,.0f} 개")
        c3.metric("💰 총 금액", f"{final_df['Total Amount'].sum():,.0f} 원")

        st.dataframe(final_df, use_container_width=True, height=600)
        
        st.download_button(
            "📥 통합 결과 엑셀 다운로드",
            data=to_excel_unified(final_df),
            file_name=f"통합수주_업로드용_{today_str}.xlsx",
            mime="application/vnd.ms-excel",
            type="primary", use_container_width=True
        )
