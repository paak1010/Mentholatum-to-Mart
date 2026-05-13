import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime

# ==========================================
# ⚙️ 페이지 및 기본 설정
# ==========================================
st.set_page_config(page_title="멘소래담 마트 통합 수주 자동화", page_icon="🏢", layout="wide")
today_str = datetime.today().strftime("%Y%m%d")

FINAL_COLUMNS = [
    '구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 
    'ME코드', '상품명', '수량', '단가', 'Total Amount'
]

# ==========================================
# 🛠️ 공통 유틸리티 함수
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
        header_format = workbook.add_format({'bold': True, 'bg_color': '#F0F2F6', 'border': 1, 'align': 'center'})
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for col_idx, col_name in enumerate(df.columns):
            if col_name in numeric_cols:
                worksheet.set_column(col_idx, col_idx, 12, num_format)
            elif col_name in ['상품명', '배송처']:
                worksheet.set_column(col_idx, col_idx, 30)
            else:
                worksheet.set_column(col_idx, col_idx, 15)
    return output.getvalue()

# ==========================================
# 🔴 [로직 1] TESCO 처리
# ==========================================
def run_tesco_logic(uploaded_file):
    # (Tesco 제품 매핑 데이터는 사용자님의 기존 데이터를 그대로 사용하시면 됩니다)
    FULL_PRODUCT_MAP = {8809020342310: 'ME90521CLA', 8809020342211: 'ME90521CLL', 8809020342419: 'ME90521CLS'} # 예시
    
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
            col_map = {k: row_strs.index(v) for k, v in {'상품명':'상품명','상품코드':'상품코드','입고타입':'입고타입','수량':'낱개수량','단가':'낱개당 단가','금액':'발주금액','납품처':'납품처','납품일자':'납품일자'}.items() if v in row_strs}
            continue
        if not col_map: continue
        try:
            barcode = int(re.sub(r'[^\d]', '', row_strs[col_map['상품코드']]))
            if barcode in FULL_PRODUCT_MAP:
                parsed_data.append({
                    '상품명': row_strs[col_map['상품명']], 'ME코드': FULL_PRODUCT_MAP[barcode],
                    '수량': float(re.sub(r'[^\d.]', '', row_strs[col_map['수량']])),
                    '단가': float(re.sub(r'[^\d.]', '', row_strs[col_map['단가']])),
                    '배송처': row_strs[col_map['납품처']], '납품일자': row_strs[col_map['납품일자']]
                })
        except: pass
    
    df = pd.DataFrame(parsed_data)
    if df.empty: return df
    df['발주코드'] = '81020000'
    df['배송코드'] = '81020000' # 테스코 배송코드 로직 적용 필요시 추가
    df['발주처'] = 'Tesco'
    df['수주날짜'] = today_str
    df['구분'] = '0'
    df['Total Amount'] = df['수량'] * df['단가']
    return df

# ==========================================
# 🟡 [로직 2] 이마트 처리 (개선됨)
# ==========================================
def run_emart_logic(uploaded_file, prod_df):
    if uploaded_file.name.endswith('.csv'):
        try: raw_df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
        except: raw_df = pd.read_csv(uploaded_file, encoding='cp949')
    else:
        raw_df = pd.read_excel(uploaded_file)

    raw_df['점포코드'] = pd.to_numeric(raw_df.get('점포코드', 0), errors='coerce').fillna(0).astype(int)
    raw_df['센터코드'] = raw_df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False).str.strip()
    
    emart_map_dict = {
        'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
        'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
        'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
    }
    delivery_name_map = {
        '81010902': '이마트 시화물류센터', '81010905': '이마트 여주물류센터', '81010903': '이마트 대구물류센터',
        '81033036': '이마트 트레이더스 평택물류', '89011174': '이마트 트레이더스 대구물류', '81011012': '이마트 트레이더스 여주물류',
        '81010904': '이마트 노브랜드 여주2물류센터', '81010968': '이마트 노브랜드 여주물류센터', '81010969': '이마트 노브랜드 시화물류센터',
        '81010901': '이마트 백암물류센터', '81010906': '이마트 광주물류센터'
    }

    def process_row(row):
        code, center = row['점포코드'], str(row['센터코드'])
        cust = 'E-mart' if (1000 <= code <= 1999 or code >= 9000) else ('E-mart(TRD)' if 2000 <= code <= 2999 else 'E-mart(노브랜드)')
        m_code = emart_map_dict.get(cust, {}).get(center, center)
        return pd.Series([cust, m_code])

    raw_df[['발주처', '배송코드']] = raw_df.apply(process_row, axis=1)
    raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    if prod_df is not None:
        merged = pd.merge(raw_df, prod_df, left_on='상품코드', right_on='바코드', how='left')
        merged['ME코드'] = merged['상품코드(기획)'].fillna(merged['상품코드'])
        merged['상품명'] = merged['상품명(기획)'].fillna(merged.get('상품명', ''))
    else:
        merged = raw_df.copy()
        merged['ME코드'] = merged['상품코드']

    merged['배송처'] = merged['배송코드'].map(delivery_name_map).fillna(merged['배송코드'])
    merged['발주코드'] = '81010000'
    merged['수주날짜'] = today_str
    merged['구분'] = '0'
    merged['단가'] = pd.to_numeric(merged.get('발주원가', 0), errors='coerce').fillna(0)
    merged['수량'] = pd.to_numeric(merged.get('수량', 0), errors='coerce').fillna(0)
    merged['Total Amount'] = merged['수량'] * merged['단가']
    
    date_col = next((c for c in ['센터입하일자', '센터입하일', '점입점일자'] if c in merged.columns), None)
    merged['납품일자'] = merged[date_col].astype(str).str.replace(r'[^0-9]', '', regex=True) if date_col else today_str
    
    return merged

# ==========================================
# 🟢 [로직 3] 롯데마트 처리 (대폭 개선)
# ==========================================
def run_lotte_logic(uploaded_file, lotte_prod_df):
    # 롯데마트 배송코드 매핑 (키워드 매칭 방식)
    def get_lotte_delivery_info(center_name):
        name = str(center_name)
        if '오산' in name: return '81030907', '롯데 오산상온센터'
        if '김해' in name: return '81030908', '롯데 김해상온센터'
        return '81030000', name

    def extract_num(val):
        s = str(val).split('(')[0]
        s = re.sub(r'[^\d.]', '', s)
        try: return float(s) if s else 0.0
        except: return 0.0

    if uploaded_file.name.endswith('.csv'): 
        try: df_edi = pd.read_csv(uploaded_file, header=None, encoding='utf-8-sig')
        except: 
            uploaded_file.seek(0)
            df_edi = pd.read_csv(uploaded_file, header=None, encoding='cp949')
    else: 
        df_edi = pd.read_excel(uploaded_file, header=None)

    parsed_list, curr_center, curr_date = [], "", ""
    for _, row in df_edi.dropna(how='all').iterrows():
        r = [str(x).strip() for x in row.tolist()]
        if r[0] == 'ORDERS':
            curr_center = re.sub(r'상온센타|상온센터|센타', '센터', r[5])
            curr_date = re.sub(r'[^0-9]', '', r[7])
            continue
        if len(r) > 1 and r[1].startswith('880'):
            barcode = r[1].replace('.0', '')
            qty = int(extract_num(r[6])) * (int(extract_num(r[5])) or 1)
            price = extract_num(r[7])
            parsed_list.append({
                '바코드': barcode, '상품명_원본': r[2], '수량': qty, '단가': price, '납품일자': curr_date, '원본_배송처': curr_center
            })
    
    df = pd.DataFrame(parsed_list)
    if df.empty: return df
    
    # 롯데마트 상품 매핑
    if lotte_prod_df is not None:
        df = pd.merge(df, lotte_prod_df, on='바코드', how='left')
        df['ME코드'] = df['ME코드'].fillna(df['바코드'])
        df['상품명'] = df['마스터_품명'].fillna(df['상품명_원본'])
    else:
        df['ME코드'] = df['바코드']
        df['상품명'] = df['상품명_원본']

    # 배송처 정보 일괄 적용
    df[['배송코드', '배송처']] = df['원본_배송처'].apply(lambda x: pd.Series(get_lotte_delivery_info(x)))
    df['발주처'] = '롯데마트'
    df['발주코드'] = '81030000' # 롯데마트 고정 발주코드
    df['수주날짜'] = today_str
    df['구분'] = '0'
    df['Total Amount'] = df['수량'] * df['단가']
    
    return df

# ==========================================
# 🚀 메인 앱 로직
# ==========================================
with st.sidebar:
    st.header("⚙️ 마스터 파일 로드")
    
    @st.cache_data
    def load_masters():
        emart_m, lotte_m = None, None
        # 이마트 마스터 로드
        e_files = ["NEW 이마트 서식파일_20260420납품.xlsx", "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx", "NEW 노브랜드_20260409납품.xlsx"]
        e_list = []
        for f in e_files:
            if os.path.exists(f):
                d = pd.read_excel(f, sheet_name=0)
                d.columns = d.columns.astype(str).str.strip()
                e_list.append(d)
        if e_list:
            emart_m = pd.concat(e_list).drop_duplicates(subset=['바코드'])
            emart_m['바코드'] = emart_m['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()

        # 롯데마트 마스터 로드
        l_file = "2022 롯데마트 서식파일 260417납품.xlsx"
        if os.path.exists(l_file):
            l_map = pd.read_excel(l_file, sheet_name=0)
            l_price = pd.read_excel(l_file, sheet_name=1)
            # 롯데마트 매핑 구조에 맞춰 컬럼 추출
            l_m = pd.merge(
                l_map.iloc[:, [3, 13]].rename(columns={l_map.columns[3]:'바코드', l_map.columns[13]:'ME코드'}),
                l_price.iloc[:, [0, 1]].rename(columns={l_price.columns[0]:'ME코드', l_price.columns[1]:'마스터_품명'}),
                on='ME코드', how='left'
            )
            l_m['바코드'] = l_m['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            lotte_m = l_m.drop_duplicates(subset=['바코드'])
            
        return emart_m, lotte_m

    emart_master, lotte_master = load_masters()
    if emart_master is not None: st.success("✅ 이마트 마스터 로드")
    if lotte_master is not None: st.success("✅ 롯데마트 마스터 로드")

st.title("📦 통합 마트 수주 자동 변환 대시보드")
uploaded_files = st.file_uploader("📂 발주서 파일 업로드", accept_multiple_files=True)

if uploaded_files:
    results = []
    for f in uploaded_files:
        f.seek(0)
        sample = str(f.read(2000))
        f.seek(0)
        
        if 'ORDERS' in sample:
            df = run_lotte_logic(f, lotte_master)
        elif '점포코드' in sample or '센터입하' in sample:
            df = run_emart_logic(f, emart_master)
        else:
            df = run_tesco_logic(f)
            
        if not df.empty:
            results.append(df)
            st.write(f"✔️ {f.name} 처리 완료")

    if results:
        final_df = pd.concat(results, ignore_index=True).fillna("")
        # 최종 병합 (동일 항목 합산)
        group_cols = ['구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '단가']
        final_df = final_df.groupby(group_cols, as_index=False).agg({'수량': 'sum', 'Total Amount': 'sum'})
        final_df = final_df[FINAL_COLUMNS]

        st.dataframe(final_df, use_container_width=True)
        st.download_button("📥 통합 결과 다운로드", data=to_excel_unified(final_df), file_name=f"통합수주_{today_str}.xlsx")
