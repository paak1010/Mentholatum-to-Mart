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
    FULL_PRODUCT_MAP = {
        8809020342310: 'ME90521CLA', 8809020342211: 'ME90521CLL', 8809020342419: 'ME90521CLS',
        8809020340804: 'ME90521MC1', 8809020340774: 'ME90521LP2', 8809020348992: 'ME90521E18',
        8809020340279: 'ME90521LR1', 8809020344444: 'ME90521EL9', 8809020344451: 'ME90521EL8',
        8809020344468: 'ME90521EL7', 8809020344192: 'ME90521EL6', 8809020344048: 'ME90521EL4',
        8809020344123: 'ME90521EL0', 8809020344239: 'ME90521E13', 8809020349821: 'ME90521CC4',
        8809020349814: 'ME90521CC2', 8809020349807: 'ME90521CC1', 8809020345212: 'ME00421186',
        8809020345236: 'ME00421183', 8809020345229: 'ME00421301', 8809020348978: 'ME00421151',
        8809020349661: 'ME90621CPS', 8809020349654: 'ME90621CPM', 8809020346516: 'ME90621AT2',
        8809020340286: 'ME00621AB5', 8809020340293: 'ME00621C21', 8809020346561: 'ME00621AT6',
        8809020346585: 'ME90621NA7', 8809020346592: 'ME90621ADI', 8809020346660: 'ME90621A07',
        8809020349425: 'ME00621A08', 8809020349685: 'ME00621AS1', 8809020349692: 'ME00621AL1',
        8809020349708: 'ME00621AR1', 8809020349715: 'ME00621AG1', 8809020349722: 'ME00621AF9',
        8809020349371: 'ME90621GK3', 8809020349418: 'ME90621GK2', 8809020349388: 'ME90621GL3',
        8809020349050: 'ME90621GLO', 8809020349067: 'ME90621GM4', 8809020349074: 'ME90621GE1',
        8809020349203: 'ME90621HCR', 8809020349098: 'ME90621HSL', 8809020349104: 'ME90621SM4',
        8809020349210: 'ME90621SCM', 8809020349166: 'ME90621GO8', 8809020349906: 'ME90621GLL',
        8809020349944: 'ME90621FGC', 8809020340200: 'ME00621H37', 8809020340217: 'ME00621H38',
        8809020340170: 'ME00621C15', 8809020340187: 'ME00621S24', 8809020340194: 'ME00621AS3',
        8809020340606: 'ME00621C22', 8809020340590: 'ME00621H44', 8809020340712: 'ME90621TC1',
        8809020341627: 'ME00621FMC', 8809020341634: 'ME00621FMR', 8809020341641: 'ME00621FBR',
        8809020341207: 'ME80421DR2', 8809020341061: 'ME81921SLL', 8809020341054: 'ME81921SVV',
        8809020341801: 'ME81921SL1', 8809020342501: 'ME90521LD9', 8809020342518: 'ME90521GT2',
        8809020342495: 'ME90521GS2', 8809020349036: 'ME00621CM5', 8809020346509: 'ME90621AFE',
        8809020349968: 'ME00621H41', 8809020342433: 'ME90621AC4', 8809020343478: 'ME00621ABN',
        8809020342525: 'ME80421DCH', 8809020343683: 'ME90521WC4', 8809020343690: 'ME90521WC5',
        8809020343706: 'ME90521WC6', 8809020344338: 'ME00621FHH', 8809020344321: 'ME90621MAM'
    }
    
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
    df['배송코드'] = '81020000' # 테스코 배송코드 적용 가능
    df['발주처'] = 'Tesco'
    df['수주날짜'] = today_str
    df['구분'] = '0'
    df['Total Amount'] = df['수량'] * df['단가']
    return df

# ==========================================
# 🟡 [로직 2] 이마트 처리
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
        merged['상품명'] = merged.get('상품명', '')

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
# 🟢 [로직 3] 롯데마트 처리
# ==========================================
def run_lotte_logic(uploaded_file, lotte_prod_df):
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
    
    if lotte_prod_df is not None:
        df = pd.merge(df, lotte_prod_df, on='바코드', how='left')
        df['ME코드'] = df['ME코드'].fillna(df['바코드'])
        df['상품명'] = df['마스터_품명'].fillna(df['상품명_원본'])
    else:
        df['ME코드'] = df['바코드']
        df['상품명'] = df['상품명_원본']

    df[['배송코드', '배송처']] = df['원본_배송처'].apply(lambda x: pd.Series(get_lotte_delivery_info(x)))
    df['발주처'] = '롯데마트'
    df['발주코드'] = '81030000'
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
        
        # 1. 이마트 마스터 파일 로드
        e_files = ["NEW 이마트 서식파일_20260420납품.xlsx", "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx", "NEW 노브랜드_20260409납품.xlsx"]
        e_list = []
        for f in e_files:
            if os.path.exists(f):
                xls = pd.ExcelFile(f)
                target = xls.sheet_names[0]
                for s in xls.sheet_names:
                    if any(x in s for x in ['제품', '상품', '단가']):
                        target = s
                        break
                d = pd.read_excel(xls, sheet_name=target)
                d.columns = d.columns.astype(str).str.strip()
                e_list.append(d)
                
        if e_list:
            temp_emart = pd.concat(e_list, ignore_index=True)
            # ⭐ 핵심 에러 방지 구역 (바코드 열이 없을 경우 대처)
            if '바코드' in temp_emart.columns:
                temp_emart['바코드'] = temp_emart['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
                emart_m = temp_emart.drop_duplicates(subset=['바코드'])
            elif '상품코드' in temp_emart.columns:
                temp_emart['바코드'] = temp_emart['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
                emart_m = temp_emart.drop_duplicates(subset=['바코드'])
            else:
                st.warning("⚠️ 이마트 마스터 파일에 '바코드' 또는 '상품코드' 열을 찾을 수 없습니다.")
                emart_m = temp_emart # 에러 내지 않고 원본 유지

        # 2. 롯데마트 마스터 파일 로드
        l_file = "2022 롯데마트 서식파일 260417납품.xlsx"
        if os.path.exists(l_file):
            try:
                l_map = pd.read_excel(l_file, sheet_name=0)
                l_price = pd.read_excel(l_file, sheet_name=1)
                
                # 열 개수가 정상인지 확인 후 매핑 진행
                if len(l_map.columns) > 13 and len(l_price.columns) > 1:
                    l_map_sub = l_map.iloc[:, [3, 13]].copy()
                    l_map_sub.columns = ['바코드', 'ME코드']
                    
                    l_price_sub = l_price.iloc[:, [0, 1]].copy()
                    l_price_sub.columns = ['ME코드', '마스터_품명']
                    
                    l_m = pd.merge(l_map_sub, l_price_sub, on='ME코드', how='left')
                    l_m['바코드'] = l_m['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
                    lotte_m = l_m.drop_duplicates(subset=['바코드'])
            except Exception as e:
                st.warning(f"⚠️ 롯데마트 맵핑 실패 (에러무시): {e}")
                
        return emart_m, lotte_m

    emart_master, lotte_master = load_masters()
    if emart_master is not None and '바코드' in emart_master.columns: st.success("✅ 이마트 마스터 로드")
    if lotte_master is not None and '바코드' in lotte_master.columns: st.success("✅ 롯데마트 마스터 로드")

st.title("📦 통합 마트 수주 자동 변환 대시보드")
uploaded_files = st.file_uploader("📂 발주서 파일 업로드", accept_multiple_files=True)

if uploaded_files:
    results = []
    for f in uploaded_files:
        f.seek(0)
        try:
            if f.name.endswith('.csv'): sample = f.read(2000).decode('utf-8-sig', errors='ignore')
            else: sample = str(pd.read_excel(f, nrows=10).values)
        except:
            sample = ""
        f.seek(0)
        
        try:
            if 'ORDERS' in sample:
                df = run_lotte_logic(f, lotte_master)
                mart_name = "롯데마트"
            elif '점포코드' in sample or '센터입하' in sample:
                df = run_emart_logic(f, emart_master)
                mart_name = "이마트"
            else:
                df = run_tesco_logic(f)
                mart_name = "Tesco"
                
            if not df.empty:
                results.append(df)
                st.write(f"✔️ **{f.name}** -> {mart_name} 변환 완료 ({len(df)}건)")
            else:
                st.error(f"❌ {f.name}: 추출된 데이터가 없습니다.")
        except Exception as e:
             st.error(f"❌ {f.name} 처리 중 오류 발생: {e}")

    if results:
        final_df = pd.concat(results, ignore_index=True).fillna("")
        
        # 데이터 병합 (동일 항목 수량/금액 합산)
        group_cols = ['구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '단가']
        final_df = final_df.groupby(group_cols, as_index=False).agg({'수량': 'sum', 'Total Amount': 'sum'})
        
        # 누락된 컬럼 방지 및 순서 정렬
        for col in FINAL_COLUMNS:
            if col not in final_df.columns:
                final_df[col] = ""
        final_df = final_df[FINAL_COLUMNS]

        st.divider()
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 총 병합 건수", f"{len(final_df):,} 건")
        c2.metric("🔢 총 납품 수량", f"{final_df['수량'].sum():,.0f} 개")
        c3.metric("💰 총 납품 금액", f"{final_df['Total Amount'].sum():,.0f} 원")

        st.dataframe(final_df, use_container_width=True)
        st.download_button("📥 통합 결과 다운로드", data=to_excel_unified(final_df), file_name=f"통합수주_{today_str}.xlsx")
