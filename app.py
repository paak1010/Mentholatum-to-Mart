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
    RAW_STORE_MAP = {
        '0903목천물류서비스센터SORTATION': 81020901, '0903목천물류서비스센터FLOW': 81020902,
        '0903목천물류서비스센터STOCK': 81020903, '0982안성ADC물류센터STOCK': 81020982,
        '0907밀양EXP센터FLOW': 81021903, '0967일죽물류서비스센터FLOW': 81021904,
        '0905기흥물류서비스센터FLOW': 81021907, '0961밀양물류센터FLOW': 81040912,
        '0961밀양물류센터STOCK': 81040913, '0906NEW함안상온물류센터FLOW': 81040912,
        '0906NEW함안상온물류센터SORTATION': 81040913, '0906NEW함안상온물류센터SORTER': 81040913,
        '0982안성ADC물류센터SORTATION': 81020980, '0982안성ADC물류센터FLOW': 81020981,
        '0970함안EXP물류센터SORTATION': 89029018, '0970함안EXP물류센터FLOW': 81040913,
        '0982안성ADC물류센터SINGLE': 81020981, '0906NEW함안상온물류센터SINGLE': 81040912,
        '0968365용인DSCDSD': 81040904, '0969남양주EXP물류센터FLOW': 81040905,
        '0968365용인DSCSTOCK': 81040904, '0969남양주EXP물류센터STOCK': 81040905,
        '0931덕평EXP물류센터FLOW': 81040906, '0934오산Exp물류센터FLOW': 81040907,
        '0935오산365물류센터STOCK': 81040908, '2001BH)영통점DSD': 81020192,
        '2002BH)강서점DSD': 81020191, '2003BH)인천송도점DSD': 81020190,
        '0934오산EXP물류센터SORTATION': 81040907, '0907밀양EXP센터SORTATION': 81021903,
        '0905기흥물류서비스센터SORTATION': 81021901, '0051강서점DSD': 81020191
    }
    NORMALIZED_STORE_MAP = {re.sub(r'^\d+', '', k).replace(" ", "").upper(): v for k, v in RAW_STORE_MAP.items()}

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
            col_map = {
                '상품명': row_strs.index('상품명') if '상품명' in row_strs else -1,
                '상품코드': row_strs.index('상품코드'),
                '입고타입': row_strs.index('입고타입') if '입고타입' in row_strs else -1,
                '수량': row_strs.index('낱개수량') if '낱개수량' in row_strs else -1,
                '단가': row_strs.index('낱개당 단가') if '낱개당 단가' in row_strs else -1,
                '금액': row_strs.index('발주금액') if '발주금액' in row_strs else -1,
                '납품처': row_strs.index('납품처') if '납품처' in row_strs else -1,
                '납품일자': row_strs.index('납품일자') if '납품일자' in row_strs else -1
            }
            continue
        if not col_map: continue
        try:
            b_idx = col_map['상품코드']
            b_str = re.sub(r'[^\d]', '', row_strs[b_idx])
            if not b_str: continue
            barcode = int(b_str)
            if barcode in FULL_PRODUCT_MAP:
                parsed_data.append({
                    '상품명': row_strs[col_map['상품명']] if col_map['상품명'] != -1 else '',
                    '바코드': barcode, '입고타입': row_strs[col_map['입고타입']] if col_map['입고타입'] != -1 else '',
                    '수량': float(re.sub(r'[^\d.]', '', row_strs[col_map['수량']])) if col_map['수량'] != -1 else 0.0,
                    '단가': float(re.sub(r'[^\d.]', '', row_strs[col_map['단가']])) if col_map['단가'] != -1 else 0.0,
                    '금액': float(re.sub(r'[^\d.]', '', row_strs[col_map['금액']])) if col_map['금액'] != -1 else 0.0,
                    '납품처': row_strs[col_map['납품처']] if col_map['납품처'] != -1 else '',
                    '납품일자': row_strs[col_map['납품일자']] if col_map['납품일자'] != -1 else ''
                })
        except: pass

    df = pd.DataFrame(parsed_data)
    if df.empty: return pd.DataFrame()
    
    df['ME코드'] = df['바코드'].map(FULL_PRODUCT_MAP)
    def get_store_code(row):
        s = re.sub(r'^\d+', '', str(row['납품처']).replace(' ', '').upper())
        t = str(row['입고타입']).replace(' ', '').upper()
        if 'HYPER_FLOW' in t: t = 'FLOW'
        elif 'MIX' in t: t = 'SORTATION'
        key = s + t
        return next((v for k, v in NORMALIZED_STORE_MAP.items() if k in key or key in k), 81040913)

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

    if '점포코드' not in raw_df.columns: return pd.DataFrame()

    raw_df = raw_df.dropna(subset=['점포코드'])
    raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)
    raw_df['센터코드'] = raw_df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False).str.strip()
    raw_df['수량'] = pd.to_numeric(raw_df.get('수량', 0), errors='coerce').fillna(0)
    
    date_col = next((c for c in ['센터입하일자', '센터입하일', '점입점일자'] if c in raw_df.columns), None)
    raw_df['납품일자'] = raw_df[date_col].astype(str).str.replace(r'[^0-9]', '', regex=True) if date_col else today_str

    emart_map_dict = {
        'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
        'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
        'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
    }
    
    # ⭐ 누락되었던 전체 이마트 배송처 맵핑 완벽 복구
    delivery_name_map = {
        '81010901': '이마트 백암물류센터', 
        '81010902': '이마트 시화물류센터', 
        '81010903': '이마트 대구물류센터',
        '81010905': '이마트 여주물류센터', 
        '81010906': '이마트 광주물류센터',
        '81010904': '이마트 노브랜드 여주2물류센터', 
        '81010968': '이마트 노브랜드 여주물류센터',
        '81010969': '이마트 노브랜드 시화물류센터', 
        '89011175': '이마트 노브랜드 대구물류(신규)',
        '81033036': '이마트 트레이더스 평택물류', 
        '89011174': '이마트 트레이더스 대구물류', 
        '81011012': '이마트 트레이더스 여주물류',
        '81011010': '이마트 트레이더스 시화물류'
    }

    def process_row(row):
        code, center = row['점포코드'], str(row['센터코드'])
        cust = 'E-mart' if (1000 <= code <= 1999 or code >= 9000) else ('E-mart(TRD)' if 2000 <= code <= 2999 else 'E-mart(노브랜드)')
        m_code = emart_map_dict.get(cust, {}).get(center, center)
        return pd.Series([cust, m_code])

    raw_df[['발주처', '배송코드']] = raw_df.apply(process_row, axis=1)
    raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
    
    merged = pd.merge(raw_df, prod_df, left_on='상품코드', right_on='바코드', how='left')
    merged['ME코드'] = merged['상품코드(기획)'].fillna(merged['상품코드'])
    name_col = '상품명(기획)' if '상품명(기획)' in prod_df.columns else '상품명'
    merged['상품명'] = merged[name_col].fillna(merged.get('상품명', ''))
    
    # 맵핑을 통해 배송처 이름 할당 (실패하면 배송코드 그대로 노출)
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

    parsed_list, curr_center, curr_doc_no, curr_date = [], "", "", ""
    
    for _, row in df_edi.dropna(how='all').iterrows():
        r = [str(x).strip() for x in row.tolist()]
        
        if r[0] == 'ORDERS':
            curr_doc_no = r[1].replace('.0', '')
            curr_center = re.sub(r'상온센타|상온센터|센타', '센터', r[5]).replace('센터센터', '센터')
            curr_date = re.sub(r'[^0-9]', '', r[7])
            continue
            
        if len(r) > 1 and r[1].startswith('880'):
            ipsu = int(extract_num(r[5])) or 1
            box_qty = int(extract_num(r[6]))
            qty = box_qty * ipsu
            price = extract_num(r[7])
            
            parsed_list.append({
                '발주코드': curr_doc_no, 
                '배송처': curr_center, 
                '납품일자': curr_date,
                'ME코드': r[1].replace('.0', ''), 
                '상품명': r[2], 
                '수량': qty, 
                '단가': price, 
                'Total Amount': qty * price
            })
    
    df = pd.DataFrame(parsed_list)
    if df.empty: return df
    
    df['배송코드'] = df['배송처'].map(lambda x: next((v for k, v in CENTER_MAP.items() if k in x), '81030000'))
    df['수주날짜'] = today_str
    df['발주처'] = '롯데마트'
    df['구분'] = "0"
    
    return df[FINAL_COLUMNS]

# ==========================================
# 🎨 사이드바 및 마스터 로드
# ==========================================
with st.sidebar:
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg", use_container_width=True)
    st.header("⚙️ 시스템 설정")
    
    @st.cache_data
    def load_master():
        files = [
            "NEW 이마트 서식파일_20260420납품.xlsx", 
            "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx", 
            "NEW 노브랜드_20260409납품.xlsx"
        ]
        appended = []
        for f in files:
            if os.path.exists(f):
                xls = pd.ExcelFile(f)
                target = xls.sheet_names[0]
                for s in xls.sheet_names:
                    if any(x in s for x in ['제품', '상품', '단가']):
                        target = s
                        break
                d = pd.read_excel(xls, sheet_name=target)
                d.columns = d.columns.astype(str).str.strip()
                appended.append(d)
        if not appended: return None
        df_m = pd.concat(appended, ignore_index=True)
        target_col = '바코드' if '바코드' in df_m.columns else ('상품코드' if '상품코드' in df_m.columns else None)
        if target_col:
            df_m['바코드'] = df_m[target_col].astype(str).str.replace('.0', '', regex=False).str.strip()
            df_m = df_m.drop_duplicates(subset=['바코드'])
            return df_m
        return None

    emart_master = load_master()
    if emart_master is not None: st.success("✅ 이마트 마스터 로드 완료")
    else: st.warning("⚠️ 이마트 마스터 파일 없음")
    st.info(f"📅 시스템 기준일: {today_str}")

# ==========================================
# 🚀 메인 대시보드
# ==========================================
st.title("📦 마트 통합 수주 자동 변환기")
st.markdown("> **Tesco, 이마트(TRD/노브랜드), 롯데마트** 파일을 구분 없이 한꺼번에 업로드하세요.")

uploaded_files = st.file_uploader("📂 발주서 파일들을 드래그하세요 (복수 선택 가능)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_results = []
    
    with st.spinner("🔄 파일을 분석하여 분류 중입니다..."):
        for f in uploaded_files:
            try:
                if f.name.endswith('.csv'):
                    sample_content = f.getvalue()[:2000].decode('utf-8-sig', errors='ignore')
                else:
                    sample_df = pd.read_excel(f, nrows=10)
                    sample_content = str(sample_df.columns.tolist()) + str(sample_df.values.tolist())
                f.seek(0)
                
                if 'ORDERS' in sample_content:
                    df_res = run_lotte_logic(f)
                    mart_name = "롯데마트"
                elif '점포코드' in sample_content or '센터입하' in sample_content:
                    df_res = run_emart_logic(f, emart_master) if emart_master is not None else pd.DataFrame()
                    mart_name = "이마트"
                else:
                    df_res = run_tesco_logic(f)
                    mart_name = "Tesco"
                
                if not df_res.empty:
                    all_results.append(df_res)
                    st.write(f"✔️ **{f.name}** -> {mart_name} 인식 완료 ({len(df_res)}건)")
            except Exception as e:
                st.error(f"❌ **{f.name}** 처리 중 오류 발생: {e}")

    if all_results:
        combined_df = pd.concat(all_results, ignore_index=True)

        # ⭐ 빈칸(NaN)을 공백 문자로 채워서 groupby 때 데이터가 통째로 사라지는 현상 방지
        combined_df = combined_df.fillna("")

        # 데이터 병합 (동일 항목 수량/금액 합산)
        group_cols = ['구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '단가']
        
        # dropna=False 를 주어 결측치가 있어도 그룹화 과정에서 행이 유실되지 않게 완벽 방어!
        final_df = combined_df.groupby(group_cols, dropna=False, as_index=False).agg({
            '수량': 'sum',
            'Total Amount': 'sum'
        })
        
        final_df = final_df[FINAL_COLUMNS]

        st.divider()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 총 병합 건수", f"{len(final_df):,} 건")
        c2.metric("🔢 총 납품 수량", f"{final_df['수량'].sum():,.0f} 개")
        c3.metric("💰 총 납품 금액", f"{final_df['Total Amount'].sum():,.0f} 원")

        with st.expander("👀 통합 병합 결과 미리보기", expanded=True):
            st.dataframe(final_df, use_container_width=True, height=500)
        
        st.download_button(
            label="📥 통합 병합본 엑셀 다운로드 (중복 수량 합산 완료)",
            data=to_excel_unified(final_df),
            file_name=f"마트통합수주_업로드용_{today_str}.xlsx",
            mime="application/vnd.ms-excel",
            type="primary",
            use_container_width=True
        )
