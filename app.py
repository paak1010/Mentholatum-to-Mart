import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime, date

# ==========================================
# ⚙️ 페이지 및 기본 설정 (Wide Layout)
# ==========================================
st.set_page_config(page_title="멘소래담 마트 통합 수주업로드", page_icon="🏢", layout="wide")

# 모든 날짜 형식을 하이픈 없이 YYYYMMDD로 통일
today_str = datetime.today().strftime("%Y%m%d")

# ==========================================
# 🎨 좌측 사이드바 (Sidebar) - 로고 및 안내
# ==========================================
with st.sidebar:
    # 요청하신 로고 이미지 삽입
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.header("💡 시스템 사용 안내")
    st.info("각 마트별 원본 발주서(Raw Data)를 업로드하면 표준 SCM 양식으로 자동 정제됩니다.")
    st.markdown("---")
    st.markdown("📌 **지원 확장자:** `.csv`, `.xls`, `.xlsx`")
    st.markdown("⚠️ **주의사항:** 서버(깃허브)에 마스터 맵핑 파일이 정상적으로 위치해야 배송처/단가가 완벽히 적용됩니다.")
    st.markdown("---")
    st.markdown(f"📅 **시스템 기준일:** `{today_str}`")

# ==========================================
# 📝 메인 화면 타이틀
# ==========================================
st.title("📦 통합 마트 수주 자동 변환 대시보드")
st.markdown("> **Tesco, 이마트 계열(TRD/노브랜드), 롯데마트**의 수주 데이터를 하나의 표준 양식으로 자동 병합·변환합니다.")
st.markdown("<br>", unsafe_allow_html=True)

# ⭐ 최종 통일 양식 컬럼 리스트 (구분 열 추가)
FINAL_COLUMNS = [
    '구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 
    'ME코드', '상품명', '수량', '단가', 'Total Amount'
]

def to_excel_unified(df, sheet_name="통합_수주업로드"):
    """데이터프레임을 엑셀 파일(메모리)로 변환하고 숫자 서식을 지정합니다."""
    
    # ⭐ [수정된 부분] 엑셀로 쓰기 전에 수량, 단가, Total Amount를 강제 숫자형으로 변환
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
        
        # 헤더 스타일 적용
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        for col_idx, col_name in enumerate(df.columns):
            if col_name in ['수량', '단가', 'Total Amount']:
                worksheet.set_column(col_idx, col_idx, 12, num_format)
            elif col_name in ['구분', '수주날짜', '납품일자', '발주코드', '배송코드']:
                worksheet.set_column(col_idx, col_idx, 14, center_format)
            elif col_name in ['상품명', '배송처']:
                worksheet.set_column(col_idx, col_idx, 30)
            else:
                worksheet.set_column(col_idx, col_idx, 15)
    return output.getvalue()

# ==========================================
# 🗂️ 대시보드 탭 디자인 적용
# ==========================================
tab_tesco, tab_emart, tab_lotte = st.tabs(["🔴 Tesco", "🟡 이마트 (TRD/노브랜드 포함)", "🟢 롯데마트"])

# =====================================================================
# 🔴 [TAB 1] TESCO 로직
# =====================================================================
with tab_tesco:
    st.subheader("🔴 Tesco 발주 데이터 업로드")
    
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
        '0906NEW함안상온물류센터SORTATION
