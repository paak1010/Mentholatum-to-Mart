import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime

# ==========================================
# ⚙️ 페이지 및 기본 설정 (Wide Layout)
# ==========================================
st.set_page_config(page_title="통합 수주업로드 시스템", page_icon="🏢", layout="wide")

# ==========================================
# 🎨 B2B SaaS 스타일 커스텀 CSS
# ==========================================
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"]  {
        font-family: 'Pretendard', sans-serif !important;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .stApp { background-color: #f8fafc; }
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    
    /* 단일 업로드 창 강조 디자인 */
    [data-testid="stFileUploadDropzone"] {
        border-radius: 12px;
        border: 2px dashed #64748b;
        background-color: #ffffff;
        padding: 40px;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #3b82f6;
        background-color: #eff6ff;
    }
    
    /* 메트릭 및 결과 박스 */
    .result-box {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

today_str = datetime.today().strftime("%Y%m%d")
FINAL_COLUMNS = ['구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '수량', '단가', 'Total Amount']

# ==========================================
# 🔍 파일 포맷 자동 감지 (Smart Routing Logic)
# ==========================================
def detect_vendor_format(uploaded_file):
    """파일의 헤더나 첫 데이터 줄을 읽어 벤더사를 자동 식별합니다."""
    filename = uploaded_file.name.lower()
    uploaded_file.seek(0) # 포인터 초기화
    
    # 1. 텍스트/CSV 기반 고유 키워드 스캔 (롯데마트 EDI)
    if filename.endswith(('.csv', '.txt')):
        try:
            content = uploaded_file.read(2000).decode('utf-8-sig', errors='ignore')
            if 'ORDERS' in content:
                uploaded_file.seek(0)
                return "LOTTE", "롯데마트 EDI"
            uploaded_file.seek(0)
        except Exception:
            uploaded_file.seek(0)

    # 2. DataFrame 헤더 스캔 (엑셀/CSV)
    try:
        if filename.endswith('.csv'):
            try:
                df_preview = pd.read_csv(uploaded_file, nrows=5, encoding='utf-8-sig')
            except:
                uploaded_file.seek(0)
                df_preview = pd.read_csv(uploaded_file, nrows=5, encoding='cp949')
        else:
            # 엑셀의 경우 첫 번째 시트의 구조 확인
            xls = pd.ExcelFile(uploaded_file)
            sheet_name = xls.sheet_names[0]
            for s in xls.sheet_names:
                temp = pd.read_excel(xls, sheet_name=s, nrows=3)
                if '점포코드' in temp.columns:
                    sheet_name = s
                    break
            df_preview = pd.read_excel(xls, sheet_name=sheet_name, nrows=5)

        columns_str = "".join(df_preview.columns.astype(str))
        first_row_str = "".join(df_preview.iloc[0].astype(str)) if len(df_preview) > 0 else ""

        # 판별 로직
        if '점포코드' in columns_str or '센터코드' in columns_str:
            uploaded_file.seek(0)
            return "EMART", "이마트/TRD/노브랜드"
        elif 'ORDERS' in first_row_str:
            uploaded_file.seek(0)
            return "LOTTE", "롯데마트 EDI"
        elif '상품명' in columns_str and ('발주금액' in columns_str or '낱개수량' in columns_str):
            uploaded_file.seek(0)
            return "TESCO", "Tesco"
        
    except Exception as e:
        pass

    uploaded_file.seek(0)
    # 특징이 명확하지 않으면 기본적으로 포괄적인 탐색을 하는 로직으로 빠지게 하거나 에러 반환
    return "UNKNOWN", "미상"

# ==========================================
# ⚙️ 벤더별 처리 코어 함수 (간략화된 뼈대)
# ==========================================
# 기존에 작성하신 코어 로직들을 함수화하여 호출만 되도록 분리합니다.

def process_tesco(file):
    # 기존 Tesco 로직 복사 후 데이터프레임(df_final) 리턴
    # ... (기존 로직) ...
    # 예시 반환값
    return pd.DataFrame(columns=FINAL_COLUMNS) 

def process_emart(file):
    # 기존 Emart 로직 복사 후 데이터프레임(df_final) 리턴
    return pd.DataFrame(columns=FINAL_COLUMNS)

def process_lotte(file):
    # 기존 Lotte 로직 복사 후 데이터프레임(df_final) 리턴
    return pd.DataFrame(columns=FINAL_COLUMNS)

def to_excel_unified(df, vendor_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=f"{vendor_name}_수주")
    return output.getvalue()

# ==========================================
# 🖥️ 메인 UI 및 인터랙션
# ==========================================
with st.sidebar:
    st.markdown("### 🏢 SCM 통합 대시보드")
    st.markdown("---")
    st.markdown(f"📅 **시스템 기준일:** `{today_str}`")
    st.markdown("💡 **Tip:** 파일 양식을 굳이 분류할 필요 없이 메인 화면에 업로드하면 시스템이 자동으로 벤더를 인식하여 매핑합니다.")

st.title("통합 마트 수주 자동 변환 시스템")
st.markdown("> **Tesco, 이마트 계열, 롯데마트** 등 거래처 상관없이 발주 파일을 업로드해 주세요.")
st.markdown("<br>", unsafe_allow_html=True)

# 1. 단일 파일 업로드 드롭존
uploaded_file = st.file_uploader("", type=['xlsx', 'xls', 'csv'], help="거래처 원본 수주 파일을 이곳에 끌어다 놓으세요.")

if uploaded_file:
    # 2. 파일 타입 자동 감지
    with st.spinner("🔍 파일 서식을 분석 중입니다..."):
        vendor_code, vendor_name = detect_vendor_format(uploaded_file)
    
    if vendor_code == "UNKNOWN":
        st.error("⚠️ 업로드된 파일의 거래처 양식을 식별할 수 없습니다. 컬럼명이나 형식이 변경되었는지 확인해 주세요.")
    else:
        st.success(f"✅ **[{vendor_name}]** 발주 양식이 감지되었습니다. 자동 병합을 시작합니다.")
        
        # 3. 진행 상태 표시
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.write("⏳ 마스터 데이터 매핑 및 수주 규격 변환 중...")
            progress_bar.progress(30)
            
            # 4. 식별된 벤더에 맞는 프로세스 라우팅
            if vendor_code == "TESCO":
                df_result = process_tesco(uploaded_file)
            elif vendor_code == "EMART":
                df_result = process_emart(uploaded_file)
            elif vendor_code == "LOTTE":
                df_result = process_lotte(uploaded_file)
                
            progress_bar.progress(90)
            status_text.write("⏳ 최종 엑셀 서식 포맷팅 중...")
            
            if df_result is not None and not df_result.empty:
                progress_bar.progress(100)
                status_text.empty() # 진행 텍스트 삭제
                
                # 5. 결과 요약 및 다운로드 UI
                st.markdown('<div class="result-box">', unsafe_allow_html=True)
                st.subheader(f"📊 {vendor_name} 처리 결과 요약")
                
                c1, c2, c3 = st.columns(3)
                c1.metric("📦 총 처리 건수", f"{len(df_result):,} 건")
                c2.metric("🔢 총 납품 수량", f"{df_result['수량'].sum():,.0f} 개")
                c3.metric("💰 총 납품 금액", f"{df_result['Total Amount'].sum():,.0f} 원")
                
                with st.expander("👀 변환된 상세 데이터 확인 (Preview)", expanded=True):
                    st.dataframe(df_result.head(50), use_container_width=True)
                
                st.download_button(
                    label=f"📥 ERP용 통일 양식 다운로드 ({vendor_name})",
                    data=to_excel_unified(df_result, vendor_name),
                    file_name=f"수주통합본_{vendor_code}_{today_str}.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                progress_bar.empty()
                st.warning("처리할 유효한 발주 데이터가 없습니다.")
                
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"❌ 데이터 변환 중 오류가 발생했습니다: {str(e)}")

]import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime

# ==========================================
# ⚙️ 페이지 및 기본 설정 (Wide Layout)
# ==========================================
st.set_page_config(page_title="통합 수주업로드 시스템", page_icon="🏢", layout="wide")

# ==========================================
# 🎨 B2B SaaS 스타일 커스텀 CSS
# ==========================================
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"]  {
        font-family: 'Pretendard', sans-serif !important;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .stApp { background-color: #f8fafc; }
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    
    /* 단일 업로드 창 강조 디자인 */
    [data-testid="stFileUploadDropzone"] {
        border-radius: 12px;
        border: 2px dashed #64748b;
        background-color: #ffffff;
        padding: 40px;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #3b82f6;
        background-color: #eff6ff;
    }
    
    /* 메트릭 및 결과 박스 */
    .result-box {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

today_str = datetime.today().strftime("%Y%m%d")
FINAL_COLUMNS = ['구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '수량', '단가', 'Total Amount']

# ==========================================
# 🔍 파일 포맷 자동 감지 (Smart Routing Logic)
# ==========================================
def detect_vendor_format(uploaded_file):
    """파일의 헤더나 첫 데이터 줄을 읽어 벤더사를 자동 식별합니다."""
    filename = uploaded_file.name.lower()
    uploaded_file.seek(0) # 포인터 초기화
    
    # 1. 텍스트/CSV 기반 고유 키워드 스캔 (롯데마트 EDI)
    if filename.endswith(('.csv', '.txt')):
        try:
            content = uploaded_file.read(2000).decode('utf-8-sig', errors='ignore')
            if 'ORDERS' in content:
                uploaded_file.seek(0)
                return "LOTTE", "롯데마트 EDI"
            uploaded_file.seek(0)
        except Exception:
            uploaded_file.seek(0)

    # 2. DataFrame 헤더 스캔 (엑셀/CSV)
    try:
        if filename.endswith('.csv'):
            try:
                df_preview = pd.read_csv(uploaded_file, nrows=5, encoding='utf-8-sig')
            except:
                uploaded_file.seek(0)
                df_preview = pd.read_csv(uploaded_file, nrows=5, encoding='cp949')
        else:
            # 엑셀의 경우 첫 번째 시트의 구조 확인
            xls = pd.ExcelFile(uploaded_file)
            sheet_name = xls.sheet_names[0]
            for s in xls.sheet_names:
                temp = pd.read_excel(xls, sheet_name=s, nrows=3)
                if '점포코드' in temp.columns:
                    sheet_name = s
                    break
            df_preview = pd.read_excel(xls, sheet_name=sheet_name, nrows=5)

        columns_str = "".join(df_preview.columns.astype(str))
        first_row_str = "".join(df_preview.iloc[0].astype(str)) if len(df_preview) > 0 else ""

        # 판별 로직
        if '점포코드' in columns_str or '센터코드' in columns_str:
            uploaded_file.seek(0)
            return "EMART", "이마트/TRD/노브랜드"
        elif 'ORDERS' in first_row_str:
            uploaded_file.seek(0)
            return "LOTTE", "롯데마트 EDI"
        elif '상품명' in columns_str and ('발주금액' in columns_str or '낱개수량' in columns_str):
            uploaded_file.seek(0)
            return "TESCO", "Tesco"
        
    except Exception as e:
        pass

    uploaded_file.seek(0)
    # 특징이 명확하지 않으면 기본적으로 포괄적인 탐색을 하는 로직으로 빠지게 하거나 에러 반환
    return "UNKNOWN", "미상"

# ==========================================
# ⚙️ 벤더별 처리 코어 함수 (간략화된 뼈대)
# ==========================================
# 기존에 작성하신 코어 로직들을 함수화하여 호출만 되도록 분리합니다.

def process_tesco(file):
    # 기존 Tesco 로직 복사 후 데이터프레임(df_final) 리턴
    # ... (기존 로직) ...
    # 예시 반환값
    return pd.DataFrame(columns=FINAL_COLUMNS) 

def process_emart(file):
    # 기존 Emart 로직 복사 후 데이터프레임(df_final) 리턴
    return pd.DataFrame(columns=FINAL_COLUMNS)

def process_lotte(file):
    # 기존 Lotte 로직 복사 후 데이터프레임(df_final) 리턴
    return pd.DataFrame(columns=FINAL_COLUMNS)

def to_excel_unified(df, vendor_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=f"{vendor_name}_수주")
    return output.getvalue()

# ==========================================
# 🖥️ 메인 UI 및 인터랙션
# ==========================================
with st.sidebar:
    st.markdown("### 🏢 SCM 통합 대시보드")
    st.markdown("---")
    st.markdown(f"📅 **시스템 기준일:** `{today_str}`")
    st.markdown("💡 **Tip:** 파일 양식을 굳이 분류할 필요 없이 메인 화면에 업로드하면 시스템이 자동으로 벤더를 인식하여 매핑합니다.")

st.title("통합 마트 수주 자동 변환 시스템")
st.markdown("> **Tesco, 이마트 계열, 롯데마트** 등 거래처 상관없이 발주 파일을 업로드해 주세요.")
st.markdown("<br>", unsafe_allow_html=True)

# 1. 단일 파일 업로드 드롭존
uploaded_file = st.file_uploader("", type=['xlsx', 'xls', 'csv'], help="거래처 원본 수주 파일을 이곳에 끌어다 놓으세요.")

if uploaded_file:
    # 2. 파일 타입 자동 감지
    with st.spinner("🔍 파일 서식을 분석 중입니다..."):
        vendor_code, vendor_name = detect_vendor_format(uploaded_file)
    
    if vendor_code == "UNKNOWN":
        st.error("⚠️ 업로드된 파일의 거래처 양식을 식별할 수 없습니다. 컬럼명이나 형식이 변경되었는지 확인해 주세요.")
    else:
        st.success(f"✅ **[{vendor_name}]** 발주 양식이 감지되었습니다. 자동 병합을 시작합니다.")
        
        # 3. 진행 상태 표시
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.write("⏳ 마스터 데이터 매핑 및 수주 규격 변환 중...")
            progress_bar.progress(30)
            
            # 4. 식별된 벤더에 맞는 프로세스 라우팅
            if vendor_code == "TESCO":
                df_result = process_tesco(uploaded_file)
            elif vendor_code == "EMART":
                df_result = process_emart(uploaded_file)
            elif vendor_code == "LOTTE":
                df_result = process_lotte(uploaded_file)
                
            progress_bar.progress(90)
            status_text.write("⏳ 최종 엑셀 서식 포맷팅 중...")
            
            if df_result is not None and not df_result.empty:
                progress_bar.progress(100)
                status_text.empty() # 진행 텍스트 삭제
                
                # 5. 결과 요약 및 다운로드 UI
                st.markdown('<div class="result-box">', unsafe_allow_html=True)
                st.subheader(f"📊 {vendor_name} 처리 결과 요약")
                
                c1, c2, c3 = st.columns(3)
                c1.metric("📦 총 처리 건수", f"{len(df_result):,} 건")
                c2.metric("🔢 총 납품 수량", f"{df_result['수량'].sum():,.0f} 개")
                c3.metric("💰 총 납품 금액", f"{df_result['Total Amount'].sum():,.0f} 원")
                
                with st.expander("👀 변환된 상세 데이터 확인 (Preview)", expanded=True):
                    st.dataframe(df_result.head(50), use_container_width=True)
                
                st.download_button(
                    label=f"📥 ERP용 통일 양식 다운로드 ({vendor_name})",
                    data=to_excel_unified(df_result, vendor_name),
                    file_name=f"수주통합본_{vendor_code}_{today_str}.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                progress_bar.empty()
                st.warning("처리할 유효한 발주 데이터가 없습니다.")
                
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"❌ 데이터 변환 중 오류가 발생했습니다: {str(e)}")
