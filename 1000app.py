import streamlit as st
import pandas as pd
import numpy as np
import os

# 1. 페이지 설정 및 브라우저 자동 번역 차단
st.set_page_config(page_title="대입 전형 및 권역별 권장과목 분석 시스템", layout="wide")
st.markdown(
    """
    <html lang="ko" class="notranslate" google="notranslate">
    <head><meta name="google" content="notranslate" /></head>
    </html>
    """, 
    unsafe_allow_html=True
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
CUT_FILE_NAME = "2026 수시정리.csv"
EXCEL_FILE_NAME = "2028학년도 권역별 대학별 권장과목(반영과목).xlsx"

MAJOR_MAPPING = {
    "의료보건(의/치/한/약/수의)": ["의예", "의학", "치의예", "치医学", "한의예", "한의학", "약학", "제약", "한약", "수의"],
    "간호/보건계열": ["간호", "물리치료", "방사선", "안경", "언어치료", "응급구조", "임상병리", "작업치료", "치기공", "치위생"],
    "문과(인문/어문/사학)": ["국문", "국어국문", "중문", "중어중문", "영문", "영어영문", "불문", "불어불문", "독문", "독어독문", "노문", "노어노문", "서문", "스페인", "역사", "사학", "철학", "미학", "고고", "미술사", "문화인류", "문헌정보", "통번역", "국제"],
    "사범(교육계열)": ["교육학", "국어교육", "영어교육", "불어교육", "독어교육", "사회교육", "역사교육", "지리교육", "윤리교육", "수학교육", "물리교육", "화학교육", "생물교육", "지구과학교육", "체육교육", "음악교육", "미술교육"],
    "이과(순수과학)": ["수학", "화학", "천문", "우주", "인공위성", "물리", "지구시스템", "대기"],
    "공과(신소재/건설/제조)": ["반도체", "신소재", "재료", "건설환경", "토목", "건축", "화공", "화학공학", "화공생명", "전기", "전자", "도시공학", "기계", "산업공학", "디스플레이", "조선", "해양공학", "원자력", "에너지", "항공우주"],
    "컴퓨터/AI/소프트웨어": ["컴퓨터", "인공지능", "AI", "IT융합", "영상예술", "소프트웨어", "정보보안", "보안"],
    "자연(바이오/생명/농림)": ["생물학", "생화학", "생명공학", "생명과학", "통계", "유전", "원예", "식물생산", "산림", "응용생물", "동물공학", "바이오시스템", "조경"],
    "사회(정치/행정/언론)": ["정치외교", "정외", "사회학", "행정", "사회복지", "언론홍보", "미디어", "경제", "지리", "러시아학", "일본학", "유럽", "지역학", "군사학", "아동", "소비자", "의류", "식품영양"],
    "법경제(경영/상경/관광)": ["법학", "부동산", "경영", "회계", "세무", "무역", "국제통상", "금융", "빅데이터응용", "호텔", "조리", "관광", "문화엔터테인먼트", "벤처"],
    "예체능": []
}

@st.cache_data
def load_admission_data():
    file_path = os.path.join(BASE_DIR, CUT_FILE_NAME)
    if not os.path.exists(file_path): return pd.DataFrame()
    try: df = pd.read_csv(file_path, encoding="utf-8")
    except Exception: df = pd.read_csv(file_path, encoding="cp949")
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == 'object': df[col] = df[col].str.strip()
    df['cut_70'] = pd.to_numeric(df['2025등급컷'], errors='coerce')
    df['cut_50'] = pd.to_numeric(df['2025등급컷2'], errors='coerce')
    return df

@st.cache_data
def load_curriculum_data():
    file_path = os.path.join(BASE_DIR, EXCEL_FILE_NAME)
    if not os.path.exists(file_path): return pd.DataFrame()
    try: df_raw = pd.read_excel(file_path, sheet_name=0, header=None, engine='openpyxl')
    except Exception: return pd.DataFrame()
    
    df_data = df_raw.iloc[4:].copy()
    df_cleaned = pd.DataFrame()
    df_cleaned['권역'] = df_data.iloc[:, 0] if df_data.shape[1] > 0 else ""
    df_cleaned['지역'] = df_data.iloc[:, 1] if df_data.shape[1] > 1 else ""
    df_cleaned['대학명'] = df_data.iloc[:, 2] if df_data.shape[1] > 2 else ""
    df_cleaned['모집단위'] = df_data.iloc[:, 3] if df_data.shape[1] > 3 else ""
    df_cleaned['핵심과목'] = df_data.iloc[:, 5] if df_data.shape[1] > 5 else ""
    df_cleaned['권장과목'] = df_data.iloc[:, 6] if df_data.shape[1] > 6 else ""
    df_cleaned['비고'] = df_data.iloc[:, 7] if df_data.shape[1] > 7 else ""
    
    for col in df_cleaned.columns:
        df_cleaned[col] = df_cleaned[col].fillna("").astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        
    df_cleaned['권역'] = df_cleaned['권역'].replace('', np.nan).ffill()
    df_cleaned['지역'] = df_cleaned['지역'].replace('', np.nan).ffill()
    df_cleaned['대학명'] = df_cleaned['대학명'].replace('', np.nan).ffill()
    df_cleaned = df_cleaned[df_cleaned['모집단위'] != ""]
    return df_cleaned

df_cut = load_admission_data()
df_curr = load_curriculum_data()

st.title("🎯 대입 분석 및 진로교과 매핑 시스템")
st.caption("천명의선택 프리미엄 입시 컨설팅 연동 솔루션")
st.markdown("---")

if df_cut.empty or df_curr.empty:
    st.error("❌ 입시 데이터 파일을 찾을 수 없습니다.")
    st.markdown(f"**현재 경로:** `{BASE_DIR}`")
    st.info(f"1. {CUT_FILE_NAME}\n2. {EXCEL_FILE_NAME}\n위 두 파일이 폴더에 있는지 확인해주세요.")
    st.stop()

st.markdown("### 📋 다음 중 선택하세요")
menu = st.radio("원하시는 메뉴를 선택하세요.", ("희망대학 컷", "진로 교과 추천"), horizontal=True, label_visibility="collapsed")
st.markdown("---")

if menu == "희망대학 컷":
    st.header("🔍 안정 · 소신 · 도전 희망대학 컷 분석")
    
    col1, col2 = st.columns(2)
    with col1:
        user_gpa = st.number_input("학생의 내신 평균 등급을 입력하세요:", min_value=1.0, max_value=9.0, value=2.3, step=0.01)
        selected_major = st.selectbox("분석 대상 계열을 선택하세요:", list(MAJOR_MAPPING.keys()))
    with col2:
        all_regions = sorted([r for r in df_cut['지역'].dropna().unique() if r != ""])
        selected_regions = st.multiselect("대상 지역을 선택하세요:", options=all_regions, default=all_regions[:1])
        
    if selected_major == "예체능":
        st.warning("⚠️ 해당 계열에 대한 자료는 불충분합니다.")
    elif not selected_regions:
        st.warning("⚠️ 지역을 1개 이상 선택해 주세요.")
    else:
        keywords = MAJOR_MAPPING[selected_major]
        pattern = "|".join(keywords)
        df_filtered = df_cut[df_cut['지역'].isin(selected_regions) & df_cut['학과명'].str.contains(pattern, na=False)].copy()
        valid_df = df_filtered.dropna(subset=['cut_70']).sort_values(by='cut_70')
        
        if valid_df.empty:
            st.error("💡 현재, 지원 가능 대학의 자료가 부족합니다.")
        else:
            st.markdown(f"#### 📊 {selected_major} 계열 통합 리포트")
            highest_row = valid_df.iloc[0]
            lowest_row = valid_df.iloc[-1]
            
            idx_col1, idx_col2 = st.columns(2)
            with idx_col1:
                st.metric("📈 계열 내 최고 입결 대학 (70%컷)", f"{highest_row['대학명']} ({highest_row['학과명']})", f"{highest_row['cut_70']} 등급", delta_color="inverse")
            with idx_col2:
                st.metric("📉 계열 내 최저 입결 대학 (70%컷)", f"{lowest_row['대학명']} ({lowest_row['학과명']})", f"{lowest_row['cut_70']} 등급")
            st.markdown("---")
            
            stable_match = valid_df[valid_df['cut_70'] >= (user_gpa + 0.5)]
            ambitious_match = valid_df[(valid_df['cut_70'] >= (user_gpa - 0.4)) & (valid_df['cut_70'] < (user_gpa + 0.5))]
            challenge_match = valid_df[(valid_df['cut_70'] >= (user_gpa - 1.0)) & (valid_df['cut_70'] < (user_gpa - 0.4))]
            
            if stable_match.empty: stable_match = valid_df[valid_df['cut_70'] > user_gpa].tail(1)
            if ambitious_match.empty:
                valid_df['diff'] = (valid_df['cut_70'] - user_gpa).abs()
                ambitious_match = valid_df.sort_values(by='diff').head(1).drop(columns=['diff'])
            if challenge_match.empty: challenge_match = valid_df[valid_df['cut_70'] < user_gpa].head(1)
                
            display_cols = ['지역', '대학명', '전형명', '학과명', '2025등급컷2', '2025등급컷']
            rename_cols = {'지역':'지역', '대학명':'대학명', '전형명':'전형 종류', '학과명':'학과명', '2025등급컷2':'50% 합격컷', '2025등급컷':'70% 합격컷'}
            
            st.subheader("🟢 안정 지원 선")
            st.dataframe(stable_match[display_cols].rename(columns=rename_cols).reset_index(drop=True), use_container_width=True)
            st.subheader("🟡 소신 지원 선")
            st.dataframe(ambitious_match[display_cols].rename(columns=rename_cols).reset_index(drop=True), use_container_width=True)
            st.subheader("🔴 도전 지원 선")
            st.dataframe(challenge_match[display_cols].rename(columns=rename_cols).reset_index(drop=True), use_container_width=True)

elif menu == "진로 교과 추천":
    st.header("📚 모집단위별 고교 핵심 / 권장선택과목 매핑")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        region_options = sorted([r for r in df_curr['지역'].unique() if r != ""])
        selected_reg = st.selectbox("1. 지역 선택", options=region_options)
    df_filtered_reg = df_curr[df_curr['지역'] == selected_reg]
    with col2:
        univ_options = sorted([u for u in df_filtered_reg['대학명'].unique() if u != ""])
        selected_univ = st.selectbox("2. 대학명 선택", options=univ_options)
    df_filtered_univ = df_filtered_reg[df_filtered_reg['대학명'] == selected_univ]
    with col3:
        dept_options = sorted([d for d in df_filtered_univ['모집단위'].unique() if d != ""])
        selected_dept = st.selectbox("3. 학과 선택", options=dept_options)
        
    target_rows = df_filtered_univ[df_filtered_univ['모집단위'] == selected_dept]
    
    if not target_rows.empty:
        final_row = target_rows.iloc[0]
        st.markdown("---")
        st.subheader(f"📋 {selected_univ} - [{selected_dept}] 이수 지표")
        
        def format_blank_cell(value):
            v_str = str(value).strip()
            if v_str in ["", "-", "nan", "NaN", "None", "- ."]: return " "
            return v_str
            
        core_subject = format_blank_cell(final_row.get('핵심과목', ' '))
        recommended_subject = format_blank_cell(final_row.get('권장과목', ' '))
        note_value = format_blank_cell(final_row.get('비고', ' '))
        
        output_df = pd.DataFrame({
            "반영 구분": ["핵심과목", "권장과목", "비고"],
            "고교 선택 교과목 가이드라인": [core_subject, recommended_subject, note_value]
        })
        st.table(output_df)