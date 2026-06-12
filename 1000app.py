import streamlit as st
import pandas as pd
import numpy as np
import os

# 1. 페이지 설정 및 브라우저 자동 번역 오작동 차단
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

# 깃허브 메뉴 등 상단 바 숨기기 (깔끔한 UI)
st.markdown("""
    <style>
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

MAJOR_MAPPING = {
    "의료보건(의/치/한/약/수의)": ["의예", "의학", "치의예", "치医学", "한의예", "한의학", "약학", "제약", "한약", "수의"],
    "간호/보건계열": ["간호", "물리치료", "방사선", "안경", "언어치료", "응급구조", "임상병리", "작업치료", "치기공", "치위생"],
    "문과(인문/어문/사학)": ["국문", "국어국문", "중문", "중어중문", "영문", "영어영문", "불문", "불어불문", "독문", "독어독문", "노문", "노어노문", "서문", "스페인", "역사", "사학", "철학", "미학", "고고", "미술사", "문화인류", "문헌정보", "통번역", "국제"],
    "사범(교육계열)": ["교육학", "국어교육", "영어교육", "불어교육", "독어교육", "사회교육", "역사교육", "지리교육", "윤리교육", "수학교육", "물리교육", "화학교육", "생물교육", "지구과학교육", "체육교육", "음악교육", "미술교육"],
    "이과(순수과학)": ["수학과", "화학과", "천문", "우주", "인공위성", "물리학과", "지구시스템", "대기"],
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
    
    # [핵심] 9등급 -> 5등급 누적 백분위 기반 정밀 환산 알고리즘
    def convert_9_to_5(grade9):
        if pd.isna(grade9): return np.nan
        # 9등급제 기준점 (등급 -> 백분위)
        g9_p = [1.0, 1.5, 2.5, 3.5, 4.5, 5.5, 6.5, 7.5, 8.5, 9.0]
        pct_p = [0, 4, 11, 23, 40, 60, 77, 89, 96, 100]
        # 5등급제 기준점 (백분위 -> 등급)
        g5_p = [1.0, 1.5, 2.5, 3.5, 4.5, 5.0]
        pct5_p = [0, 10, 34, 66, 90, 100]
        
        # 1. 9등급 점수를 백분위로 변환
        pct = np.interp(grade9, g9_p, pct_p)
        # 2. 백분위를 다시 5등급 점수로 변환
        g5 = np.interp(pct, pct5_p, g5_p)
        return round(g5, 2)
        
    df['cut_70_5g'] = df['cut_70'].apply(convert_9_to_5)
    df['cut_50_5g'] = df['cut_50'].apply(convert_9_to_5)
    return df

@st.cache_data
def load_curriculum_data():
    file_path = os.path.join(BASE_DIR, EXCEL_FILE_NAME)
    if not os.path.exists(file_path): return pd.DataFrame()
    try:
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None, engine='openpyxl')
        if df_raw.empty: return pd.DataFrame()
        df_data = df_raw.iloc[4:].copy()
        ncol = df_data.shape[1]
        
        df_cleaned = pd.DataFrame()
        mapping_indices = {'권역': 0, '지역': 1, '대학명': 2, '모집단위': 3, '핵심과목': 5, '권장과목': 6, '비고': 7}
        for k, idx in mapping_indices.items():
            if ncol > idx:
                df_cleaned[k] = df_data.iloc[:, idx]
            else:
                df_cleaned[k] = ""
                
        for col in df_cleaned.columns:
            df_cleaned[col] = df_cleaned[col].fillna("").astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
            
        df_cleaned['권역'] = df_cleaned['권역'].replace('', np.nan).ffill()
        df_cleaned['지역'] = df_cleaned['지역'].replace('', np.nan).ffill()
        df_cleaned['대학명'] = df_cleaned['대학명'].replace('', np.nan).ffill()
        df_cleaned = df_cleaned[df_cleaned['모집단위'] != ""]
        return df_cleaned
    except Exception:
        return pd.DataFrame()

df_cut = load_admission_data()
df_curr = load_curriculum_data()

st.title("🎯 대입 분석 및 진로교과 매핑 시스템")
st.caption("천명의선택 프리미엄 입시 컨설팅 고도화 연동 솔루션")
st.markdown("---")

if df_cut.empty or df_curr.empty:
    st.error("❌ 입시 데이터 파일을 찾을 수 없습니다. (2026 수시정리.csv / 권장과목.xlsx)")
    st.stop()

st.markdown("### 📋 다음 중 선택하세요")
menu = st.radio("원하시는 메뉴를 선택하세요.", ("희망대학 컷", "진로 교과 추천"), horizontal=True, label_visibility="collapsed")
st.markdown("---")

if menu == "희망대학 컷":
    st.header("🔍 안정 · 소신 · 도전 희망대학 컷 분석")
    
    # [추가됨] 9등급 / 5등급 선택 UI
    grade_system = st.radio(
        "📊 평가 체제 선택", 
        ("현 고3, N수생 (9등급제 기준)", "현 고1, 고2 (5등급제 환산 기준)"), 
        horizontal=True
    )
    is_5g = "5등급제" in grade_system
    
    if is_5g:
        st.info("💡 **[입결 데이터 안내]** 본 5등급제 환산 입결은 단순 추정이 아닌, 기존 9등급제 데이터를 누적 백분위 과학적 보간법(Interpolation)으로 산출한 **분석에 의한 예측 자료**입니다.")
    st.write("") # 간격 띄우기

    col1, col2 = st.columns(2)
    with col1:
        if is_5g:
            user_gpa = st.number_input("학생의 내신 평균 등급 (5등급제)을 입력하세요:", min_value=1.0, max_value=5.0, value=1.50, step=0.01)
        else:
            user_gpa = st.number_input("학생의 내신 평균 등급 (9등급제)을 입력하세요:", min_value=1.0, max_value=9.0, value=2.30, step=0.01)
        
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
        
        # 선택된 등급제에 따라 기준 컬럼 변경
        target_col = 'cut_70_5g' if is_5g else 'cut_70'
        
        valid_df = df_filtered.dropna(subset=[target_col]).sort_values(by=target_col).copy()
        
        if valid_df.empty:
            st.error("💡 현재, 지원 가능 대학의 자료가 부족합니다.")
        else:
            st.markdown(f"#### 📊 {selected_major} 계열 통합 리포트")
            highest_row = valid_df.iloc[0]
            lowest_row = valid_df.iloc[-1]
            
            unit = "등급 (5G)" if is_5g else "등급 (9G)"
            
            idx_col1, idx_col2 = st.columns(2)
            with idx_col1:
                st.metric("📈 계열 내 최고 입결 대학 (70%컷)", f"{highest_row['대학명']} ({highest_row['학과명']})", f"{highest_row[target_col]:.2f} {unit}", delta_color="inverse")
            with idx_col2:
                st.metric("📉 계열 내 최저 입결 대학 (70%컷)", f"{lowest_row['대학명']} ({lowest_row['학과명']})", f"{lowest_row[target_col]:.2f} {unit}")
            st.markdown("---")
            
            # 5등급제는 등급 촘촘함이 다르므로 margin(오차범위)을 다르게 적용
            if is_5g:
                s_margin, a_up, a_down, c_up, c_down = 0.25, 0.25, 0.15, 0.15, 0.6
            else:
                s_margin, a_up, a_down, c_up, c_down = 0.5, 0.5, 0.4, 0.4, 1.0
                
            stable_match = valid_df[valid_df[target_col] >= (user_gpa + s_margin)].copy()
            ambitious_match = valid_df[(valid_df[target_col] >= (user_gpa - a_down)) & (valid_df[target_col] < (user_gpa + a_up))].copy()
            challenge_match = valid_df[(valid_df[target_col] >= (user_gpa - c_down)) & (valid_df[target_col] < (user_gpa - c_up))].copy()
            
            if stable_match.empty: stable_match = valid_df[valid_df[target_col] > user_gpa].tail(1).copy()
            if ambitious_match.empty:
                valid_df['diff'] = (valid_df[target_col] - user_gpa).abs()
                ambitious_match = valid_df.sort_values(by='diff').head(1).copy()
                if 'diff' in ambitious_match.columns: ambitious_match = ambitious_match.drop(columns=['diff'])
            if challenge_match.empty: challenge_match = valid_df[valid_df[target_col] < user_gpa].head(1).copy()
                
            # 출력 컬럼명 동적 변경
            if is_5g:
                display_cols = ['지역', '대학명', '전형명', '학과명', 'cut_50_5g', 'cut_70_5g']
                rename_cols = {'지역':'지역', '대학명':'대학명', '전형명':'전형 종류 (농어촌 등)', '학과명':'학과명', 'cut_50_5g':'50% 합격컷(5등급)', 'cut_70_5g':'70% 합격컷(5등급)'}
            else:
                display_cols = ['지역', '대학명', '전형명', '학과명', '2025등급컷2', '2025등급컷']
                rename_cols = {'지역':'지역', '대학명':'대학명', '전형명':'전형 종류 (농어촌 등)', '학과명':'학과명', '2025등급컷2':'50% 합격컷(9등급)', '2025등급컷':'70% 합격컷(9등급)'}
            
            st.subheader("🟢 안정 지원 선 (합격 안착 확률 최상)")
            if not stable_match.empty:
                st.dataframe(stable_match[display_cols].rename(columns=rename_cols).reset_index(drop=True), use_container_width=True)
            else:
                st.write("현재 지원 가능 대학의 자료가 부족합니다.")
                
            st.subheader("🟡 소신 지원 선 (적정 적합 분석권)")
            if not ambitious_match.empty:
                st.dataframe(ambitious_match[display_cols].rename(columns=rename_cols).reset_index(drop=True), use_container_width=True)
            else:
                st.write("현재 지원 가능 대학의 자료가 부족합니다.")
                
            st.subheader("🔴 도전 지원 선 (상향 예측 추적권)")
            if not challenge_match.empty:
                st.dataframe(challenge_match[display_cols].rename(columns=rename_cols).reset_index(drop=True), use_container_width=True)
            else:
                st.write("현재 지원 가능 대학의 자료가 부족합니다.")

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
