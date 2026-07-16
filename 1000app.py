import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import os

# 1. 페이지 설정 및 브라우저 자동 번역 오작동 차단 (카카오톡 공유 제목 일관성 유지)
st.set_page_config(page_title="천명의선택 입시 NAVI", layout="wide")

# 2. 카카오톡 미리보기 고정을 위한 Open Graph 메타 태그 강제 주입
meta_tags = """
<head>
    <meta property="og:title" content="천명의선택 입시 NAVI" />
    <meta property="og:description" content="천명의선택 프리미엄 입시 컨설팅 고도화 연동 솔루션" />
    <meta property="og:type" content="website" />
    <html lang="ko" class="notranslate" google="notranslate">
    <meta name="google" content="notranslate" />
</head>
"""
components.html(meta_tags, height=0)

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
CUT_FILE_NAME = "2026 수시정리.csv"
EXCEL_FILE_NAME = "2028학년도 권역별 대학별 권장과목(반영과목).xlsx"

# 깃허브 메뉴 등 상단 바 숨기기 및 타이틀 한 줄 강제 정렬 스타일 적용
st.markdown("""
    <style>
    header {visibility: hidden;}
    .brand-title {
        font-size: calc(1.6rem + 1.2vw) !important;
        font-weight: 800 !important;
        white-space: nowrap !important;
        word-break: keep-all !important;
        margin-bottom: 5px;
    }
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

# 💡 9등급 -> 5등급 누적 백분위 기반 정밀 환산 알고리즘
def convert_9_to_5(grade9):
    if pd.isna(grade9) or str(grade9).strip() in ["", "-", "nan", "NaN", "None", "- ."]:
        return np.nan
    try:
        val = float(grade9)
        # 등급 범위 외 노이즈 차단
        if val < 1.0 or val > 9.0:
            return np.nan
        g9_p = [1.0, 1.5, 2.5, 3.5, 4.5, 5.5, 6.5, 7.5, 8.5, 9.0]
        pct_p = [0, 4, 11, 23, 40, 60, 77, 89, 96, 100]
        g5_p = [1.0, 1.5, 2.5, 3.5, 4.5, 5.0]
        pct5_p = [0, 10, 34, 66, 90, 100]
        
        pct = np.interp(val, g9_p, pct_p)
        g5 = np.interp(pct, pct5_p, g5_p)
        return round(g5, 2)
    except:
        return np.nan

@st.cache_data(ttl=3600)
def load_admission_data():
    file_path = os.path.join(BASE_DIR, CUT_FILE_NAME)
    if not os.path.exists(file_path): return pd.DataFrame()
    try: df = pd.read_csv(file_path, encoding="utf-8-sig")
    except Exception: df = pd.read_csv(file_path, encoding="cp949")
    
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == 'object': df[col] = df[col].str.strip()
        
    # 2025학년도 등급컷 정제 및 숫자 매핑
    df['cut_70'] = pd.to_numeric(df.get('2025등급컷', np.nan), errors='coerce')
    df['cut_50'] = pd.to_numeric(df.get('2025등급컷2', np.nan), errors='coerce')
    
    df['cut_70_5g'] = df['cut_70'].apply(convert_9_to_5)
    df['cut_50_5g'] = df['cut_50'].apply(convert_9_to_5)
    
    # 2026학년도 등급컷I, II, III 정밀 정제 및 환산
    if '2026등급컷I' in df.columns:
        df['cut_26_1_numeric'] = pd.to_numeric(df['2026등급컷I'], errors='coerce')
        df['cut_26_1_5g'] = df['cut_26_1_numeric'].apply(convert_9_to_5)
    else:
        df['cut_26_1_5g'] = np.nan

    if '2026등급컷II' in df.columns:
        df['cut_26_2_numeric'] = pd.to_numeric(df['2026등급컷II'], errors='coerce')
        df['cut_26_2_5g'] = df['cut_26_2_numeric'].apply(convert_9_to_5)
    else:
        df['cut_26_2_5g'] = np.nan

    if '2026등급컷III' in df.columns:
        df['cut_26_3_numeric'] = pd.to_numeric(df['2026등급컷III'], errors='coerce')
        df['cut_26_3_5g'] = df['cut_26_3_numeric'].apply(convert_9_to_5)
    else:
        df['cut_26_3_5g'] = np.nan
        
    return df

@st.cache_data(ttl=3600)
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
            if ncol > idx: df_cleaned[k] = df_data.iloc[:, idx]
            else: df_cleaned[k] = ""
                
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

st.markdown('<div class="brand-title">🎯 천명의선택 입시 NAVI</div>', unsafe_allow_html=True)
st.caption("천명의선택 프리미엄 입시 컨설팅 고도화 연동 솔루션")
st.markdown("---")

if df_cut.empty or df_curr.empty:
    st.error("❌ 입시 데이터 파일을 찾을 수 없습니다. (2026 수시정리.csv / 권장과목.xlsx)")
    st.stop()

st.markdown("### 📋 다음 중 선택하세요")
menu = st.radio("원하시는 메뉴를 선택하세요.", ("희망대학 컷", "진로 교과 추천"), horizontal=True, label_visibility="collapsed")
st.markdown("---")

def check_overlap(target, text):
    if pd.isna(text): return False
    text = str(text)
    for i in range(len(target) - 1):
        if target[i:i+2] in text: return True
    return False

if menu == "희망대학 컷":
    st.header("🔍 안정 · 소신 · 도전 희망대학 컷 분석")
    
    grade_system = st.radio(
        "📊 평가 체제 선택", 
        ("현 고3, N수생 (9등급제 기준)", "현 고1, 고2 (5등급제 환산 기준)"), 
        horizontal=True
    )
    is_5g = "5등급제" in grade_system
    
    if is_5g:
        st.info("💡 **[입결 데이터 안내]** 본 5등급제 환산 입결은 단순 추정이 아닌, 기존 9등급제 데이터를 누적 백분위 과학적 보간법(Interpolation)으로 산출한 **분석에 의한 예측 자료**입니다.")
    st.write("") 

    admission_types = [
        "모든전형", "교과전형", "종합전형", "논술전형", "기회전형", "농어촌전형", 
        "지역균형전형", "특성화고교전형", "실기우수자", "계약학과전형", 
        "기초생활전형", "운동선수전형", "지도자추천전형"
    ]
    selected_adm = st.radio("🎯 분석 전형 필터", options=admission_types, horizontal=True)
    st.write("")

    col1, col2 = st.columns(2)
    with col1:
        if is_5g:
            user_gpa = st.number_input("학생의 내신 평균 등급 (5등급제)을 입력하세요:", min_value=1.0, max_value=5.0, value=1.50, step=0.01)
        else:
            user_gpa = st.number_input("학생의 내신 평균 등급 (9등급제)을 입력하세요:", min_value=1.0, max_value=9.0, value=2.30, step=0.01)
        
        selected_major = st.selectbox("분석 대상 계열을 선택하세요:", list(MAJOR_MAPPING.keys()))
        
    with col2:
        all_regions = sorted([str(r).strip() for r in df_cut['지역'].dropna().unique() if str(r).strip() != "" and len(str(r).strip()) > 1])
        default_val = all_regions[:1] if all_regions else None
        selected_regions = st.multiselect("대상 지역을 선택하세요:", options=all_regions, default=default_val)
        
    if selected_major == "예체능":
        st.warning("⚠️ 해당 계열에 대한 자료는 불충분합니다.")
    elif not selected_regions:
        st.warning("⚠️ 지역을 1개 이상 선택해 주세요.")
    else:
        keywords = MAJOR_MAPPING[selected_major]
        pattern = "|".join(keywords)
        df_filtered = df_cut[df_cut['지역'].isin(selected_regions) & df_cut['학과명'].str.contains(pattern, na=False)].copy()
        
        if selected_adm != "모든전형":
            mask = df_filtered['전형명'].apply(lambda x: check_overlap(selected_adm, x))
            df_temp = df_filtered[mask].copy()
            
            if df_temp.empty:
                st.warning(f"⚠️ 선택하신 '{selected_adm}'에 해당하는 데이터가 없어 '종합전형' 기준으로 대체하여 보여드립니다.")
                mask_fallback = df_filtered['전형명'].apply(lambda x: check_overlap("종합전형", x))
                df_filtered = df_filtered[mask_fallback].copy()
            else:
                df_filtered = df_temp
        
        # 기본 70%컷 컬럼 정의
        target_col = 'cut_70_5g' if is_5g else 'cut_70'
        
        # 💡 [핵심 보완] 등급컷이 비어있지 않고 이상치(0 등)가 아닌 진짜 유효 데이터만 명확하게 추출
        valid_df = df_filtered.dropna(subset=[target_col]).copy()
        
        # 9등급 혹은 5등급 체제에 맞게 이상한 등급 범위를 가진 쓰레기 데이터 일괄 정제
        if is_5g:
            valid_df = valid_df[(valid_df[target_col] >= 1.0) & (valid_df[target_col] <= 5.0)]
        else:
            valid_df = valid_df[(valid_df[target_col] >= 1.0) & (valid_df[target_col] <= 9.0)]
            
        # 입결 오름차순 정렬 (숫자가 낮을수록 최상위 등급)
        valid_df = valid_df.sort_values(by=target_col)
        
        if valid_df.empty:
            st.error("💡 현재, 지원 가능 대학의 자료가 부족합니다.")
        else:
            st.markdown(f"#### 📊 {selected_major} 계열 통합 리포트")
            highest_row = valid_df.iloc[0] # 최상위 입결 (숫자가 가장 작은 등급)
            lowest_row = valid_df.iloc[-1]  # 최하위 입결 (숫자가 가장 큰 등급)
            
            unit = "등급 (5G)" if is_5g else "등급 (9G)"
            
            idx_col1, idx_col2 = st.columns(2)
            with idx_col1:
                st.metric("📈 계열 내 최고 입결 대학 (70%컷)", f"{highest_row['대학명']} ({highest_row['학과명']})", f"{highest_row[target_col]:.2f} {unit}", delta_color="inverse")
            with idx_col2:
                st.metric("📉 계열 내 최저 입결 대학 (70%컷)", f"{lowest_row['대학명']} ({lowest_row['학과명']})", f"{lowest_row[target_col]:.2f} {unit}")
            st.markdown("---")
            
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
                
            def format_match_df(df_match):
                if df_match.empty: return df_match
                df_sorted = df_match.sort_values(by=['대학명', '학과명', '전형명']).copy()
                
                if is_5g:
                    raw_cols = [
                        '지역', '대학명', '학과명', '전형명', 
                        '25년합격cutI기준', 'cut_70_5g', '25년합격CUTII기준', 'cut_50_5g',
                        '26년합격cut I기준', 'cut_26_1_5g', '26년합격cut II 기준', 'cut_26_2_5g', '26년합격cut III기준', 'cut_26_3_5g'
                    ]
                    rename_dict = {
                        '전형명': '전형종류', 'cut_70_5g': '2025컷(5G)', 'cut_50_5g': '2025컷2(5G)',
                        'cut_26_1_5g': '2026컷I(5G)', 'cut_26_2_5g': '2026컷II(5G)', 'cut_26_3_5g': '2026컷III(5G)'
                    }
                else:
                    raw_cols = [
                        '지역', '대학명', '학과명', '전형명', 
                        '25년합격cutI기준', '2025등급컷', '25년합격CUTII기준', '2025등급컷2',
                        '26년합격cut I기준', '2026등급컷I', '26년합격cut II 기준', '2026등급컷II', '26년합격cut III기준', '2026등급컷III'
                    ]
                    rename_dict = {'전형명': '전형종류'}

                valid_cols = [c for c in raw_cols if c in df_sorted.columns]
                df_out = df_sorted[valid_cols].copy()
                df_out = df_out.fillna('-') 
                return df_out.rename(columns=rename_dict).reset_index(drop=True)

            st.subheader("🟢 안정 지원 선 (합격 안착 확률 최상)")
            if not stable_match.empty:
                st.dataframe(format_match_df(stable_match), use_container_width=True)
            else:
                st.write("현재 지원 가능 대학의 자료가 부족합니다.")
                
            st.subheader("🟡 소신 지원 선 (적정 적합 분석권)")
            if not ambitious_match.empty:
                st.dataframe(format_match_df(ambitious_match), use_container_width=True)
            else:
                st.write("현재 지원 가능 대학의 자료가 부족합니다.")
                
            st.subheader("🔴 도전 지원 선 (상향 예측 추적권)")
            if not challenge_match.empty:
                st.dataframe(format_match_df(challenge_match), use_container_width=True)
            else:
                st.write("현재 지원 가능 대학의 자료가 부족합니다.")

elif menu == "진로 교과 추천":
    st.header("📚 모집단위별 고교 핵심 / 권장선택과목 매핑")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        region_options = sorted([str(r).strip() for r in df_curr['지역'].dropna().unique() if str(r).strip() != "" and len(str(r).strip()) > 1])
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

st.markdown("<br><br><hr>", unsafe_allow_html=True)

st.markdown(
    """
    <style>
        .footer-link {
            text-decoration: none; 
            color: #111111; 
            font-weight: bold;
            transition: color 0.2s;
        }
        .footer-link:hover {
            color: #1f77b4;
            text-decoration: underline;
        }
    </style>
    
    <div style='text-align: center; line-height: 2.0; color: #444444; font-size: 0.95rem; font-family: sans-serif;'>
        <div style='margin-bottom: 8px;'>
            대표 <a href='tel:010-3715-0994' class='footer-link'>김태영 010-3715-0994</a> &nbsp;&nbsp;·&nbsp;&nbsp; 
            소장 <a href='tel:010-3164-4029' class='footer-link'>채훈 010-3164-4029</a> &nbsp;&nbsp;·&nbsp;&nbsp; 
            <a href='http://choice1000.com' target='_blank' style='text-decoration: none; color: #1f77b4; font-weight: bold;'>choice1000.com</a>
        </div>
        <div style='color: #888888; font-size: 0.85rem;'>
            ⓒ 2026. 천명의선택 All Rights Reserved.
        </div>
    </div>
    """, 
    unsafe_allow_html=True
)
