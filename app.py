import streamlit as st
from logic import AdmissionLogic
import re

st.set_page_config(page_title="Find My University", layout="wide")

@st.cache_resource
def load_engine():
    return AdmissionLogic('univ_data.db')

engine = load_engine()

st.title("🎓 대입 정시 환산점수 계산기 (표점/백분위/등급 완벽대응)")

# --- 좌측 사이드바: 학년도 선택 ---
with st.sidebar:
    st.header("📅 입시 학년도")
    selection = st.selectbox("선택", ["2026학년도 정시", "2025학년도 정시"], index=0)
    year_match = re.findall(r'\d+', selection)[0][-2:]
    target_table = f"T_{year_match}정시"
    st.divider()

# --- 메인 상단: 성적 입력 (가로 배치) ---
st.subheader("📝 수능 성적 입력")

def score_row(label, key_prefix):
    cols = st.columns([1, 2, 2, 2])
    with cols[0]: st.markdown(f"**{label}**")
    with cols[1]: std = st.number_input("표준점수", 0, 200, 130, key=f"{key_prefix}_s")
    with cols[2]: pct = st.number_input("백분위", 0, 100, 95, key=f"{key_prefix}_p")
    with cols[3]: grd = st.selectbox("등급", list(range(1, 10)), index=0, key=f"{key_prefix}_g")
    return std, pct, grd

# 1. 표점/백분위/등급 모두 필요한 과목들
kor_s, kor_p, kor_g = score_row("국어", "kor")
mat_s, mat_p, mat_g = score_row("수학", "mat")
inq1_s, inq1_p, inq1_g = score_row("탐구1", "inq1")
inq2_s, inq2_p, inq2_g = score_row("탐구2", "inq2")
for_s, for_p, for_g = score_row("제2외국어", "for")

st.divider()

# 2. 등급만 필요한 과목들
st.markdown("**[등급 필수 과목]**")
ec1, ec2 = st.columns(2)
with ec1: eng = st.selectbox("🔤 영어 등급", list(range(1, 10)), index=0)
with ec2: his = st.selectbox("🇰🇷 한국사 등급", list(range(1, 10)), index=0)

user_scores = {
    'kor': kor_s, 'kor_p': kor_p, 'kor_g': kor_g,
    'mat': mat_s, 'mat_p': mat_p, 'mat_g': mat_g,
    'inq1_s': inq1_s, 'inq1_p': inq1_p, 'inq1_g': inq1_g,
    'inq2_s': inq2_s, 'inq2_p': inq2_p, 'inq2_g': inq2_g,
    'foreign_s': for_s, 'foreign_p': for_p, 'foreign_g': for_g,
    'eng': eng, 'his': his
}

st.divider()

# --- 메인 하단: 대학 선택 및 결과 ---
if target_table not in engine.tables:
    st.error(f"⚠️ {target_table} 데이터를 로드할 수 없습니다.")
else:
    df = engine.tables[target_table]
    
    col_u, col_d = st.columns(2)
    with col_u:
        univ_col = '대학교명'
        univ_list = sorted(df[univ_col].unique())
        selected_univ = st.selectbox("🏫 대학교 선택", univ_list)
    with col_d:
        dept_col = '모집단위'
        depts = df[df[univ_col] == selected_univ][dept_col].unique()
        selected_dept = st.selectbox("📋 모집단위 선택", depts)

    if st.button("🚀 내 점수 환산하기", use_container_width=True):
        result = engine.calculate(selected_univ, selected_dept, user_scores, selection)
        if result:
            st.balloons()
            st.success(f"## {selected_univ} [{selected_dept}] 환산점수: {result:,.2f} 점")
            
            # 작년 컷 정보 표시
            target_row = df[(df[univ_col] == selected_univ) & (df[dept_col] == selected_dept)].iloc[0]
            if '25총점_70_cut' in target_row:
                st.info(f"💡 해당 학과 2025학년도 70% 합격 컷: **{target_row['25총점_70_cut']}점**")