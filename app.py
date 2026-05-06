import streamlit as st
import sqlite3
import pandas as pd

st.set_page_config(page_title="성적 분석 시스템", layout="centered") # 가독성을 위해 화면 중앙 정렬

st.title("🎓 성적 입력 및 분석")
st.info("각 과목의 성적을 위에서 아래 순서대로 입력해 주세요.")

# --- 1. 기본 정보 섹션 ---
with st.expander("📅 시험 정보 설정", expanded=True):
    c1, c2, c3 = st.columns(3)
    year = c1.selectbox("년도", ["2026", "2025"], index=0)
    grade = c2.selectbox("학년", ["고1", "고2", "고3"], index=0)
    month = c3.selectbox("회차/월", ["3월", "6월", "9월", "수능"], index=0)

st.markdown("---")

# --- 2. 과목별 입력 섹션 (위아래 배열) ---
# 반복되는 입력 구조를 함수로 만들어 코드를 깔끔하게 관리합니다.
def score_input(subject_name, has_std=True):
    st.markdown(f"### 📘 {subject_name}")
    cols = st.columns(3)
    if has_std:
        std = cols[0].number_input(f"{subject_name} 표준점수", value=0, key=f"{subject_name}_s")
        pct = cols[1].number_input(f"{subject_name} 백분위", value=0, key=f"{subject_name}_p")
        rank = cols[2].number_input(f"{subject_name} 등급", value=1, key=f"{subject_name}_r")
        return std, pct, rank
    else:
        rank = cols[0].number_input(f"{subject_name} 등급", value=1, key=f"{subject_name}_r")
        return None, None, rank

# 위에서 아래로 과목 나열
k_s, k_p, k_r = score_input("국어")
m_s, m_p, m_r = score_input("수학")
_, _, e_r = score_input("영어", has_std=False)
_, _, h_r = score_input("한국사", has_std=False)

# 탐구 과목 (과목명 입력 포함)
st.markdown("### 🔬 탐구 영역")
t1_col, t2_col = st.columns(2)
with t1_col:
    t1_name = st.text_input("탐구1 과목명", value="사회탐구")
    t1_s = st.number_input("탐구1 표준점수", value=0)
    t1_p = st.number_input("탐구1 백분위", value=0)
    t1_r = st.number_input("탐구1 등급", value=1)
with t2_col:
    t2_name = st.text_input("탐구2 과목명", value="과학탐구")
    t2_s = st.number_input("탐구2 표준점수 ", value=0)
    t2_p = st.number_input("탐구2 백분위 ", value=0)
    t2_r = st.number_input("탐구2 등급 ", value=1)

st.markdown("---")

# --- 3. 목표 및 검색 ---
target_col1, target_col2 = st.columns(2)
criteria = target_col1.selectbox("기준 정시", ["2025학년도 정시", "2024학년도 정시"])
target_univ = target_col2.text_input("목표 대학 그룹", value="SKY")

if st.button("🔍 분석 결과 확인", use_container_width=True):
    # 계산 및 DB 조회 로직 (여기에 VBA 로직이 들어갈 예정)
    st.success("데이터베이스 분석을 시작합니다.")
    # (생략: 이전과 동일한 DB 호출 코드)