import sqlite3
import pandas as pd
import re

class AdmissionLogic:
    def __init__(self, db_path='univ_data.db'):
        self.db_path = db_path
        self.tables = self._load_all_tables()

    def _load_all_tables(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        t_names = [row[0] for row in cursor.fetchall() if row[0].startswith('T_')]
        
        db_data = {}
        for name in t_names:
            db_data[name] = pd.read_sql(f'SELECT * FROM "{name}"', conn)
        conn.close()
        return db_data

    def get_conv_score(self, univ_name, percentile):
        """변환표준점수 테이블 연동 (VLOOKUP)"""
        df = self.tables.get('T_26정시_변환표준점수')
        if df is not None:
            col_name = '대학교명' if '대학교명' in df.columns else '대학명'
            res = df[(df[col_name] == univ_name) & (df['백분위'] == percentile)]
            if not res.empty:
                return res.iloc[0]['변환점수']
        return percentile

    def calculate(self, univ_name, dept_name, user, selection_text):
        # 테이블명 매칭 (T_26정시 등)
        year_num = re.findall(r'\d+', selection_text)[0][-2:]
        table_name = f"T_{year_num}정시"
        
        main_df = self.tables.get(table_name)
        if main_df is None: return None
        
        target = main_df[(main_df['대학교명'] == univ_name) & (main_df['모집단위'] == dept_name)]
        if target.empty: return None
        row = target.iloc[0]

        # 1. 탐구 변환점수 처리
        inq1_v = self.get_conv_score(univ_name, user['inq1_p'])
        inq2_v = self.get_conv_score(univ_name, user['inq2_p'])
        inq_sum = inq1_v + inq2_v

        # 2. 가중치 및 등급 점수 추출
        w_kor, w_mat, w_eng, w_inq = row['국어'], row['수학'], row['영어'], row['탐구']
        eng_score = row[f"영{user['eng']}"]
        his_score = row[f"한{user['his']}"]

        # 3. 대학별 특수 로직 (VBA 로직 기반)
        if "연세대" in univ_name:
            part_sum = (user['kor'] * w_kor / 200) + (user['mat'] * w_mat / 200) + (eng_score * w_eng / 100)
            inq_factor = 2 if "미래" in univ_name else 1
            part_sum += (inq_sum * w_inq * inq_factor / 200)
            divide_sum = w_kor + w_mat + w_eng + w_inq
            return (part_sum * row['만점점수'] / divide_sum) + his_score

        elif "아주대" in univ_name:
            top_df = self.tables.get('T_수능표준점수최고점')
            max_k = top_df[top_df['과목명']=='국어'].iloc[0]['최고점'] if top_df is not None else 140
            max_m = top_df[top_df['과목명']=='수학'].iloc[0]['최고점'] if top_df is not None else 140
            score = (user['kor']*w_kor/max_k*0.01) + (user['mat']*w_mat/max_m*0.01) + \
                    (eng_score*w_eng/100*0.01) + (inq_sum*w_inq/260*0.01) 
            return (score * 1000) + his_score

        else:
            # 일반 대학 합산 로직
            return (user['kor']*w_kor*0.01) + (user['mat']*w_mat*0.01) + (inq_sum*w_inq*0.01) + eng_score + his_score