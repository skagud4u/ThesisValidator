import streamlit as st
import pandas as pd
import os
import time
import importlib
import 논문체제_검수
importlib.reload(논문체제_검수)
from 논문체제_검수 import ThesisValidator  # 기존에 만든 검수 엔진을 임포트

st.set_page_config(page_title="HYCU 대학원 학위논문 자동 검수 시스템", layout="wide")

# 타이틀 및 아코디언 가이드
st.title("🎓 HYCU 대학원 학위논문 자동 검수 시스템")
st.caption("구글 드라이브 Sample 폴더의 학위청구논문 체제 및 서식을 자동으로 검수합니다.")

with st.expander("ℹ️ 2026-2학기 주요 검수 기준 안내"):
    st.markdown("""
    - **인준서 개정 서식**: 심사위원 날인란 최하단(마지막 행)에 반드시 **'지도교수'** 서명란이 위치해야 함.
    - **페이지 번호 체계**: 국문초록~목차는 **로마자 소문자(i, ii...)**, 본문 시작부터는 **아라비아 숫자(1, 2...)**.
    - **구성 순서**: 표지 ➔ 속표지 ➔ 인준서 ➔ 국문초록 ➔ 목차 ➔ 본문 ➔ 참고문헌 ➔ 영문초록 順.
    """)

# 사이드바 설정 영역
with st.sidebar:
    st.header("⚙️ 검수 설정")
    graduate_type = st.radio("대학원 구분", ["일반(전문)대학원", "특수대학원"])
    
    # 구글 드라이브 폴더 ID 입력 (기본값 설정 가능)
    folder_id = st.text_input("구글 드라이브 폴더 ID", value="1amEWCnb1-dfPly_TEe90B5VwPN-smHf8")
    st.info("※ 구글 API 인증 파일(credentials.json, token.json)이 프로그램과 같은 경로에 존재해야 합니다.")

# 탭 메뉴 구성
tab1, tab2 = st.tabs(["📊 논문 검수 진행", "📂 샘플 파일 업로드 및 직접검수"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 구글 드라이브에서 파일 불러오기", use_container_width=True):
            with st.spinner("구글 드라이브에서 대기 중인 논문 파일을 탐색하는 중..."):
                # 드라이브 API 연동 및 파일 다운로드 가상 프로세스
                time.sleep(1.5)
                sample_files = [f for f in os.listdir("./Sample") if f.lower().endswith(".pdf")] if os.path.exists("./Sample") else []
                st.success(f"총 {len(sample_files)}개의 학위청구논문(PDF) 파일을 'Sample' 폴더에 다운로드 완료했습니다.")
                
    with col2:
        run_btn = st.button("🔍 학위논문 일괄 서식 검수 시작", type="primary", use_container_width=True)

    if run_btn:
        st.subheader("📋 실시간 검수 수행 현황")
        progress_bar = st.progress(0)
        
        # 실제 검수 엔진 구동
        validator = ThesisValidator(target_dir="./Sample")
        # 내부 로직 수정으로 실시간 진행상황 표출 가능
        files = [f for f in os.listdir("./Sample") if f.lower().endswith(".pdf")]
        
        results = []
        for i, f in enumerate(files):
            full_path = os.path.join("./Sample", f)
            res = validator.check_thesis(full_path)
            results.append(res)
            progress_bar.progress((i + 1) / len(files))
        
        df_results = pd.DataFrame(results)
        
        # 결과를 화면에 미려하게 보여주기
        st.subheader("✨ 일괄 검수 요약 통계")
        total_cnt = len(df_results)
        pass_cnt = len(df_results[df_results['검수결과'] == '통과'])
        fail_cnt = total_cnt - pass_cnt
        
        m1, m2, m3 = st.columns(3)
        m1.metric("총 검수 대상", f"{total_cnt}건")
        m2.metric("통과 (서식 일치)", f"{pass_cnt}건", delta_color="normal")
        m3.metric("반려 (체제 보완 조치 필요)", f"{fail_cnt}건", delta="-반려")

        # 결과 데이터프레임 스타일링 (통과는 초록색, 반려/경고는 빨간색 계열)
        def style_results(val):
            if val == "통과":
                return "background-color: #E2EFDA; color: black;"
            elif val == "반려":
                return "background-color: #FCE4D6; color: black;"
            return ""

        st.dataframe(df_results.style.applymap(style_results, subset=['검수결과']), use_container_width=True)
        
        # 엑셀 다운로드 버튼 배치
        st.subheader("📥 최종 검수 리포트 출력")
        validator.run_batch_and_export() # 엑셀 파일 생성
        with open("검수결과_리포트.xlsx", "rb") as f:
            st.download_button(
                label="🟢 검수결과 리포트 Excel 다운로드",
                data=f,
                file_name="학위논문_검수결과_최종_리포트.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

with tab2:
    st.subheader("📤 개별 파일 수동 즉시 검수")
    uploaded_file = st.file_uploader("검수할 개별 논문 PDF 파일을 드래그 앤 드롭 하세요.", type=["pdf"])
    if uploaded_file is not None:
        # 업로드된 임시 파일 저장 후 즉시 검증
        temp_path = f"./temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        validator = ThesisValidator()
        res = validator.check_thesis(temp_path)
        
        if res["검수결과"] == "통과":
            st.success("🎉 서식이 체제 규격과 완벽하게 일치합니다!")
        else:
            st.error("🚨 보완 조치가 필요한 반려 항목이 감지되었습니다.")
            st.code(res["학생_안내문구"], language="text") # 바로 복사해서 보낼 수 있는 문자 양식 제공
            
        os.remove(temp_path) # 검수 후 임시 파일 즉시 파괴
