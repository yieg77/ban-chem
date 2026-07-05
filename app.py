"""
화학물질 통합 시스템
- 탭 1: 화학물질 유해성 정보 수집 (KOSHA MSDS) - 다중 파일 처리
- 탭 2: 화학물질 통계 정보공개 검색 (ICIS)
"""
import streamlit as st
from createForm import main_ui
from makeResult import makeResult_multi_ui

st.set_page_config(
    page_title='화학물질 통합 시스템',
    layout='wide',
    initial_sidebar_state='collapsed'
)

st.markdown("""
<style>
/* 탭 버튼 - 너비 및 굵기 */
button[data-baseweb="tab"] {
    background-color: #ffffff !important;
    min-width: 200px !important;
    padding: 10px 32px !important;
}
button[data-baseweb="tab"] p {
    font-size: 15px !important;
    font-weight: 700 !important;
    color: #555 !important;
}
button[data-baseweb="tab"][aria-selected="true"] p {
    color: #1a73e8 !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    border-bottom-color: #1a73e8 !important;
    background-color: #ffffff !important;
}
button[data-baseweb="tab"]:hover {
    background-color: #f5f5f5 !important;
}
div[role="tablist"] {
    background-color: #ffffff !important;
}
</style>
""", unsafe_allow_html=True)


# 상단 공간을 만들기 위해 컬럼 활용
col_tab, col_ver = st.columns([0.85, 0.15])

with col_ver:
    # 탭 바로 윗줄 오른쪽 끝에 위치하도록 정렬
    st.markdown(
        """<div style="text-align: right; color: #999; font-size: 15px; margin-top: 10px;">
        v2.260705
        </div>""", 
        unsafe_allow_html=True
    )

tab1, tab2 = st.tabs(["유해성 정보 수집", "업체별 입력폼 생성"])

with tab1:
    st.markdown("## 화학물질 유해성 정보 수집")
    st.write("")
    makeResult_multi_ui()

with tab2:
    st.markdown("## 업체별 입력폼 생성")
    st.write("")
    main_ui(tab_mode=True)
