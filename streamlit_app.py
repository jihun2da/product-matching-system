import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
import tempfile
import zipfile
from io import BytesIO

# 기존 매칭 엔진 import
from match import PerfectSolutionMatcher
import config

# 페이지 설정
st.set_page_config(
    page_title="상품 매칭 시스템",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 세션 상태 초기화
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = None
if 'match_results' not in st.session_state:
    st.session_state.match_results = None
if 'config_data' not in st.session_state:
    st.session_state.config_data = {
        'name_synonyms': config.NAME_SYNONYM_GROUPS,
        'size_synonyms': config.SIZE_SYNONYMS,
        'color_aliases': config.COLOR_ALIASES,
        'matching_settings': config.MATCHING,
        'excel_colors': config.EXCEL_COLORS
    }

def main():
    st.title("🔍 상품 매칭 시스템")
    st.markdown("---")
    
    # 사이드바 메뉴
    menu = st.sidebar.selectbox(
        "메뉴 선택",
        ["🚀 매칭 실행", "⚙️ 설정 관리", "📊 결과 분석"]
    )
    
    if menu == "🚀 매칭 실행":
        matching_page()
    elif menu == "⚙️ 설정 관리":
        config_page()
    elif menu == "📊 결과 분석":
        results_page()

def matching_page():
    st.header("🚀 매칭 실행")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📁 파일 업로드")
        
        receipt_file = st.file_uploader(
            "주문서 파일 (Receipt)",
            type=['xlsx', 'xls'],
            help="매칭할 주문 데이터가 포함된 Excel 파일"
        )
        
        matched_file = st.file_uploader(
            "기준 데이터 파일 (Matched)",
            type=['xlsx', 'xls'],
            help="매칭 기준이 되는 상품 데이터 Excel 파일"
        )
        
        use_fast = st.checkbox("고속 매칭 사용 (권장)", value=True)
        
        if receipt_file and matched_file:
            st.success("✅ 파일 업로드 완료!")
            
            # 파일 미리보기 버튼
            if st.button("👀 파일 미리보기"):
                preview_files(receipt_file, matched_file)
    
    with col2:
        st.subheader("🎯 매칭 실행")
        
        if receipt_file and matched_file:
            if st.button("▶️ 매칭 시작", type="primary"):
                run_matching(receipt_file, matched_file, use_fast)
        else:
            st.info("먼저 두 파일을 모두 업로드해주세요.")
        
        # 결과 표시
        if st.session_state.match_results:
            display_results()

def preview_files(receipt_file, matched_file):
    st.subheader("📋 파일 미리보기")
    
    tab1, tab2 = st.tabs(["주문서 파일", "기준 데이터 파일"])
    
    with tab1:
        try:
            df_receipt = pd.read_excel(receipt_file)
            st.write(f"**총 행 수:** {len(df_receipt)}행")
            st.dataframe(df_receipt.head(10), use_container_width=True)
        except Exception as e:
            st.error(f"파일 읽기 오류: {str(e)}")
    
    with tab2:
        try:
            df_matched = pd.read_excel(matched_file)
            st.write(f"**총 행 수:** {len(df_matched)}행")
            st.dataframe(df_matched.head(10), use_container_width=True)
        except Exception as e:
            st.error(f"파일 읽기 오류: {str(e)}")

def run_matching(receipt_file, matched_file, use_fast):
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        status_text.text("파일 준비 중...")
        progress_bar.progress(10)
        
        # 임시 파일로 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_receipt:
            tmp_receipt.write(receipt_file.getvalue())
            receipt_path = tmp_receipt.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_matched:
            tmp_matched.write(matched_file.getvalue())
            matched_path = tmp_matched.name
        
        progress_bar.progress(30)
        status_text.text("매칭 엔진 초기화 중...")
        
        # 출력 파일 경로
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_receipt = f"temp_receipt_{timestamp}.xlsx"
        output_matched = f"temp_matched_{timestamp}.xlsx"
        
        progress_bar.progress(50)
        status_text.text("매칭 처리 중...")
        
        # 매칭 실행
        matcher = PerfectSolutionMatcher()
        matched_count, match_results = matcher.process_excel_perfect_solution(
            receipt_path, matched_path, output_receipt, output_matched, use_fast=use_fast
        )
        
        progress_bar.progress(80)
        status_text.text("결과 처리 중...")
        
        # 리포트 생성
        report = matcher.generate_match_report(match_results)
        
        # 결과 저장
        st.session_state.match_results = {
            'matched_count': matched_count,
            'report': report,
            'output_receipt': output_receipt,
            'output_matched': output_matched,
            'timestamp': timestamp
        }
        
        progress_bar.progress(100)
        status_text.text("완료!")
        
        # 임시 파일 정리
        try:
            os.unlink(receipt_path)
            os.unlink(matched_path)
        except:
            pass  # 파일 삭제 실패해도 무시
        
        st.success("🎉 매칭 완료!")
        
    except Exception as e:
        st.error(f"매칭 중 오류 발생: {str(e)}")
        st.error(f"상세 오류: {type(e).__name__}")
    finally:
        progress_bar.empty()
        status_text.empty()

def display_results():
    results = st.session_state.match_results
    
    st.subheader("📊 매칭 결과")
    
    # 통계 표시
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("총 매칭 건수", f"{results['matched_count']:,}건")
    
    with col2:
        st.metric("평균 신뢰도", f"{results['report']['average_confidence']:.1f}%")
    
    with col3:
        st.metric("고신뢰도", f"{results['report']['match_distribution']['high_confidence']}건")
    
    with col4:
        st.metric("중신뢰도", f"{results['report']['match_distribution']['medium_confidence']}건")
    
    # 다운로드 버튼
    col1, col2 = st.columns(2)
    
    with col1:
        if os.path.exists(results['output_receipt']):
            with open(results['output_receipt'], 'rb') as f:
                st.download_button(
                    "📥 주문서 결과 다운로드",
                    f.read(),
                    file_name=f"주문서_매칭결과_{results['timestamp']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    with col2:
        if os.path.exists(results['output_matched']):
            with open(results['output_matched'], 'rb') as f:
                st.download_button(
                    "📥 기준데이터 결과 다운로드",
                    f.read(),
                    file_name=f"기준데이터_매칭결과_{results['timestamp']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

def config_page():
    st.header("⚙️ 설정 관리")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏷️ 상품명 동의어", 
        "📏 사이즈 동의어", 
        "🎨 색상 동의어", 
        "⚖️ 매칭 설정", 
        "📊 엑셀 색상"
    ])
    
    with tab1:
        manage_name_synonyms()
    
    with tab2:
        manage_size_synonyms()
    
    with tab3:
        manage_color_aliases()
    
    with tab4:
        manage_matching_settings()
    
    with tab5:
        manage_excel_colors()
    
    # 설정 저장 버튼
    st.markdown("---")
    col1, col2 = st.columns([1, 4])
    
    with col1:
        if st.button("💾 설정 저장", type="primary"):
            save_config()
            st.success("설정이 저장되었습니다!")
    
    with col2:
        if st.button("🔄 초기화"):
            reset_config()
            st.info("설정이 초기화되었습니다.")

def manage_name_synonyms():
    st.subheader("🏷️ 상품명 동의어 관리")
    st.write("상품명에서 사용되는 동의어와 약어를 관리합니다.")
    
    # 기존 동의어 표시 및 편집
    name_synonyms = st.session_state.config_data['name_synonyms'].copy()
    
    for i, (key, values) in enumerate(list(name_synonyms.items())):
        col1, col2, col3 = st.columns([2, 4, 1])
        
        with col1:
            new_key = st.text_input(f"표준 용어", value=key, key=f"name_key_{i}")
        
        with col2:
            values_str = ", ".join(values) if isinstance(values, list) else str(values)
            new_values = st.text_input(f"동의어 (쉼표로 구분)", value=values_str, key=f"name_values_{i}")
        
        with col3:
            if st.button("🗑️", key=f"name_del_{i}", help="삭제"):
                if key in st.session_state.config_data['name_synonyms']:
                    del st.session_state.config_data['name_synonyms'][key]
                    st.rerun()
        
        # 변경사항 반영
        if new_key != key:
            if key in st.session_state.config_data['name_synonyms']:
                del st.session_state.config_data['name_synonyms'][key]
            st.session_state.config_data['name_synonyms'][new_key] = [v.strip() for v in new_values.split(',') if v.strip()]
        else:
            st.session_state.config_data['name_synonyms'][key] = [v.strip() for v in new_values.split(',') if v.strip()]
    
    # 새 동의어 그룹 추가
    st.markdown("**새 동의어 그룹 추가**")
    col1, col2, col3 = st.columns([2, 4, 1])
    
    with col1:
        new_key = st.text_input("새 표준 용어", key="new_name_key")
    
    with col2:
        new_values = st.text_input("새 동의어들 (쉼표로 구분)", key="new_name_values")
    
    with col3:
        if st.button("➕ 추가", key="add_name_synonym"):
            if new_key and new_values:
                st.session_state.config_data['name_synonyms'][new_key] = [v.strip() for v in new_values.split(',') if v.strip()]
                st.rerun()

def manage_size_synonyms():
    st.subheader("📏 사이즈 동의어 관리")
    st.write("사이즈 표기의 동의어를 관리합니다. (예: 2XL → XXL)")
    
    # 기존 사이즈 동의어 관리 (name_synonyms와 유사한 로직)
    size_synonyms = st.session_state.config_data['size_synonyms'].copy()
    
    for i, (key, values) in enumerate(list(size_synonyms.items())):
        col1, col2, col3 = st.columns([2, 4, 1])
        
        with col1:
            new_key = st.text_input(f"표준 사이즈", value=key, key=f"size_key_{i}")
        
        with col2:
            values_str = ", ".join(values) if isinstance(values, list) else str(values)
            new_values = st.text_input(f"동의어 사이즈", value=values_str, key=f"size_values_{i}")
        
        with col3:
            if st.button("🗑️", key=f"size_del_{i}", help="삭제"):
                if key in st.session_state.config_data['size_synonyms']:
                    del st.session_state.config_data['size_synonyms'][key]
                    st.rerun()
        
        # 변경사항 반영
        if new_key != key:
            if key in st.session_state.config_data['size_synonyms']:
                del st.session_state.config_data['size_synonyms'][key]
            st.session_state.config_data['size_synonyms'][new_key] = [v.strip() for v in new_values.split(',') if v.strip()]
        else:
            st.session_state.config_data['size_synonyms'][key] = [v.strip() for v in new_values.split(',') if v.strip()]

def manage_color_aliases():
    st.subheader("🎨 색상 동의어 관리")
    st.write("색상 표기의 정규화 규칙을 관리합니다.")
    
    # 색상 동의어 관리
    color_aliases = st.session_state.config_data['color_aliases'].copy()
    
    for i, (key, value) in enumerate(list(color_aliases.items())):
        col1, col2, col3 = st.columns([3, 3, 1])
        
        with col1:
            new_key = st.text_input(f"원본 색상", value=key, key=f"color_key_{i}")
        
        with col2:
            new_value = st.text_input(f"정규화 결과", value=value, key=f"color_value_{i}")
        
        with col3:
            if st.button("🗑️", key=f"color_del_{i}", help="삭제"):
                if key in st.session_state.config_data['color_aliases']:
                    del st.session_state.config_data['color_aliases'][key]
                    st.rerun()
        
        # 변경사항 반영
        if new_key != key:
            if key in st.session_state.config_data['color_aliases']:
                del st.session_state.config_data['color_aliases'][key]
            st.session_state.config_data['color_aliases'][new_key] = new_value
        else:
            st.session_state.config_data['color_aliases'][key] = new_value

def manage_matching_settings():
    st.subheader("⚖️ 매칭 설정")
    
    # 임계값 설정
    cutoff = st.slider(
        "상품명 매칭 임계값",
        min_value=0, max_value=100,
        value=st.session_state.config_data['matching_settings'].get('name_score_cutoff', 70),
        help="상품명 퍼지 매칭의 최소 점수"
    )
    
    st.session_state.config_data['matching_settings']['name_score_cutoff'] = cutoff
    
    # 가중치 설정
    st.write("**매칭 가중치** (합계가 1.0이 되어야 합니다)")
    
    weights = st.session_state.config_data['matching_settings'].get('weights', {})
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        brand_weight = st.number_input("브랜드", min_value=0.0, max_value=1.0, value=weights.get('brand', 0.25), step=0.05)
    
    with col2:
        name_weight = st.number_input("상품명", min_value=0.0, max_value=1.0, value=weights.get('name', 0.35), step=0.05)
    
    with col3:
        color_weight = st.number_input("색상", min_value=0.0, max_value=1.0, value=weights.get('color', 0.25), step=0.05)
    
    with col4:
        size_weight = st.number_input("사이즈", min_value=0.0, max_value=1.0, value=weights.get('size', 0.15), step=0.05)
    
    # 가중치 합계 확인
    total_weight = brand_weight + name_weight + color_weight + size_weight
    
    if abs(total_weight - 1.0) > 0.01:
        st.warning(f"⚠️ 가중치의 합이 1.0이 아닙니다. 현재 합: {total_weight:.3f}")
    else:
        st.success("✅ 가중치 설정이 올바릅니다.")
    
    # 가중치 저장
    st.session_state.config_data['matching_settings']['weights'] = {
        'brand': brand_weight,
        'name': name_weight,
        'color': color_weight,
        'size': size_weight
    }

def manage_excel_colors():
    st.subheader("📊 엑셀 색상 설정")
    st.write("매칭 결과 파일에서 사용할 색상을 설정합니다. (ARGB 8자리 HEX)")
    
    excel_colors = st.session_state.config_data['excel_colors'].copy()
    
    for i, (key, value) in enumerate(list(excel_colors.items())):
        col1, col2, col3 = st.columns([2, 3, 1])
        
        with col1:
            new_key = st.text_input(f"색상 이름", value=key, key=f"excel_key_{i}")
        
        with col2:
            new_value = st.text_input(f"ARGB 코드", value=value, key=f"excel_value_{i}", max_chars=8)
        
        with col3:
            if st.button("🗑️", key=f"excel_del_{i}", help="삭제"):
                if key in st.session_state.config_data['excel_colors']:
                    del st.session_state.config_data['excel_colors'][key]
                    st.rerun()
        
        # 변경사항 반영
        if new_key != key:
            if key in st.session_state.config_data['excel_colors']:
                del st.session_state.config_data['excel_colors'][key]
            st.session_state.config_data['excel_colors'][new_key] = new_value
        else:
            st.session_state.config_data['excel_colors'][key] = new_value

def save_config():
    # config.py 파일 업데이트
    config.NAME_SYNONYM_GROUPS = st.session_state.config_data['name_synonyms']
    config.SIZE_SYNONYMS = st.session_state.config_data['size_synonyms']
    config.COLOR_ALIASES = st.session_state.config_data['color_aliases']
    config.MATCHING = st.session_state.config_data['matching_settings']
    config.EXCEL_COLORS = st.session_state.config_data['excel_colors']

def reset_config():
    st.session_state.config_data = {
        'name_synonyms': {"맨투맨": ["mtm"], "자켓": ["jk", "재킷", "쟈켓"]},
        'size_synonyms': {"xxl": ["2xl"]},
        'color_aliases': {"gray": "회색", "black": "검정"},
        'matching_settings': {"name_score_cutoff": 70, "weights": {"brand": 0.25, "name": 0.35, "color": 0.25, "size": 0.15}},
        'excel_colors': {"cyan": "0000FFFF", "yellow": "00FFFF00", "red": "00FF0000"}
    }

def results_page():
    st.header("📊 결과 분석")
    
    if st.session_state.match_results:
        results = st.session_state.match_results
        
        # 상세 통계
        st.subheader("📈 매칭 통계")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric("총 매칭 건수", f"{results['matched_count']:,}건")
            st.metric("평균 신뢰도", f"{results['report']['average_confidence']:.2f}%")
        
        with col2:
            # 신뢰도 분포 차트
            dist = results['report']['match_distribution']
            chart_data = pd.DataFrame({
                '신뢰도': ['고신뢰도\n(90%+)', '중신뢰도\n(70-90%)', '저신뢰도\n(<70%)'],
                '건수': [dist['high_confidence'], dist['medium_confidence'], dist['low_confidence']]
            })
            st.bar_chart(chart_data.set_index('신뢰도'))
    
    else:
        st.info("매칭을 실행한 후 결과를 확인할 수 있습니다.")

if __name__ == "__main__":
    main() 