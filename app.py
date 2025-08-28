#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Flask 웹앱 - 상품 매칭 시스템
- 파일 업로드 및 매칭 실행
- Config 설정 관리 (동의어, 색상, 사이즈 등)
"""

import os
import json
import logging
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd

from match import PerfectSolutionMatcher
import config

# Flask 앱 설정
app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# 업로드 폴더 설정
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')

@app.route('/config')
def config_page():
    """설정 관리 페이지"""
    current_config = {
        'name_synonyms': config.NAME_SYNONYM_GROUPS,
        'size_synonyms': config.SIZE_SYNONYMS,
        'color_aliases': config.COLOR_ALIASES,
        'matching_settings': config.MATCHING,
        'excel_colors': config.EXCEL_COLORS
    }
    return render_template('config.html', config=current_config)

@app.route('/api/config', methods=['GET'])
def get_config():
    """현재 설정 반환"""
    current_config = {
        'name_synonyms': config.NAME_SYNONYM_GROUPS,
        'size_synonyms': config.SIZE_SYNONYMS,
        'color_aliases': config.COLOR_ALIASES,
        'matching_settings': config.MATCHING,
        'excel_colors': config.EXCEL_COLORS
    }
    return jsonify(current_config)

@app.route('/api/config', methods=['POST'])
def update_config():
    """설정 업데이트"""
    try:
        data = request.get_json()
        
        # config.py 파일을 동적으로 업데이트
        if 'name_synonyms' in data:
            config.NAME_SYNONYM_GROUPS = data['name_synonyms']
        
        if 'size_synonyms' in data:
            config.SIZE_SYNONYMS = data['size_synonyms']
            
        if 'color_aliases' in data:
            config.COLOR_ALIASES = data['color_aliases']
            
        if 'matching_settings' in data:
            config.MATCHING = data['matching_settings']
            
        if 'excel_colors' in data:
            config.EXCEL_COLORS = data['excel_colors']
        
        # config.py 파일에 저장
        save_config_to_file()
        
        return jsonify({'success': True, 'message': '설정이 성공적으로 업데이트되었습니다.'})
    
    except Exception as e:
        logger.error(f"설정 업데이트 오류: {e}")
        return jsonify({'success': False, 'message': f'오류: {str(e)}'})

def save_config_to_file():
    """현재 설정을 config.py 파일에 저장"""
    config_content = f'''# -*- coding: utf-8 -*-
"""
설정 파일 (config.py)
- 동의어/약어 캐논화
- 색상 동의어
- 엑셀 칠하기 색상
- 매칭 임계값
- 파일 경로
- 로깅/성능 옵션
"""

# =========================
# 파일 경로 (입출력)
# =========================
DEFAULT_PATHS = {{
    "receipt": "주문_receipt.xlsx",             # 입력: 주문서
    "matched": "주문_matched.xlsx",             # 입력: 기준데이터
    "output_receipt": "주문_receipt_매칭완료.xlsx",
    "output_matched": "주문_matched_매칭완료.xlsx",
}}

# =========================
# 상품명 동의어/약어 (양방향 캐논화)
# - key(캐논라벨)로 통일됨
# - values 리스트에 있는 모든 표현은 key로 변환됨
# - 대/소문자 구분 없음, 단어 내부 포함치환
# =========================
NAME_SYNONYM_GROUPS = {repr(config.NAME_SYNONYM_GROUPS)}

# =========================
# 사이즈 동의어/약어 (토큰 캐논화)
# - key(캐논라벨)로 통일
# - 완전히 같은 토큰일 때만 매칭 (2XL → XXL 등)
# =========================
SIZE_SYNONYMS = {repr(config.SIZE_SYNONYMS)}

# =========================
# 색상 동의어 (정규화용)
# - normalize_color_ultra에 반영됨
# =========================
COLOR_ALIASES = {repr(config.COLOR_ALIASES)}

# =========================
# 엑셀 칠하기 색 (ARGB 8자리 HEX)
# - alpha(앞 2자리)는 "00"이면 불투명
# =========================
EXCEL_COLORS = {repr(config.EXCEL_COLORS)}

# =========================
# 매칭 임계값/가중치
# =========================
MATCHING = {repr(config.MATCHING)}

# =========================
# 로깅/성능
# =========================
LOGGING = {{
    "level": "INFO",  # DEBUG/INFO/WARN/ERROR
}}

PERFORMANCE = {{
    "progress_update_every": 50,  # tqdm postfix 업데이트 간격(최소)
}}
'''
    
    with open('config.py', 'w', encoding='utf-8') as f:
        f.write(config_content)

@app.route('/api/upload', methods=['POST'])
def upload_files():
    """파일 업로드 처리"""
    try:
        if 'receipt_file' not in request.files or 'matched_file' not in request.files:
            return jsonify({'success': False, 'message': '두 파일을 모두 업로드해야 합니다.'})
        
        receipt_file = request.files['receipt_file']
        matched_file = request.files['matched_file']
        
        if receipt_file.filename == '' or matched_file.filename == '':
            return jsonify({'success': False, 'message': '파일을 선택해주세요.'})
        
        if not (allowed_file(receipt_file.filename) and allowed_file(matched_file.filename)):
            return jsonify({'success': False, 'message': 'Excel 파일만 업로드 가능합니다. (.xlsx, .xls)'})
        
        # 파일 저장
        receipt_filename = secure_filename(receipt_file.filename)
        matched_filename = secure_filename(matched_file.filename)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        receipt_filename = f"{timestamp}_receipt_{receipt_filename}"
        matched_filename = f"{timestamp}_matched_{matched_filename}"
        
        receipt_path = os.path.join(UPLOAD_FOLDER, receipt_filename)
        matched_path = os.path.join(UPLOAD_FOLDER, matched_filename)
        
        receipt_file.save(receipt_path)
        matched_file.save(matched_path)
        
        return jsonify({
            'success': True, 
            'message': '파일 업로드 완료',
            'receipt_path': receipt_path,
            'matched_path': matched_path,
            'timestamp': timestamp
        })
    
    except Exception as e:
        logger.error(f"파일 업로드 오류: {e}")
        return jsonify({'success': False, 'message': f'업로드 오류: {str(e)}'})

@app.route('/api/process', methods=['POST'])
def process_matching():
    """매칭 처리 실행"""
    try:
        data = request.get_json()
        receipt_path = data.get('receipt_path')
        matched_path = data.get('matched_path')
        timestamp = data.get('timestamp')
        use_fast = data.get('use_fast', True)
        
        if not receipt_path or not matched_path:
            return jsonify({'success': False, 'message': '파일 경로가 누락되었습니다.'})
        
        # 출력 파일 경로 생성
        output_receipt = os.path.join(OUTPUT_FOLDER, f"{timestamp}_주문_receipt_매칭완료.xlsx")
        output_matched = os.path.join(OUTPUT_FOLDER, f"{timestamp}_주문_matched_매칭완료.xlsx")
        
        # 매칭 처리 실행
        matcher = PerfectSolutionMatcher()
        matched_count, match_results = matcher.process_excel_perfect_solution(
            receipt_path, matched_path, output_receipt, output_matched, use_fast=use_fast
        )
        
        # 리포트 생성
        report = matcher.generate_match_report(match_results)
        
        return jsonify({
            'success': True,
            'message': '매칭 처리 완료',
            'matched_count': matched_count,
            'report': report,
            'output_receipt': output_receipt,
            'output_matched': output_matched,
            'timestamp': timestamp
        })
    
    except Exception as e:
        logger.error(f"매칭 처리 오류: {e}")
        return jsonify({'success': False, 'message': f'매칭 오류: {str(e)}'})

@app.route('/api/download/<path:filename>')
def download_file(filename):
    """결과 파일 다운로드"""
    try:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'success': False, 'message': '파일을 찾을 수 없습니다.'})
    except Exception as e:
        logger.error(f"파일 다운로드 오류: {e}")
        return jsonify({'success': False, 'message': f'다운로드 오류: {str(e)}'})

@app.route('/api/preview', methods=['POST'])
def preview_files():
    """업로드된 파일 미리보기"""
    try:
        data = request.get_json()
        receipt_path = data.get('receipt_path')
        matched_path = data.get('matched_path')
        
        if not receipt_path or not matched_path:
            return jsonify({'success': False, 'message': '파일 경로가 누락되었습니다.'})
        
        # 파일 읽기 (첫 10행만)
        receipt_df = pd.read_excel(receipt_path, engine="openpyxl").head(10)
        matched_df = pd.read_excel(matched_path, engine="openpyxl").head(10)
        
        return jsonify({
            'success': True,
            'receipt_preview': {
                'columns': receipt_df.columns.tolist(),
                'data': receipt_df.to_dict('records'),
                'total_rows': len(pd.read_excel(receipt_path, engine="openpyxl"))
            },
            'matched_preview': {
                'columns': matched_df.columns.tolist(),
                'data': matched_df.to_dict('records'),
                'total_rows': len(pd.read_excel(matched_path, engine="openpyxl"))
            }
        })
    
    except Exception as e:
        logger.error(f"파일 미리보기 오류: {e}")
        return jsonify({'success': False, 'message': f'미리보기 오류: {str(e)}'})

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port) 