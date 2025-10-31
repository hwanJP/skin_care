"""
app_recipe.py - 제형 레시피 OCR (기존 app.py에서 최소 수정)
"""

import streamlit as st
import pandas as pd
import os
import sys
import tempfile
import uuid
from pathlib import Path
from datetime import datetime
import io
import fitz
import copy

# ========================================
# 🔧 수정 1: import 변경
# ========================================
# ❌ 기존: from backend import process_pdf_page, ExcelIncrementalSaver
# ✅ 신규: from backend_recipe import process_recipe_page, RecipeExcelSaver

current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

from backend import PDFProcessor, logger  # ✅ PDFProcessor는 재사용
from backend_recipe import (              # 🆕 제형 레시피 전용
    process_recipe_page,
    RecipeExcelSaver
)

# ========================================
# ✅ 동일: 페이지 설정
# ========================================
st.set_page_config(
    page_title="한국콜마 실험 처방 READER",  # 🔧 제목만 변경
    layout="wide",
    initial_sidebar_state="collapsed"
)

MAX_PDF_PAGES = 50       # 최대 페이지 수
MAX_FILE_SIZE_MB = 40    # 최대 파일 크기 (MB)
# ========================================
# ✅ 동일: 세션 상태 초기화 (99% 동일)
# ========================================
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

if "ocr_data_frames" not in st.session_state:
    st.session_state.ocr_data_frames = {}

if "current_page" not in st.session_state:
    st.session_state.current_page = 1

if "saved_pages" not in st.session_state:
    st.session_state.saved_pages = set()

if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None

if "current_file_bytes" not in st.session_state:
    st.session_state.current_file_bytes = None

if "confirm_reset" not in st.session_state:
    st.session_state.confirm_reset = False

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = {}

if "excel_saver" not in st.session_state:
    temp_dir = tempfile.gettempdir()
    excel_path = os.path.join(temp_dir, f"제형레시피_{st.session_state.session_id}.xlsx")
    st.session_state.excel_saver = RecipeExcelSaver(excel_path)
    st.session_state.excel_path = excel_path
# ============================================
# 🆕 저장 함수 (공통)
# ============================================
def save_current_page():
    """현재 페이지 데이터 Excel 저장"""
    key = (st.session_state.current_file_name, st.session_state.current_page)
    
    if key not in st.session_state.ocr_data_frames:
        return True
    
    bundle = st.session_state.ocr_data_frames[key]
    data = bundle.get('data', [])
    
    if not data:
        return True
    
    # ✅ 임시 저장소에서 edited_df 가져오기
    temp_df = st.session_state.get(f'_temp_edited_df_{key}')
    
    if temp_df is not None and len(temp_df) > 0:
        # 원본 _corrections 백업
        original_corrections = {
            ing.get('Code', f'idx_{i}'): ing.get('_corrections', {})
            for i, ing in enumerate(data)
        }
        
        # 메모 저장
        memo_content = temp_df.iloc[0].to_dict()
        if '_is_separator' in memo_content:
            del memo_content['_is_separator']
        bundle['memo'] = memo_content
        
        # 데이터 저장
        if len(temp_df) > 1:
            edited_data = []
            
            for _, row in temp_df.iloc[1:].iterrows():
                ingredient = row.to_dict()
                
                if ingredient.get('_is_separator', False):
                    continue
                
                if '_is_separator' in ingredient:
                    del ingredient['_is_separator']
                
                code = ingredient.get('Code', '')
                if code in original_corrections:
                    ingredient['_corrections'] = original_corrections[code]
                
                edited_data.append(ingredient)
            
            bundle['data'] = edited_data
    
    # Excel 저장
    metadata_with_memo = bundle['metadata'].copy()
    metadata_with_memo['memo'] = bundle.get('memo', {})
    
    if 'saved_sheet_name' in bundle:
        metadata_with_memo['saved_sheet_name'] = bundle['saved_sheet_name']
    
    with st.spinner('저장 중...'):
        result = st.session_state.excel_saver.add_recipe_data(
            data=bundle['data'],
            metadata=metadata_with_memo,
            experiment_cols=bundle['experiment_columns']
        )
    
    if result['success']:
        st.session_state.ocr_data_frames[key]['saved_sheet_name'] = result['sheet_name']
        st.session_state.saved_pages.add(key)
        return True
    else:
        st.error('저장 실패. 다시 시도해주세요.')
        return False
# ========================================
# ✅ 동일: CSS 스타일
# ========================================
st.markdown("""
<style>
    .compact-header {
        background: linear-gradient(90deg, #0066cc 0%, #0099ff 100%) !important;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        color: white !important;  /* !important 추가 */
        margin-bottom: 1rem;
    }
    .status-bar {
        background-color: #f0f2f6 !important;
        padding: 0.5rem;
        border-radius: 5px;
        margin: 0.5rem 0;
        color: #000000 !important;  /* 텍스트 색 명시 */
    }
    
    /* 다크 모드 대응 */
    [data-testid="stAppViewContainer"] .compact-header {
        background: linear-gradient(90deg, #0066cc 0%, #0099ff 100%) !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ========================================
# 🔧 수정: 헤더 (제목만 변경)
# ========================================
st.markdown("""
<div class="compact-header">
    <h1>한국콜마 실험 처방 READER</h1>
    <p>Azure Document Intelligence 기반 PDF to Excel 자동 변환</p>
</div>
""", unsafe_allow_html=True)

# ========================================
# ✅ 동일: 파일 업로드 영역 (100% 동일)
# ========================================
header_col1, header_col2 = st.columns([4, 1])

with header_col1:
    has_work = len(st.session_state.ocr_data_frames) > 0
    
    if not has_work:
        uploaded_file = st.file_uploader(
            "PDF 파일 선택",
            type=["pdf"],
            accept_multiple_files=False,
            label_visibility="collapsed",
            key="file_uploader"
        )
        
        if uploaded_file:
            file_id = f"{uploaded_file.name}_{len(uploaded_file.getvalue())}"
            
            if st.session_state.current_file_name != uploaded_file.name:
                if file_id not in st.session_state.processed_files:
                    with st.spinner("🔐 파일 확인 중..."):
                        original_bytes = uploaded_file.getvalue()
                        
                        # ============================================
                        # 🆕 1. 파일 크기 체크
                        # ============================================
                        file_size_mb = len(original_bytes) / (1024 * 1024)
                        
                        if file_size_mb > MAX_FILE_SIZE_MB:
                            st.error(f"파일 크기가 제한을 초과했습니다. ({file_size_mb:.1f}MB / {MAX_FILE_SIZE_MB}MB)")
                            st.info(f"현재 파일 크기: {file_size_mb:.1f}MB")
                            st.stop()
                        
                        # DRM 처리
                        drm_success, processed_bytes, drm_message = PDFProcessor.process_drm_if_needed(original_bytes)
                        
                        if not drm_success:
                            st.error(f"파일 처리 실패: {drm_message}")
                            logger.error(f"DRM 처리 실패: {drm_message}")
                            st.stop()
                        
                        # ============================================
                        # 🆕 2. 페이지 수 체크
                        # ============================================
                        try:
                            doc = fitz.open(stream=processed_bytes, filetype="pdf")
                            page_count = doc.page_count
                            doc.close()
                            
                            # 페이지 수 제한 체크
                            if page_count > MAX_PDF_PAGES:
                                st.error(f"PDF 페이지 수가 제한을 초과했습니다. (최대 {MAX_PDF_PAGES}페이지)")
                                st.info(f"현재 PDF: {page_count}페이지")
                                st.info(f"ℹPDF를 {MAX_PDF_PAGES}페이지 이하로 분할하거나, 필요한 페이지만 추출해주세요.")
                                
                                st.stop()
                            
                        except Exception as e:
                            st.error(f"❌ PDF 열기 실패: {e}")
                            st.stop()
                        
                        st.session_state.processed_files[file_id] = {
                            'bytes': processed_bytes,
                            'message': drm_message,
                            'name': uploaded_file.name,
                            'page_count': page_count
                        }
                        
                        if "DRM 처리 완료" in drm_message or "DRM 해제" in drm_message:
                            st.success(f"{drm_message} | 총 {page_count} 페이지")
                        else:
                            st.success(f"파일 로드 완료 | 총 {page_count} 페이지")
                
                processed_file_info = st.session_state.processed_files[file_id]
                st.session_state.current_file_name = uploaded_file.name
                st.session_state.current_file_bytes = processed_file_info['bytes']
                st.session_state.current_file_id = file_id
                st.session_state.current_page = 1
                st.rerun()
# ========================================
# 🆕 새로 시작하기 버튼 (2단계 확인)
# ========================================
with header_col2:
    if has_work:
        # 1단계: 일반 버튼
        if not st.session_state.get('reset_confirm', False):
            if st.button("🔄 새로 시작하기", use_container_width=True, type="secondary"):
                st.session_state.reset_confirm = True
                st.rerun()
        
        # 2단계: 경고 + 확인 버튼
        else:
            col1, col2 = st.columns(2)
            with col1:
                if st.button("취소", use_container_width=True, type="secondary"):
                    st.session_state.reset_confirm = False
                    st.rerun()
            with col2:
                if st.button("모두 삭제", use_container_width=True, type="primary"):
                    # Excel 파일 삭제
                    if os.path.exists(st.session_state.excel_path):
                        os.remove(st.session_state.excel_path)
                    
                    # 초기화
                    st.session_state.ocr_data_frames = {}
                    st.session_state.saved_pages = set()
                    st.session_state.current_page = 1
                    st.session_state.current_file_name = None
                    st.session_state.current_file_bytes = None
                    st.session_state.current_file_id = None
                    st.session_state.processed_files = {}
                    st.session_state.reset_confirm = False
                    
                    # 새 Excel 생성
                    new_session_id = str(uuid.uuid4())
                    excel_path = os.path.join(tempfile.gettempdir(), f"제형레시피_{new_session_id}.xlsx")
                    st.session_state.excel_saver = RecipeExcelSaver(excel_path)
                    st.session_state.excel_path = excel_path
                    st.session_state.session_id = new_session_id
                    
                    st.success("초기화 완료")
                    st.rerun()
        
        # 경고 메시지 (2단계일 때)
        if st.session_state.get('reset_confirm', False):
            st.warning("모든 작업(PDF, OCR 결과, Excel)이 영구 삭제됩니다!")
            
# ========================================
# ✅ 동일: 현재 파일 설정
# ========================================
current_file = None
page_count = 0

if st.session_state.get('current_file_bytes'):
    current_file = type('obj', (object,), {
        'name': st.session_state.current_file_name,
        'getvalue': lambda self: st.session_state.current_file_bytes
    })()
    
    page_count = PDFProcessor.extract_page_count(st.session_state.current_file_bytes)
    
    if st.session_state.current_page > page_count:
        st.session_state.current_page = page_count
    if st.session_state.current_page < 1:
        st.session_state.current_page = 1

# ========================================
# 메인 컨텐츠
# ========================================
if current_file:
    # ✅ 동일: 상단 액션바
    action_col1, action_col2, action_col3, action_col4, action_col5 = st.columns([2, 2, 2, 1, 2])
    
    # ============================================
    # 버튼 1: OCR 시작 (col1) - 상태 관리
    # ============================================
    with action_col1:
        key = (current_file.name, st.session_state.current_page)
        ocr_completed = key in st.session_state.ocr_data_frames
        has_data = len(st.session_state.ocr_data_frames.get(key, {}).get('data', [])) > 0
        
        if ocr_completed and has_data:
            button_label = "OCR 완료"
            disabled = True
        elif ocr_completed and not has_data:
            button_label = "OCR 재시도"
            disabled = False
        else:
            button_label = "OCR 시작"
            disabled = False
        
        if st.button(button_label, type="primary", use_container_width=True, disabled=disabled):
            with st.spinner(f"페이지 {st.session_state.current_page} 처리 중..."):
                result = process_recipe_page(
                    current_file.getvalue(), 
                    st.session_state.current_page - 1
                )
                
                if result['success']:
                    st.session_state.ocr_data_frames[key] = {
                        "data": result['data'],
                        "metadata": result['metadata'],
                        "experiment_columns": result['experiment_columns']
                    }
                    st.success(f"{len(result['data'])}개 원료 추출 완료")
                    st.rerun()  # ✅ 필수 - OCR 결과를 UI에 반영
                else:
                    st.session_state.ocr_data_frames[key] = {
                        "data": [],
                        "metadata": {},
                        "experiment_columns": [],
                        "_error": result['message']
                    }
                    st.error(f"OCR 실패: {result['message']}")
                    st.info("'OCR 재시도' 버튼을 클릭하여 다시 시도하세요")
                    st.rerun()  # ✅ 필수 - 버튼 상태 변경 반영 (재시도로 변경)
    
    # ============================================
    # 버튼 2: ◀ 이전 (col2)
    # ============================================
    with action_col2:
        if st.button("◀ 이전", use_container_width=True, 
                    disabled=(st.session_state.current_page <= 1)):
            
            if save_current_page():  # ✅ 저장 성공 시에만 이동
                st.session_state.current_page -= 1
                st.rerun()  # ✅ 필수 - 페이지 변경 반영
    
    # ============================================
    # 버튼 3: ▶ 다음 (col3)
    # ============================================
    with action_col3:
        # OCR 상태 확인
        key = (current_file.name, st.session_state.current_page)
        ocr_completed = key in st.session_state.ocr_data_frames
        has_data = len(st.session_state.ocr_data_frames.get(key, {}).get('data', [])) > 0
        
        # 마지막 페이지 확인
        is_last_page = (st.session_state.current_page >= page_count)
        
        # 비활성화 조건
        if is_last_page:
            disabled = False  # 마지막 페이지는 항상 활성화 (저장 전용)
        else:
            disabled = not (ocr_completed and has_data)  # OCR 완료되어야 활성화
        
        if st.button("▶ 다음", type="primary", use_container_width=True, disabled=disabled):
            if save_current_page():
                if is_last_page:
                    st.success("✅ 마지막 페이지 저장 완료!")
                    # ❌ rerun 제거 - 저장만 하고 현재 페이지 유지
                else:
                    st.session_state.current_page += 1
                    st.rerun()  # ✅ 필수 - 페이지 변경 반영
    
    # ============================================
    # 버튼 4: 💾 N/M (col4) - 저장 현황
    # ============================================
    with action_col4:
        saved_count = len(st.session_state.saved_pages)
        st.button(f"{saved_count}/{page_count}", 
                  use_container_width=True, disabled=True)
    
    # ============================================
    # 버튼 5: 📥 Excel 다운로드 (col5)
    # ============================================
    with action_col5:
        if len(st.session_state.saved_pages) > 0 and os.path.exists(st.session_state.excel_path):
            excel_bytes = st.session_state.excel_saver.get_excel_bytes()
            
            if excel_bytes:
                stats = st.session_state.excel_saver.get_statistics()
                file_size_mb = stats.get('file_size_mb', 0)
                
                st.download_button(
                    label=f"Excel 다운로드 ({file_size_mb:.1f}MB)",
                    data=excel_bytes,
                    file_name=f"제형레시피_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.button("Excel 다운로드", use_container_width=True, disabled=True)
    
    # ========================================
    # ✅ 상태 표시줄
    # ========================================
    key = (current_file.name, st.session_state.current_page)
    processed_pages = len(st.session_state.ocr_data_frames)
    
    status_html = f"""
    <div class="status-bar">
        <strong>페이지:</strong> {st.session_state.current_page}/{page_count} | 
        <strong>처리 완료:</strong> {processed_pages}/{page_count}
    </div>
    """
    st.markdown(status_html, unsafe_allow_html=True)
    
    # ========================================
    # 메인 컨텐츠 영역 (2단 레이아웃)
    # ========================================
    left_col, right_col = st.columns([4, 6])

    # ✅ 동일: 좌측 PDF 미리보기
    with left_col:
        st.markdown("### PDF 미리보기")
        
        # ✅ PDFProcessor 메서드 사용 (과거 완성형 방식)
        img_bytes = PDFProcessor.render_page_image(
            current_file.getvalue(), 
            st.session_state.current_page - 1, 
            zoom=2.5  # ✅ 높은 zoom으로 OCR 결과 확인에 유리
        )
        
        if img_bytes:
            st.image(
                img_bytes,
                caption=f"{current_file.name} - 페이지 {st.session_state.current_page}/{page_count}",
                use_column_width=True
            )
        else:
            st.error("이미지 렌더링 실패")

    # ============================================
    # 우측: OCR 결과 (자동 OCR 포함)
    # ============================================
    with right_col:
        st.markdown("### OCR 결과")
        
        key = (current_file.name, st.session_state.current_page)
        
        # ========================================
        # 🆕 자동 OCR 로직 (2페이지 이상, OCR 안 됨)
        # ========================================
        if key not in st.session_state.ocr_data_frames and st.session_state.current_page > 1:
            with st.spinner("페이지 분석 중... (약 5초 소요)"):
                result = process_recipe_page(
                    current_file.getvalue(), 
                    st.session_state.current_page - 1
                )
                
                if result['success']:
                    st.session_state.ocr_data_frames[key] = {
                        "data": result['data'],
                        "metadata": result['metadata'],
                        "experiment_columns": result['experiment_columns']
                    }
                    st.success(f"자동 OCR 완료: {len(result['data'])}개 원료")
                    st.rerun()  # ✅ 추가 - 결과 즉시 반영
                else:
                    st.session_state.ocr_data_frames[key] = {
                        "data": [],
                        "metadata": {},
                        "experiment_columns": [],
                        "_error": result['message']
                    }
                    st.error(f"자동 OCR 실패: {result['message']}")
                    st.info("상단 'OCR 재시도' 버튼으로 다시 시도하세요")
                    st.rerun()  # ✅ 추가 - 버튼 상태 변경 반영
        
        # ========================================
        # OCR 결과 표시
        # ========================================
        if key in st.session_state.ocr_data_frames:
            bundle = st.session_state.ocr_data_frames[key]
            
            # 에러가 있으면 표시
            if '_error' in bundle:
                st.warning(f"⚠️ 이전 OCR 시도 실패: {bundle['_error']}")
                st.info("데이터를 수정하거나 'OCR 재시도' 버튼을 클릭하세요")
            
            # 데이터가 있으면 표시
            if bundle.get('data'):
                # ========================================
                # 📋 메타데이터 편집
                # ========================================
                metadata = bundle.get('metadata', {})
                
                st.markdown("**문서 정보**")
                
                meta_data = [
                    {'항목': '처방번호', '내용': metadata.get('formula_number', '')},
                    {'항목': '제품명', '내용': metadata.get('product_name', '')},
                    {'항목': '처방특성', '내용': metadata.get('characteristics', '')}
                ]
                
                meta_df = pd.DataFrame(meta_data)
                
                edited_meta_df = st.data_editor(
                    meta_df,
                    column_config={
                        '항목': st.column_config.TextColumn("항목", width="small", disabled=True),
                        '내용': st.column_config.TextColumn("내용", width="large")
                    },
                    hide_index=True,
                    use_container_width=True,
                    key=f"meta_editor_{current_file.name}_{st.session_state.current_page}"
                )
                
                # 편집된 메타데이터 저장
                updated_metadata = {
                    'formula_number': edited_meta_df.iloc[0]['내용'],
                    'product_name': edited_meta_df.iloc[1]['내용'],
                    'characteristics': edited_meta_df.iloc[2]['내용']
                }
                st.session_state.ocr_data_frames[key]['metadata'] = updated_metadata
                
                st.markdown("---")
                
                # ========================================
                # 📊 OCR 결과 데이터 테이블
                # ========================================
                st.markdown("**OCR 결과 데이터**")

                data = bundle.get('data', [])
                if data:
                    data_copy = copy.deepcopy(data)
                    
                    # 원본 _corrections 백업
                    original_corrections = {
                        ing.get('Code', f'idx_{i}'): ing.get('_corrections', {})
                        for i, ing in enumerate(data_copy)
                    }
                    
                    # Phase 기준 정렬
                    sorted_data = sorted(data_copy, key=lambda x: x.get('Phase', ''))
                    
                    # Phase 구분 빈 행 추가
                    data_with_separators = []
                    previous_phase = None
                    
                    for ingredient in sorted_data:
                        current_phase = ingredient.get('Phase', '')
                        
                        if previous_phase and current_phase != previous_phase:
                            separator = {
                                'Phase': '',
                                'Code': '',
                                'Raw_Materials': '',
                                '_is_separator': True
                            }
                            
                            experiment_cols = bundle.get('experiment_columns', [])
                            for exp_col in experiment_cols:
                                separator[exp_col] = ''
                            
                            data_with_separators.append(separator)
                        
                        data_with_separators.append(ingredient)
                        previous_phase = current_phase
                    
                    # DataFrame 생성
                    df = pd.DataFrame(data_with_separators)
                    
                    base_cols = ['Phase', 'Code', 'Raw_Materials']
                    experiment_cols = bundle.get('experiment_columns', [])
                    
                    # DataFrame 재생성
                    df = pd.DataFrame(data_with_separators)
                    all_cols = base_cols + [col for col in experiment_cols if col in df.columns]
                    if '_is_separator' in df.columns:
                        all_cols.append('_is_separator')

                    df = df[all_cols]
                    
                    # 메모용 빈 행 추가
                    memo_data = bundle.get('memo', {})
                    memo_row = pd.DataFrame([{col: memo_data.get(col, '') for col in df.columns}])
                    df_with_memo = pd.concat([memo_row, df], ignore_index=True)
                    
                    # 컬럼 구성
                    col_config = {
                        'Phase': st.column_config.TextColumn("Phase", width="small"),
                        'Code': st.column_config.TextColumn("Code", width="small"),
                        'Raw_Materials': st.column_config.TextColumn("Raw_Materials", width="medium")
                    }
                    
                    for exp_col in experiment_cols:
                        if exp_col in df.columns:
                            col_config[exp_col] = st.column_config.TextColumn(exp_col, width="small")
                    
                    if '_is_separator' in df.columns:
                        col_config['_is_separator'] = None
                    
                    edited_df = st.data_editor(
                        df_with_memo,
                        column_config=col_config,
                        num_rows="dynamic",
                        hide_index=True,
                        key=f"data_editor_{current_file.name}_{st.session_state.current_page}",
                        use_container_width=True,
                        height=700
                    )
                    st.session_state[f'_temp_edited_df_{key}'] = edited_df
                    # 저장 시 구분선 제거 + _corrections 복원
                    if len(edited_df) > 1:
                        edited_data = []
                        
                        for _, row in edited_df.iloc[1:].iterrows():
                            ingredient = row.to_dict()
                            
                            if ingredient.get('_is_separator', False):
                                continue
                            
                            if '_is_separator' in ingredient:
                                del ingredient['_is_separator']
                            
                            code = ingredient.get('Code', '')
                            if code in original_corrections:
                                ingredient['_corrections'] = original_corrections[code]
                            
                            edited_data.append(ingredient)
                    else:
                        edited_data = []
                    
                    # st.session_state.ocr_data_frames[key]['data'] = edited_data
                                    
                    # 메모 행 저장
                    if len(edited_df) > 0:
                        memo_content = edited_df.iloc[0].to_dict()
                        # st.session_state.ocr_data_frames[key]['memo'] = memo_content
                else:
                    st.info("원료 데이터가 없습니다.")
            else:
                st.info("📋 OCR 데이터가 없습니다")
        
        else:
            st.info("🔍 OCR 시작 버튼을 클릭하여 데이터를 추출하세요")

else:
    st.info("PDF 파일을 업로드하여 시작하세요")
    
    # # ✅ 동일: 하단 통계
    # st.markdown("---")
    # st.markdown("### 전체 현황")
    
    # total_ingredients = sum(
    #     len(bundle.get('data', [])) 
    #     for bundle in st.session_state.ocr_data_frames.values()
    # )
    
    # stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
    
    # with stats_col1:
    #     st.metric("처리된 페이지", processed_pages)
    # with stats_col2:
    #     st.metric("추출된 원료", total_ingredients)
    # with stats_col3:
    #     st.metric("저장된 레시피", len(st.session_state.saved_pages))
    # with stats_col4:
    #     avg_per_page = round(total_ingredients / processed_pages, 1) if processed_pages > 0 else 0
    #     st.metric("페이지당 평균", f"{avg_per_page}개")

# else:
#     # ✅ 동일: 초기 화면
#     st.info("PDF 파일을 업로드하여 시작하세요")