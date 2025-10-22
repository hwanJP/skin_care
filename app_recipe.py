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
    page_title="제형 레시피 OCR 도구",  # 🔧 제목만 변경
    layout="wide",
    initial_sidebar_state="collapsed"
)

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

# 🔧 수정: ExcelIncrementalSaver → RecipeExcelSaver
if "excel_saver" not in st.session_state:
    temp_dir = tempfile.gettempdir()
    excel_path = os.path.join(temp_dir, f"제형레시피_{st.session_state.session_id}.xlsx")
    st.session_state.excel_saver = RecipeExcelSaver(excel_path)
    st.session_state.excel_path = excel_path

# ========================================
# ✅ 동일: CSS 스타일
# ========================================
st.markdown("""
<style>
    .compact-header {
        background: linear-gradient(90deg, #0066cc 0%, #0099ff 100%);
        padding: 0.5rem 1rem;
        border-radius: 5px;
        color: white;
        margin-bottom: 1rem;
    }
    /* ... 나머지 CSS 동일 ... */
</style>
""", unsafe_allow_html=True)

# ========================================
# 🔧 수정: 헤더 (제목만 변경)
# ========================================
st.markdown("""
<div class="compact-header">
    <h1>제형 레시피 OCR 도구</h1>
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
                        
                        # ✅ PDFProcessor 재사용
                        drm_success, processed_bytes, drm_message = PDFProcessor.process_drm_if_needed(original_bytes)
                        
                        if not drm_success:
                            st.error(f"❌ 파일 처리 실패: {drm_message}")
                            logger.error(f"DRM 처리 실패: {drm_message}")
                            st.stop()
                        
                        try:
                            doc = fitz.open(stream=processed_bytes, filetype="pdf")
                            page_count = doc.page_count
                            doc.close()
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
                            st.success(f"✅ {drm_message} | 총 {page_count} 페이지")
                        else:
                            st.success(f"✅ 파일 로드 완료 | 총 {page_count} 페이지")
                
                processed_file_info = st.session_state.processed_files[file_id]
                st.session_state.current_file_name = uploaded_file.name
                st.session_state.current_file_bytes = processed_file_info['bytes']
                st.session_state.current_file_id = file_id
                st.session_state.current_page = 1
                st.rerun()

# ========================================
# ✅ 동일: 새로 시작하기 버튼
# ========================================
with header_col2:
    if has_work:
        if st.button("🔄 새로 시작하기", use_container_width=True, type="secondary"):
            if st.session_state.get('confirm_reset', False):
                st.session_state.ocr_data_frames = {}
                st.session_state.saved_pages = set()
                st.session_state.current_page = 1
                st.session_state.current_file_name = None
                st.session_state.current_file_bytes = None
                st.session_state.current_file_id = None
                st.session_state.confirm_reset = False
                
                # 🔧 Excel 초기화 (RecipeExcelSaver)
                temp_dir = tempfile.gettempdir()
                excel_path = os.path.join(temp_dir, f"제형레시피_{st.session_state.session_id}.xlsx")
                st.session_state.excel_saver = RecipeExcelSaver(excel_path)
                st.session_state.excel_path = excel_path
                
                st.success("✅ 새로 시작합니다")
                st.rerun()
            else:
                st.session_state.confirm_reset = True
                st.warning("⚠️ 작업 내용이 삭제됩니다. 다시 클릭하면 초기화됩니다.")
                st.rerun()

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
    
    with action_col1:
        if st.button("OCR 시작", type="primary", use_container_width=True):
            with st.spinner(f"페이지 {st.session_state.current_page} 처리 중..."):
                # 🔧 수정: process_pdf_page → process_recipe_page
                result = process_recipe_page(
                    current_file.getvalue(), 
                    st.session_state.current_page - 1
                )
                
                if result['success']:
                    key = (current_file.name, st.session_state.current_page)
                    
                    # 🔧 수정: 데이터 구조 변경
                    st.session_state.ocr_data_frames[key] = {
                        "data": result['data'],
                        "metadata": result['metadata'],
                        "experiment_columns": result['experiment_columns']
                    }
                    
                    st.success(result['message'])
                    st.rerun()
                else:
                    st.error(f"처리 실패: {result['message']}")
    
    with action_col2:
        key = (current_file.name, st.session_state.current_page)
        if key in st.session_state.ocr_data_frames:
            if st.button("Excel에 저장", use_container_width=True):
                bundle = st.session_state.ocr_data_frames[key]
                
                # 🔧 수정: experiment_columns → experiment_cols
                success = st.session_state.excel_saver.add_recipe_data(
                    data=bundle['data'],
                    metadata=bundle['metadata'],
                    experiment_cols=bundle['experiment_columns']  # ← 파라미터명 수정
                )
                
                if success:
                    st.session_state.saved_pages.add(key)
                    st.success(f"{len(bundle['data'])}개 원료가 저장되었습니다")
                    st.rerun()
                else:
                    st.error("Excel 저장 실패")
        else:
            st.button("Excel에 저장", use_container_width=True, disabled=True)
    
    # ✅ 동일: action_col3, action_col4, action_col5 (코드 동일)
    with action_col3:
        if st.session_state.excel_saver:
            stats = st.session_state.excel_saver.get_statistics()
            sheet_count = stats['test_sheets']
        else:
            sheet_count = 0
        st.button(f"저장: {sheet_count}개", use_container_width=True, disabled=True)
    
    with action_col4:
        key = (current_file.name, st.session_state.current_page)
        is_saved = key in st.session_state.saved_pages
        has_data = key in st.session_state.ocr_data_frames
        
        if has_data and not is_saved:
            st.button("다음", use_container_width=True, disabled=True)
            st.caption("저장 후 이동")
        else:
            if st.button("다음", use_container_width=True):
                if st.session_state.current_page < page_count:
                    st.session_state.current_page += 1
                    st.rerun()
    
    with action_col5:
        if os.path.exists(st.session_state.excel_path):
            excel_bytes = st.session_state.excel_saver.get_excel_bytes()
            if excel_bytes:
                stats = st.session_state.excel_saver.get_statistics()
                file_size_mb = stats.get('file_size_mb', 0)
                
                st.download_button(
                    label=f"Excel 다운로드 ({file_size_mb}MB)",
                    data=excel_bytes,
                    file_name=f"제형레시피_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.button("Excel 다운로드", use_container_width=True, disabled=True)
    
    # ✅ 동일: 상태 표시줄
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
        
        try:
            doc = fitz.open(stream=current_file.getvalue(), filetype="pdf")
            page = doc.load_page(st.session_state.current_page - 1)
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")
            st.image(img_bytes, use_column_width=True)
            doc.close()
        except Exception as e:
            st.error(f"PDF 렌더링 오류: {e}")

    # 🔧 수정: 우측 OCR 결과 (더 넓은 공간)
    with right_col:
        st.markdown("### OCR 결과")
        
        key = (current_file.name, st.session_state.current_page)
        
        if key in st.session_state.ocr_data_frames:
            bundle = st.session_state.ocr_data_frames[key]
            
            # ========================================
            # 📋 메타데이터 편집 가능 (상단)
            # ========================================
            metadata = bundle.get('metadata', {})
            
            st.markdown("**문서 정보**")
            
            # 메타데이터를 DataFrame으로 만들어 편집 가능하게
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
            
            st.markdown("---")  # 구분선
            
            # ========================================
            # 📊 OCR 결과 데이터 테이블
            # ========================================
            st.markdown("**OCR 결과 데이터**")
            
            data = bundle.get('data', [])
            if data:
                df = pd.DataFrame(data)
                
                # 컬럼 순서
                base_cols = ['Phase', 'Code', 'Raw_Materials']
                experiment_cols = bundle.get('experiment_columns', [])
                
                # 실제 컬럼명 (U, V, W, X, Y, Z 등)
                all_cols = base_cols + [col for col in experiment_cols if col in df.columns]
                df = df[all_cols]
                
                # ========================================
                # 🆕 메모용 빈 행 추가 (헤더 바로 아래)
                # ========================================
                memo_row = pd.DataFrame([{col: '' for col in df.columns}])
                df_with_memo = pd.concat([memo_row, df], ignore_index=True)
                
                # 컬럼 구성
                col_config = {
                    'Phase': st.column_config.TextColumn("Phase", width="small"),
                    'Code': st.column_config.TextColumn("Code", width="small"),
                    'Raw_Materials': st.column_config.TextColumn("Raw_Materials", width="medium")
                }
                
                # 실험 컬럼 동적 추가
                for exp_col in experiment_cols:
                    if exp_col in df.columns:
                        col_config[exp_col] = st.column_config.TextColumn(
                            exp_col,
                            width="small"
                        )
                
                edited_df = st.data_editor(
                    df_with_memo,
                    column_config=col_config,
                    num_rows="dynamic",
                    hide_index=True,
                    key=f"data_editor_{current_file.name}_{st.session_state.current_page}",
                    use_container_width=True,
                    height=700
                )
                
                # 편집된 데이터 저장 (메모 행 제외)
                if len(edited_df) > 1:
                    edited_data = edited_df.iloc[1:].to_dict('records')
                else:
                    edited_data = []
                
                st.session_state.ocr_data_frames[key]['data'] = edited_data
                
                # 메모 행 내용도 저장
                if len(edited_df) > 0:
                    memo_content = edited_df.iloc[0].to_dict()
                    st.session_state.ocr_data_frames[key]['memo'] = memo_content
            else:
                st.info("원료 데이터가 없습니다.")
        else:
            st.info("OCR 결과 데이터가 없습니다. OCR 시작 버튼을 클릭하세요.")
    
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

else:
    # ✅ 동일: 초기 화면
    st.info("PDF 파일을 업로드하여 시작하세요")