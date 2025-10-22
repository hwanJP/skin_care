"""
제형 레시피 OCR 백엔드 (최소 구현)
기존 backend.py의 PDFProcessor만 재사용
"""

import io
import logging
import os
import tempfile  # 🔧 추가!

# ✅ 기존 backend에서 PDFProcessor만 import
from backend import PDFProcessor

# 🆕 Azure OCR
from azure_ai import KolmarCosmeticOCR

logger = logging.getLogger(__name__)


def process_recipe_page(pdf_bytes: bytes, page_index: int) -> dict:
    """
    제형 레시피 페이지 처리 (간소화)
    """
    result = {
        'success': False,
        'data': [],
        'metadata': {},
        'experiment_columns': [],
        'message': ''
    }
    
    temp_image_path = None
    
    try:
        # 1. DRM 처리
        drm_success, processed_bytes, drm_message = PDFProcessor.process_drm_if_needed(pdf_bytes)
        if not drm_success:
            result['message'] = drm_message
            return result
        
        logger.info(f"📄 DRM 처리: {drm_message}")
        
        # 2. 이미지 렌더링
        img_bytes = PDFProcessor.render_page_image(processed_bytes, page_index, zoom=2.0)
        if not img_bytes:
            result['message'] = "이미지 렌더링 실패"
            return result
        
        # 3. 임시 파일 저장
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
            temp_image_path = temp_file.name
            temp_file.write(img_bytes)
        
        logger.info(f"💾 임시 이미지 저장: {temp_image_path}")
        
        # 4. Azure OCR
        ocr = KolmarCosmeticOCR()
        formula_data = ocr.extract_cosmetic_formula_table(temp_image_path)
        
        if not formula_data or not formula_data.get('ingredients'):
            result['message'] = "데이터 추출 실패"
            return result
        
        # 5. 결과 포맷팅
        result['success'] = True
        result['data'] = formula_data['ingredients']
        result['metadata'] = {
            'formula_number': formula_data.get('formula_number', ''),
            'product_name': formula_data.get('product_name', ''),
            'characteristics': formula_data.get('characteristics', '')
        }
        
        # 🔧 experiment_columns 처리
        exp_cols = formula_data.get('experiment_columns', [])
        
        # Col_4, Col_5 형태면 자동으로 알파벳 생성
        if exp_cols and all(col.startswith('Col_') for col in exp_cols):
            logger.warning(f"⚠️ 기본 컬럼명 감지: {exp_cols}")
            
            # 알파벳으로 변환 (A부터 시작하거나 U부터 시작)
            # 보통 U, V, W, X, Y, Z 패턴이므로 U부터 시작
            alphabet_start = ord('U')
            new_exp_cols = []
            
            for i, old_col in enumerate(exp_cols):
                # U, V, W, X, Y, Z, AA, AB...
                if i < 26:
                    new_col = chr(alphabet_start + i)
                else:
                    # 26개 넘으면 AA, AB...
                    first = chr(alphabet_start + (i // 26) - 1)
                    second = chr(alphabet_start + (i % 26))
                    new_col = first + second
                
                new_exp_cols.append(new_col)
                
                # 데이터에서도 컬럼명 변경
                for ingredient in result['data']:
                    if old_col in ingredient:
                        ingredient[new_col] = ingredient.pop(old_col)
            
            exp_cols = new_exp_cols
            logger.info(f"✅ 컬럼명 자동 변환: {exp_cols}")
        
        # 빈 리스트면 원료 데이터에서 추출
        elif not exp_cols and result['data']:
            first_ingredient = result['data'][0]
            base_cols = ['Phase', 'Code', 'Raw_Materials']
            exp_cols = [col for col in first_ingredient.keys() if col not in base_cols]
            logger.info(f"🔧 experiment_columns 자동 생성: {exp_cols}")
        
        result['experiment_columns'] = exp_cols
        
        logger.info(f"✅ OCR 성공: {len(result['data'])}개 원료, 실험 컬럼: {exp_cols}")
        result['message'] = f"{len(formula_data['ingredients'])}개 원료 추출 완료"
        
        return result
        
    except Exception as e:
        logger.error(f"❌ 처리 오류: {e}")
        import traceback
        traceback.print_exc()
        result['message'] = str(e)
        return result
    
    finally:
        # 임시 파일 정리
        if temp_image_path and os.path.exists(temp_image_path):
            try:
                os.remove(temp_image_path)
                logger.info(f"🗑️ 임시 파일 삭제: {temp_image_path}")
            except Exception as e:
                logger.warning(f"⚠️ 임시 파일 삭제 실패: {e}")



class RecipeExcelSaver:
    """제형 레시피 Excel 저장 (단순화)"""
    
    def __init__(self, output_path: str):
        self.output_path = output_path
        
        if not os.path.exists(self.output_path):
            from openpyxl import Workbook
            wb = Workbook()
            # ❌ 기존: wb.remove(wb.active)  # 이 줄이 문제!
            # ✅ 수정: 기본 시트 그대로 두기
            wb.save(self.output_path)
            wb.close()
    
    def add_recipe_data(self, data, metadata, experiment_cols):
        """제형 데이터 추가"""
        try:
            from openpyxl import load_workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            import pandas as pd
            
            if not data:
                return False
            
            # DataFrame 생성
            df = pd.DataFrame(data)
            base_cols = ['Phase', 'Code', 'Raw_Materials']
            exp_cols = [col for col in experiment_cols if col in df.columns]
            df = df[base_cols + exp_cols]
            
            workbook = load_workbook(self.output_path)
            
            # 시트명: 처방번호
            sheet_name = metadata.get('formula_number', 'Recipe')
            counter = 1
            original_name = sheet_name
            while sheet_name in workbook.sheetnames:
                sheet_name = f"{original_name}_{counter}"
                counter += 1
            
            worksheet = workbook.create_sheet(title=sheet_name)
            
            # 스타일
            info_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
            info_font = Font(bold=True, size=10)
            header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            header_font = Font(bold=True, color='FFFFFF', size=11)
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            
            # 상단 정보 (1-3행)
            doc_info = [
                ['처방번호', metadata.get('formula_number', '')],
                ['제품명', metadata.get('product_name', '')],
                ['처방특성', metadata.get('characteristics', '')]
            ]
            
            for row_idx, (label, value) in enumerate(doc_info, start=1):
                cell_label = worksheet.cell(row=row_idx, column=1, value=label)
                cell_label.fill = info_fill
                cell_label.font = info_font
                cell_label.border = thin_border
                
                cell_value = worksheet.cell(row=row_idx, column=2, value=value)
                cell_value.border = thin_border
                worksheet.merge_cells(start_row=row_idx, start_column=2, 
                                    end_row=row_idx, end_column=len(df.columns))
            
            # 헤더 (6행)
            header_row = 6
            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = worksheet.cell(row=header_row, column=col_idx, value=col_name)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 데이터 (8행부터)
            data_start_row = 8
            for df_row_idx, row_data in df.iterrows():
                excel_row = data_start_row + df_row_idx
                for col_idx, value in enumerate(row_data, start=1):
                    cell = worksheet.cell(row=excel_row, column=col_idx, value=value)
                    cell.border = thin_border
            
            # 열 너비 조정
            for col_idx in range(1, len(df.columns) + 1):
                max_length = 10
                col_letter = get_column_letter(col_idx)
                for row_idx in range(1, data_start_row + len(df)):
                    cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))
                worksheet.column_dimensions[col_letter].width = min(max_length + 2, 50)
            
            worksheet.freeze_panes = 'D8'
            
            workbook.save(self.output_path)
            workbook.close()
            
            logger.info(f"💾 Excel 저장: {sheet_name} ({len(df)}개 원료)")
            return True
            
        except Exception as e:
            logger.error(f"❌ Excel 저장 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_excel_bytes(self):
        """Excel 바이트 반환"""
        try:
            if os.path.exists(self.output_path):
                with open(self.output_path, 'rb') as f:
                    return f.read()
        except:
            pass
        return None
    
    def get_statistics(self):
        """통계 반환"""
        try:
            from openpyxl import load_workbook
            if os.path.exists(self.output_path):
                wb = load_workbook(self.output_path, read_only=True)
                sheet_count = len(wb.sheetnames)
                wb.close()
                
                file_size = os.path.getsize(self.output_path)
                return {
                    'test_sheets': sheet_count,
                    'file_size': file_size,
                    'file_size_mb': round(file_size / (1024 * 1024), 2)
                }
        except:
            pass
        
        return {'test_sheets': 0, 'file_size': 0, 'file_size_mb': 0}