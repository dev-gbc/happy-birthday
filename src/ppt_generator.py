from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from datetime import datetime
from typing import List, Dict, Tuple
import os

class PPTGeneratorError(Exception):
    """PPT 생성 관련 커스텀 에러"""
    pass

class PPTGenerator:
    def __init__(self):
        self.prs = Presentation()
        self._set_slide_size()
        
    def _set_slide_size(self):
        """슬라이드 크기를 16:9 비율로 설정"""
        try:
            self.prs.slide_width = Inches(16)
            self.prs.slide_height = Inches(9)
        except Exception as e:
            raise PPTGeneratorError(f"슬라이드 크기 설정 중 오류 발생: {str(e)}")
    
    def _validate_save_path(self, save_path: str) -> None:
        """저장 경로 검증"""
        if not os.path.exists(save_path):
            raise PPTGeneratorError(f"저장 경로가 존재하지 않습니다: {save_path}")
        if not os.access(save_path, os.W_OK):
            raise PPTGeneratorError(f"저장 경로에 쓰기 권한이 없습니다: {save_path}")
            
    def _validate_birthday_data(self, birthday_list: List[Dict]) -> None:
        """생일자 데이터 검증"""
        if not birthday_list:
            raise PPTGeneratorError("생일자 데이터가 비어있습니다")
            
        required_fields = {'이름', '성별', '생년월일', '나이'}
        for person in birthday_list:
            missing_fields = required_fields - set(person.keys())
            if missing_fields:
                raise PPTGeneratorError(f"필수 필드가 누락되었습니다: {', '.join(missing_fields)}")
                
            if not isinstance(person['이름'], str) or not person['이름'].strip():
                raise PPTGeneratorError(f"잘못된 이름 형식: {person['이름']}")
                
            if not isinstance(person['나이'], int) or person['나이'] <= 0:
                raise PPTGeneratorError(f"잘못된 나이 형식: {person['나이']}")
    
    def create_title_slide(self, month: int) -> None:
        """월별 타이틀 슬라이드 생성"""
        try:
            if not 1 <= month <= 12:
                raise PPTGeneratorError(f"잘못된 월 값: {month}")
                
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
            if not slide:
                raise PPTGeneratorError("타이틀 슬라이드 생성 실패")
            
            title = slide.shapes.title
            title.text = f"{month}월 생일자"
            
            title_format = title.text_frame.paragraphs[0].font
            title_format.name = '맑은 고딕'
            title_format.size = Pt(44)
            title_format.bold = True
            title_format.color.rgb = RGBColor(0, 120, 212)
        except Exception as e:
            raise PPTGeneratorError(f"타이틀 슬라이드 생성 중 오류: {str(e)}")
        
    def create_birthday_slide(self, person: Dict) -> None:
        """생일자 정보 슬라이드 생성"""
        try:
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            if not slide:
                raise PPTGeneratorError("생일자 슬라이드 생성 실패")
            
            # 생일자 이름 추가
            name_box = slide.shapes.add_textbox(
                left=Inches(2), top=Inches(1),
                width=Inches(12), height=Inches(1.5)
            )
            name_frame = name_box.text_frame
            name_frame.text = person['이름']
            name_para = name_frame.paragraphs[0]
            name_para.alignment = PP_ALIGN.CENTER
            name_para.font.size = Pt(40)
            name_para.font.bold = True
            name_para.font.name = '맑은 고딕'
            
            # 생일 정보 추가
            info_box = slide.shapes.add_textbox(
                left=Inches(2), top=Inches(3),
                width=Inches(12), height=Inches(2)
            )
            info_frame = info_box.text_frame
            
            try:
                birth_date = datetime.strptime(person['생년월일'], '%Y-%m-%d')
            except ValueError:
                raise PPTGeneratorError(f"잘못된 생년월일 형식: {person['생년월일']}")
            
            info_frame.text = (
                f"생년월일: {birth_date.strftime('%Y년 %m월 %d일')}\n"
                f"나이: {person['나이']}세\n"
                f"성별: {person['성별']}"
            )
            
            for paragraph in info_frame.paragraphs:
                paragraph.font.size = Pt(28)
                paragraph.font.name = '맑은 고딕'
                paragraph.alignment = PP_ALIGN.CENTER
                
        except PPTGeneratorError:
            raise
        except Exception as e:
            raise PPTGeneratorError(f"생일자 슬라이드 생성 중 오류: {str(e)}")
    
    def generate_ppt(self, month: int, birthday_list: List[Dict], save_path: str) -> Tuple[bool, str]:
        """생일자 PPT 생성"""
        try:
            # 입력값 검증
            self._validate_save_path(save_path)
            self._validate_birthday_data(birthday_list)
            
            # 타이틀 슬라이드 생성
            self.create_title_slide(month)
            
            # 생일자별 슬라이드 생성
            for person in birthday_list:
                self.create_birthday_slide(person)
            
            # PPT 저장
            output_path = f"{save_path}/{month}월_생일자.pptx"
            try:
                self.prs.save(output_path)
            except Exception as e:
                raise PPTGeneratorError(f"PPT 파일 저장 중 오류: {str(e)}")
            
            return True, "PPT 생성 완료"
            
        except PPTGeneratorError as e:
            return False, str(e)
        except Exception as e:
            return False, f"PPT 생성 중 예기치 못한 오류 발생: {str(e)}"