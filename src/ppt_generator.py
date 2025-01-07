from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from datetime import datetime
from typing import List, Dict

class PPTGenerator:
    def __init__(self):
        self.prs = Presentation()
        self._set_slide_size()
        
    def _set_slide_size(self):
        """슬라이드 크기를 16:9 비율로 설정"""
        self.prs.slide_width = Inches(16)
        self.prs.slide_height = Inches(9)
    
    def create_title_slide(self, month: int) -> None:
        """월별 타이틀 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
        
        title = slide.shapes.title
        title.text = f"{month}월 생일자"
        
        title_format = title.text_frame.paragraphs[0].font
        title_format.name = '맑은 고딕'
        title_format.size = Pt(44)
        title_format.bold = True
        title_format.color.rgb = RGBColor(0, 120, 212)
        
    def create_birthday_slide(self, person: Dict) -> None:
        """생일자 정보 슬라이드 생성"""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        
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
        
        birth_date = datetime.strptime(person['생년월일'], '%Y-%m-%d')
        info_frame.text = (
            f"생년월일: {birth_date.strftime('%Y년 %m월 %d일')}\n"
            f"나이: {person['나이']}세\n"
            f"성별: {person['성별']}"
        )
        
        for paragraph in info_frame.paragraphs:
            paragraph.font.size = Pt(28)
            paragraph.font.name = '맑은 고딕'
            paragraph.alignment = PP_ALIGN.CENTER
    
    def generate_ppt(self, month: int, birthday_list: List[Dict], save_path: str) -> bool:
        """생일자 PPT 생성"""
        try:
            if not birthday_list:
                return False
            
            # 타이틀 슬라이드 생성
            self.create_title_slide(month)
            
            # 생일자별 슬라이드 생성
            for person in birthday_list:
                self.create_birthday_slide(person)
            
            # PPT 저장
            self.prs.save(f"{save_path}/{month}월_생일자.pptx")
            return True
            
        except Exception as e:
            print(f"PPT 생성 중 오류 발생: {str(e)}")
            return False