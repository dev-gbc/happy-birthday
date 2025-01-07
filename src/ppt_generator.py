from pptx import Presentation
from datetime import datetime
from typing import List, Dict, Tuple
import os
from pptx.enum.shapes import MSO_SHAPE_TYPE
from io import BytesIO

class PPTGeneratorError(Exception):
    pass

class PPTGenerator:
    def __init__(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.template_path = os.path.join(current_dir, '..', 'resources', 'templates', 'template.pptx')
        
        if not os.path.exists(self.template_path):
            raise PPTGeneratorError(f"템플릿 파일을 찾을 수 없습니다: {self.template_path}")
            
        try:
            self.prs = Presentation(self.template_path)
            # 템플릿 구조 출력
            print(f"템플릿 슬라이드 수: {len(self.prs.slides)}")
            for idx, slide in enumerate(self.prs.slides):
                print(f"\n슬라이드 {idx + 1} 분석:")
                print(f"- 레이아웃: {slide.slide_layout.name}")
                print("- 도형 목록:")
                for shape_idx, shape in enumerate(slide.shapes):
                    print(f"  도형 {shape_idx + 1}:")
                    print(f"    유형: {shape.shape_type}")
                    if hasattr(shape, 'name'):
                        print(f"    이름: {shape.name}")
                    if shape.has_text_frame:
                        print(f"    텍스트: {shape.text}")
        except Exception as e:
            raise PPTGeneratorError(f"템플릿 분석 실패: {str(e)}")

    def create_birthday_slide(self, person: Dict) -> None:
        """생일자 슬라이드 생성"""
        try:
            template_slide = self.prs.slides[1]
            new_slide = self.prs.slides.add_slide(template_slide.slide_layout)
            
            # 템플릿의 도형들 복사
            for shape in template_slide.shapes:
                if hasattr(shape, 'shape_type'):
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        # 이미지 복사
                        image = new_slide.shapes.add_picture(
                            image_file=BytesIO(shape.image.blob),
                            left=left, top=top,
                            width=width, height=height
                        )
                    elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                        # 텍스트박스 복사
                        textbox = new_slide.shapes.add_textbox(
                            left=left, top=top,
                            width=width, height=height
                        )
                        
                        if shape.has_text_frame and shape.text:
                            # 텍스트와 서식 복사
                            text_frame = textbox.text_frame
                            text_frame.text = shape.text_frame.text
                            
                            # 생일자 정보로 텍스트 교체
                            birth_date = datetime.strptime(person['생년월일'], '%Y-%m-%d')
                            text = text_frame.text
                            text = text.replace("{name}", person['이름'])
                            text = text.replace("{month}", str(birth_date.month))
                            text = text.replace("{day}", str(birth_date.day))
                            text_frame.text = text
                            
                            # 폰트 복사 (선택적)
                            if len(text_frame.paragraphs) > 0:
                                p = text_frame.paragraphs[0]
                                if len(p.runs) > 0:
                                    font = p.runs[0].font
                                    font.name = '맑은 고딕'  # 기본 폰트 지정
                                    font.size = shape.text_frame.paragraphs[0].runs[0].font.size
                                    
        except Exception as e:
            print(f"슬라이드 생성 중 오류: {str(e)}")
            raise PPTGeneratorError(f"슬라이드 생성 오류: {str(e)}")

    def create_title_slide(self, month: int) -> None:
        try:
            print(f"\n타이틀 슬라이드 수정 (월: {month})")
            title_slide = self.prs.slides[0]
            
            print("현재 도형들의 텍스트:")
            for shape in title_slide.shapes:
                if shape.has_text_frame:
                    print(f"- {shape.text}")
                    
            for shape in title_slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    if "{month}" in text_frame.text:
                        original_text = text_frame.text
                        new_text = original_text.replace("{month}", str(month))
                        text_frame.text = new_text
                        print(f"텍스트 교체: {original_text} -> {new_text}")

        except Exception as e:
            print(f"타이틀 슬라이드 수정 중 오류: {str(e)}")
            raise PPTGeneratorError(f"타이틀 슬라이드 수정 오류: {str(e)}")

    def generate_ppt(self, month: int, birthday_list: List[Dict], save_path: str) -> Tuple[bool, str]:
        try:
            print(f"\nPPT 생성 시작:")
            print(f"- 월: {month}")
            print(f"- 생일자 수: {len(birthday_list)}")
            print(f"- 저장 경로: {save_path}")
            
            self._validate_save_path(save_path)
            self._validate_birthday_data(birthday_list)
            
            self.create_title_slide(month)
            
            for person in birthday_list:
                self.create_birthday_slide(person)
            
            # 템플릿 슬라이드 제거
            xml_slides = self.prs.slides._sldIdLst
            slides = list(xml_slides)
            xml_slides.remove(slides[1])
            print("템플릿 슬라이드 제거됨")
            
            output_path = os.path.join(save_path, f"{month}월_생일자.pptx")
            self.prs.save(output_path)
            print(f"파일 저장 완료: {output_path}")
            
            return True, f"PPT 파일이 생성되었습니다: {output_path}"
            
        except Exception as e:
            print(f"PPT 생성 실패: {str(e)}")
            return False, f"PPT 생성 실패: {str(e)}"

    def _validate_save_path(self, save_path: str) -> None:
        if not os.path.exists(save_path):
            raise PPTGeneratorError(f"저장 경로가 존재하지 않습니다: {save_path}")
        if not os.access(save_path, os.W_OK):
            raise PPTGeneratorError(f"저장 경로에 쓰기 권한이 없습니다: {save_path}")

    def _validate_birthday_data(self, birthday_list: List[Dict]) -> None:
        if not birthday_list:
            raise PPTGeneratorError("생일자 데이터가 비어있습니다")
        
        required_fields = {'이름', '성별', '생년월일', '나이'}
        for person in birthday_list:
            missing_fields = required_fields - set(person.keys())
            if missing_fields:
                raise PPTGeneratorError(f"필수 필드가 누락되었습니다: {', '.join(missing_fields)}")