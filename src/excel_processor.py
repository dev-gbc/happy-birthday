import pandas as pd
from datetime import datetime
from typing import List, Dict, Tuple

class ExcelProcessor:
    REQUIRED_COLUMNS = ['이름', '성별', '생년월일']
    
    def __init__(self):
        self.df = None
        
    def validate_columns(self, df: pd.DataFrame) -> bool:
        """필수 컬럼이 모두 있는지 확인"""
        return all(col in df.columns for col in self.REQUIRED_COLUMNS)
    
    def validate_date_format(self, date_str: str) -> bool:
        """날짜 형식이 YYYY-MM-DD 형식인지 확인"""
        try:
            datetime.strptime(str(date_str), '%Y-%m-%d')
            return True
        except ValueError:
            return False
    
    def read_excel(self, file_path: str) -> Tuple[bool, str]:
        """엑셀 파일 읽기 및 검증"""
        try:
            self.df = pd.read_excel(file_path)
            
            # 필수 컬럼 검증
            if not self.validate_columns(self.df):
                missing_cols = [col for col in self.REQUIRED_COLUMNS if col not in self.df.columns]
                return False, f"필수 컬럼이 없습니다: {', '.join(missing_cols)}"
            
            # 데이터 형식 검증
            invalid_dates = []
            for idx, row in self.df.iterrows():
                if not self.validate_date_format(str(row['생년월일'])):
                    invalid_dates.append(f"{row['이름']}: {row['생년월일']}")
            
            if invalid_dates:
                return False, f"잘못된 날짜 형식이 있습니다:\n{chr(10).join(invalid_dates)}"
            
            # 데이터 전처리
            self.df['생년월일'] = pd.to_datetime(self.df['생년월일'])
            
            return True, "파일 검증 성공"
            
        except Exception as e:
            return False, f"파일 읽기 오류: {str(e)}"
    
    def get_birthdays_by_month(self, month: int) -> List[Dict]:
        """특정 월의 생일자 목록 반환"""
        if self.df is None:
            return []
        
        # 해당 월의 생일자만 필터링
        monthly_df = self.df[self.df['생년월일'].dt.month == month].copy()
        
        # 결과를 리스트로 변환
        birthday_list = []
        for _, row in monthly_df.iterrows():
            birthday_list.append({
                '이름': row['이름'],
                '성별': row['성별'],
                '생년월일': row['생년월일'].strftime('%Y-%m-%d'),
                '나이': datetime.now().year - row['생년월일'].year + 1
            })
            
        return birthday_list