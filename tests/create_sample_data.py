import pandas as pd

# 샘플 데이터 생성
data = {
    '이름': ['홍길동', '김영희', '이철수', '박미란', '정민수', '윤서연'],
    '성별': ['남', '여', '남', '여', '남', '여'],
    '생년월일': ['1990-01-15', '1992-01-22', '1988-01-30', 
              '1995-01-18', '1993-01-05', '1991-01-10']
}

# DataFrame 생성
df = pd.DataFrame(data)

# 엑셀 파일로 저장
df.to_excel('tests/test_data/sample_birthday.xlsx', index=False)
print("샘플 엑셀 파일이 생성되었습니다: tests/test_data/sample_birthday.xlsx")