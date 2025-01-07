import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QPushButton, QComboBox, 
                           QFileDialog, QProgressBar, QMessageBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from excel_processor import ExcelProcessor
from ppt_generator import PPTGenerator

class BirthdayPPTApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('생일 PPT 생성기')
        self.setFixedSize(500, 400)
        
        # 전체 앱 스타일 설정
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QWidget {
                background-color: #f0f0f0;
                color: #333333;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QComboBox {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 3px;
                padding: 5px;
            }
            QProgressBar {
                background-color: #e0e0e0;
                border: 1px solid #cccccc;
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
            }
        """)
        
        # 메인 위젯 및 레이아웃 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # 제목
        title_label = QLabel('생일 PPT 생성기')
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        layout.addWidget(title_label)
        
        # 구분선 추가
        line = QLabel()
        line.setFrameStyle(QLabel.Shape.Box | QLabel.Shadow.Plain)
        line.setFixedHeight(2)
        layout.addWidget(line)
        layout.addSpacing(20)
        
        # 1. 엑셀 파일 선택
        excel_group = QWidget()
        excel_layout = QHBoxLayout()
        excel_group.setLayout(excel_layout)
        
        self.excel_path_label = QLabel('선택된 파일 없음')
        self.excel_path_label.setStyleSheet('color: #888888; background-color: white; padding: 5px; border: 1px solid #cccccc; border-radius: 3px;')
        excel_button = QPushButton('엑셀 파일 선택')
        excel_button.clicked.connect(self.select_excel)
        excel_layout.addWidget(QLabel('1. 엑셀 파일:'))
        excel_layout.addWidget(self.excel_path_label, stretch=1)
        excel_layout.addWidget(excel_button)
        self.excel_path_label.setWordWrap(True)
        layout.addWidget(excel_group)
        
        # 2. PPT 저장 위치
        save_group = QWidget()
        save_layout = QHBoxLayout()
        save_group.setLayout(save_layout)
        
        self.save_path_label = QLabel('선택된 경로 없음')
        self.save_path_label.setStyleSheet('color: #888888; background-color: white; padding: 5px; border: 1px solid #cccccc; border-radius: 3px;')
        save_button = QPushButton('저장 위치 선택')
        save_button.clicked.connect(self.select_save_path)
        save_layout.addWidget(QLabel('2. 저장 위치:'))
        save_layout.addWidget(self.save_path_label, stretch=1)
        save_layout.addWidget(save_button)
        self.save_path_label.setWordWrap(True)
        layout.addWidget(save_group)
        
        layout.addSpacing(20)
        
        # 3. 감지된 월 표시
        month_group = QWidget()
        month_layout = QHBoxLayout()
        month_group.setLayout(month_layout)
        
        month_layout.addWidget(QLabel('3. 감지된 월:'))
        self.month_label = QLabel('파일을 선택하세요')
        self.month_label.setStyleSheet('color: #888888; background-color: white; padding: 5px; border: 1px solid #cccccc; border-radius: 3px;')
        month_layout.addWidget(self.month_label)
        month_layout.addStretch()
        layout.addWidget(month_group)
        
        layout.addSpacing(20)

        
        # 진행 상태 바
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(30)
        layout.addWidget(self.progress_bar)
        
        # 생성 버튼
        generate_button = QPushButton('PPT 생성하기')
        generate_button.setFixedHeight(50)
        generate_button.clicked.connect(self.generate_ppt)
        layout.addWidget(generate_button)
        
        # 상태 메시지
        self.status_label = QLabel('파일을 선택해주세요')
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
    def select_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            excel_processor = ExcelProcessor()
            success, message = excel_processor.read_excel(file_name)
            
            if not success:
                QMessageBox.warning(self, '오류', message)
                return
                
            self.excel_path_label.setText(file_name)
            self.excel_path_label.setStyleSheet('color: #333333; background-color: white; padding: 5px; border: 1px solid #cccccc; border-radius: 3px;')
            
            # 감지된 월 표시
            self.month_label.setText(f"{excel_processor.detected_month}월")
            self.month_label.setStyleSheet('color: #333333; background-color: white; padding: 5px; border: 1px solid #cccccc; border-radius: 3px;')
            
            self.status_label.setText(message)
            
    def select_save_path(self):
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "PPT 저장 위치 선택"
        )
        if folder_path:
            self.save_path_label.setText(folder_path)
            self.save_path_label.setStyleSheet('color: #333333; background-color: white; padding: 5px; border: 1px solid #cccccc; border-radius: 3px;')
            self.status_label.setText(f'저장 위치가 선택되었습니다: {folder_path}')
            
    def generate_ppt(self):
        # 입력 검증
        if self.excel_path_label.text() == '선택된 파일 없음':
            QMessageBox.warning(self, '경고', '엑셀 파일을 선택해주세요.')
            return
        if self.save_path_label.text() == '선택된 경로 없음':
            QMessageBox.warning(self, '경고', '저장 위치를 선택해주세요.')
            return
        
        # 엑셀 파일 처리
        self.status_label.setText('엑셀 파일 읽는 중...')
        self.progress_bar.setValue(10)
        
        excel_processor = ExcelProcessor()
        success, message = excel_processor.read_excel(self.excel_path_label.text())
        
        if not success:
            QMessageBox.warning(self, '오류', message)
            self.status_label.setText('엑셀 파일 처리 실패')
            self.progress_bar.setValue(0)
            return
            
        # 생일자 목록 가져오기
        birthday_list = excel_processor.get_birthdays()
        
        if not birthday_list:
            QMessageBox.information(self, '알림', '생일자 데이터가 없습니다.')
            self.status_label.setText('데이터 없음')
            self.progress_bar.setValue(0)
            return
            
        # PPT 생성
        self.status_label.setText('PPT 생성 중...')
        self.progress_bar.setValue(50)
        
        # PPT 생성
        self.status_label.setText('PPT 생성 중...')
        self.progress_bar.setValue(50)
        
        ppt_generator = PPTGenerator() 
        success = ppt_generator.generate_ppt(
            excel_processor.detected_month,
            birthday_list,
            self.save_path_label.text()
        )
        
        if success:
            self.status_label.setText('PPT 생성 완료')
            self.progress_bar.setValue(100)
            QMessageBox.information(self, '완료', 'PPT 파일이 생성되었습니다.')
        else:
            self.status_label.setText('PPT 생성 실패')
            self.progress_bar.setValue(0)
            QMessageBox.warning(self, '오류', 'PPT 생성 중 오류가 발생했습니다.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = BirthdayPPTApp()
    ex.show()
    sys.exit(app.exec())