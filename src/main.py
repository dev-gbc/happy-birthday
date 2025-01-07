import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                           QProgressBar, QMessageBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from excel_processor import ExcelProcessor
from ppt_generator import PPTGenerator

class BirthdayPPTApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_path_label = None
        self.save_path_label = None
        self.month_label = None
        self.progress_bar = None
        self.status_label = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('생일 PPT 생성기')
        self.setFixedSize(500, 400)
        
        # 스타일 설정
        self.setStyleSheet("""
        QMainWindow, QWidget {
            background-color: #FFFFFF;
            font-family: Pretendard;
        }
        
        QPushButton {
            background-color: #2563EB;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            font-weight: 500;
            font-size: 13px;
            min-width: 100px;
        }
        
        QPushButton:hover {
            background-color: #1D4ED8;
        }
        
        QPushButton:pressed {
            background-color: #1E40AF;
        }
        
        QLabel.path-label {
            background-color: #F3F4F6;
            border: 1px solid #E5E7EB;
            border-radius: 6px;
            padding: 8px;
            color: #6B7280;
        }
        
        QLabel.section-label {
            color: #374151;
            font-weight: 500;
            min-width: 100px;
        }
        
        QProgressBar {
            background-color: #E5E7EB;
            border: none;
            border-radius: 4px;
            height: 8px;
            text-align: center;
        }
        
        QProgressBar::chunk {
            background-color: #2563EB;
            border-radius: 4px;
        }
        """)
        
        # 메인 위젯 및 레이아웃 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(20)
        main_widget.setLayout(layout)
        
        # 제목
        title_label = QLabel('생일 PPT 생성기')
        title_font = QFont("Pretendard", 17, QFont.Weight.Bold)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #111827; margin: 0px 0;")
        layout.addWidget(title_label)
        
        # 구분선
        line = QLabel()
        line.setStyleSheet("background-color: #E5E7EB; min-height: 1px; max-height: 1px;")
        layout.addWidget(line)
        layout.addSpacing(16)
        
        # 1. 엑셀 파일 선택
        excel_group = QWidget()
        excel_layout = QHBoxLayout()
        excel_layout.setContentsMargins(0, 0, 0, 0)
        excel_group.setLayout(excel_layout)
        
        excel_label = QLabel('1. 엑셀 파일:')
        excel_label.setProperty('class', 'section-label')
        excel_layout.addWidget(excel_label)
        
        self.excel_path_label = QLabel('선택된 파일 없음')
        self.excel_path_label.setProperty('class', 'path-label')
        self.excel_path_label.setWordWrap(True)
        excel_layout.addWidget(self.excel_path_label, stretch=1)
        
        excel_button = QPushButton('파일 선택')
        excel_button.clicked.connect(self.select_excel)
        excel_layout.addWidget(excel_button)
        
        layout.addWidget(excel_group)
        
        # 2. PPT 저장 위치
        save_group = QWidget()
        save_layout = QHBoxLayout()
        save_layout.setContentsMargins(0, 0, 0, 0)
        save_group.setLayout(save_layout)
        
        save_label = QLabel('2. 저장 위치:')
        save_label.setProperty('class', 'section-label')
        save_layout.addWidget(save_label)
        
        self.save_path_label = QLabel('선택된 경로 없음')
        self.save_path_label.setProperty('class', 'path-label')
        self.save_path_label.setWordWrap(True)
        save_layout.addWidget(self.save_path_label, stretch=1)
        
        save_button = QPushButton('위치 선택')
        save_button.clicked.connect(self.select_save_path)
        save_layout.addWidget(save_button)
        
        layout.addWidget(save_group)
        
        # 3. 감지된 월 표시
        month_group = QWidget()
        month_layout = QHBoxLayout()
        month_layout.setContentsMargins(0, 0, 0, 0)
        month_group.setLayout(month_layout)
        
        month_label = QLabel('3. 감지된 월:')
        month_label.setProperty('class', 'section-label')
        month_layout.addWidget(month_label)
        
        self.month_label = QLabel('파일을 선택하세요')
        self.month_label.setProperty('class', 'path-label')
        month_layout.addWidget(self.month_label)
        month_layout.addStretch()
        
        layout.addWidget(month_group)
        
        layout.addSpacing(16)
        
        # 진행 상태 바
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(8)
        layout.addWidget(self.progress_bar)
        
        # 생성 버튼
        generate_button = QPushButton('PPT 생성하기')
        generate_button.setFixedHeight(50)
        generate_button.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                font-weight: bold;
            }
        """)
        generate_button.clicked.connect(self.generate_ppt)
        layout.addWidget(generate_button)
        
        # 상태 메시지
        self.status_label = QLabel('파일을 선택해주세요')
        self.status_label.setStyleSheet("color: #6B7280; font-size: 13px;")
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
            self.excel_path_label.setStyleSheet("""
                background-color: white;
                border: 1px solid #2563EB;
                border-radius: 6px;
                padding: 8px;
                color: #374151;
            """)
            
            # 감지된 월 표시
            self.month_label.setText(f"{excel_processor.detected_month}월")
            self.month_label.setStyleSheet("""
                background-color: white;
                border: 1px solid #2563EB;
                border-radius: 6px;
                padding: 8px;
                color: #374151;
            """)
            
            self.status_label.setText(message)
            
    def select_save_path(self):
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "PPT 저장 위치 선택"
        )
        if folder_path:
            self.save_path_label.setText(folder_path)
            self.save_path_label.setStyleSheet("""
                background-color: white;
                border: 1px solid #2563EB;
                border-radius: 6px;
                padding: 8px;
                color: #374151;
            """)
            self.status_label.setText(f'저장 위치가 선택되었습니다: {folder_path}')
            
    def generate_ppt(self):
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
        
        ppt_generator = PPTGenerator(font_name="Pretendard")
        success, message = ppt_generator.generate_ppt(
            excel_processor.detected_month,
            birthday_list,
            self.save_path_label.text()
        )
        
        if success:
            self.status_label.setText('PPT 생성 완료')
            self.progress_bar.setValue(100)
            QMessageBox.information(self, '완료', message)
        else:
            self.status_label.setText('PPT 생성 실패')
            self.progress_bar.setValue(0)
            QMessageBox.warning(self, '오류', message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = BirthdayPPTApp()
    ex.show()
    sys.exit(app.exec())