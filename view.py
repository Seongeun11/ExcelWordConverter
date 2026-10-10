#view.py
from PySide6.QtWidgets import (
    QMainWindow, QSpinBox, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QLineEdit, QTextEdit, QGroupBox, QScrollArea
)

class ExcelWordView(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to JSON & Word Converter (MVC)")
        self.resize(900, 800)
        self.file_inputs = {}  # {Path: QLineEdit}
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 프로그램 안내 및 R1C1 참조 스타일 설정 방법
        desc_group = QGroupBox("프로그램 안내 및 R1C1 참조 스타일 설정 방법")
        desc_layout = QVBoxLayout()
        desc_label = QLabel(
            "엑셀의 열 번호를 입력하면 그 열의 모든 정보를 불러옵니다.\n\n"
            "※ 엑셀 상단 R1C1 참조 스타일 설정 방법:\n"
            "1. 엑셀 상단의 [파일] 메뉴 클릭\n"
            "2. 하단의 [옵션] 클릭\n"
            "3. 왼쪽 메뉴에서 [수식] 선택\n"
            "4. '수식 작업' 항목에서 [R1C1 참조 스타일] 체크박스 선택 후 [확인]"
        )
        desc_layout.addWidget(desc_label)
        desc_group.setLayout(desc_layout)
        main_layout.addWidget(desc_group)

        # 1. 폴더 선택 섹션
        dir_group = QGroupBox("엑셀 파일 폴더 선택")
        dir_layout = QHBoxLayout()
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("폴더 선택 버튼을 눌러 엑셀 폴더를 지정하세요")
        self.btn_select_dir = QPushButton("폴더 선택")
        dir_layout.addWidget(self.path_input)
        dir_layout.addWidget(self.btn_select_dir)
        dir_group.setLayout(dir_layout)
        main_layout.addWidget(dir_group)

        # 2. 파일별 1대1 열 번호 매칭 입력 섹션 (동적 스크롤 영역)
        self.file_col_group = QGroupBox("파일별 추출할 엑셀 열 번호 설정 (1:1 매칭)")
        self.file_col_layout = QVBoxLayout()
        
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        
        self.file_col_layout.addWidget(self.scroll_area)
        self.file_col_group.setLayout(self.file_col_layout)
        main_layout.addWidget(self.file_col_group)

        # 초기 안내 문구 설정
        self.show_empty_message("폴더를 선택하면 파일별 열 번호 입력창이 여기에 표시됩니다.")

        # [추가] 추출 설정 섹션 (min_length 조절)
        setting_group = QGroupBox("추출 설정")
        setting_layout = QHBoxLayout()
        
        lbl_min_length = QLabel("이 글자 수 미만의 응답은 JSON 추출 시 제외됩니다. :")
        self.spin_min_length = QSpinBox()
        self.spin_min_length.setRange(1, 100)  # 최소 1자부터 100자까지 조절 가능
        self.spin_min_length.setValue(30)       # 기본값 30
        self.spin_min_length.setToolTip("이 글자 수 미만의 응답은 JSON 추출 시 제외됩니다.")
        
        setting_layout.addWidget(lbl_min_length)
        setting_layout.addWidget(self.spin_min_length)
        setting_group.setLayout(setting_layout)
        main_layout.addWidget(setting_group)
        
        # 3. 변환 작업 버튼 섹션
        btn_layout = QHBoxLayout()
        self.btn_excel_to_json = QPushButton("1. 엑셀 -> JSON 일괄 변환 (20KB 분할)")
        self.btn_json_to_word = QPushButton("2. JSON -> Word 변환 (병합 및 _word 생성)")
        btn_layout.addWidget(self.btn_excel_to_json)
        btn_layout.addWidget(self.btn_json_to_word)
        main_layout.addLayout(btn_layout)

        # 4. 로그 및 상태 출력 창
        log_group = QGroupBox("작업 로그")
        log_layout = QVBoxLayout()
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        log_layout.addWidget(self.log_view)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

    def show_empty_message(self, message):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item is None:
                continue
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.file_inputs.clear()
        lbl = QLabel(message)
        self.scroll_layout.addWidget(lbl)

    def create_file_column_inputs(self, excel_files):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item is None:
                continue
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        self.file_inputs.clear()

        if not excel_files:
            self.show_empty_message("선택한 폴더에 엑셀(.xlsx) 파일이 없습니다.")
            return

        for f in excel_files:
            row_layout = QHBoxLayout()
            lbl_name = QLabel(f.name)
            lbl_name.setFixedWidth(220)
            
            line_edit = QLineEdit()
            line_edit.setPlaceholderText("예: 6, 8, 10, 12, 14, 16, 21")
            
            row_layout.addWidget(lbl_name)
            row_layout.addWidget(line_edit)
            
            container = QWidget()
            container.setLayout(row_layout)
            self.scroll_layout.addWidget(container)
            
            self.file_inputs[f] = line_edit

    def log(self, message: str):
        self.log_view.append(message)