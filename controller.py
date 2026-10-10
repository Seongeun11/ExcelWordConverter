#controller.py
from pathlib import Path
from PySide6.QtWidgets import QFileDialog, QMessageBox

class ExcelWordController:
    def __init__(self, model, view):
        self.model = model
        self.view = view
        self.selected_dir = None
        self.excel_files = []

        self.connect_signals()

    def connect_signals(self):
        self.view.btn_select_dir.clicked.connect(self.select_folder)
        self.view.btn_excel_to_json.clicked.connect(self.run_excel_to_json)
        self.view.btn_json_to_word.clicked.connect(self.run_json_to_word)

    def select_folder(self):
        dir_path = QFileDialog.getExistingDirectory(self.view, "엑셀 폴더 선택", "")
        if dir_path:
            self.selected_dir = Path(dir_path)
            self.view.path_input.setText(str(self.selected_dir))
            
            self.excel_files = sorted(list(self.selected_dir.glob("*.xlsx")))
            self.view.log(f"📁 폴더 선택됨: {self.selected_dir}")
            self.view.log(f"   └─ 필터링된 엑셀 파일 수: {len(self.excel_files)}개")
            
            self.view.create_file_column_inputs(self.excel_files)
            for f in self.excel_files:
                self.view.log(f"      - 매칭 슬롯 생성: {f.name}")

    def run_excel_to_json(self):
        if not self.selected_dir or not self.selected_dir.exists():
            QMessageBox.warning(self.view, "경고", "먼저 엑셀 폴더를 선택해주세요.")
            return

        if not self.excel_files:
            QMessageBox.warning(self.view, "경고", "선택한 폴더에 처리할 엑셀(.xlsx) 파일이 없습니다.")
            return

        # View에서 사용자가 설정한 min_length 값 가져오기
        min_len = self.view.spin_min_length.value()

        file_target_map = {}
        for f, line_edit in self.view.file_inputs.items():
            text = line_edit.text().strip()
            if not text:
                QMessageBox.warning(self.view, "경고", f"[{f.name}] 파일의 열 번호가 입력되지 않았습니다.")
                return
            try:
                cols = [int(x.strip()) for x in text.split(",") if x.strip().isdigit()]
                if not cols:
                    raise ValueError
                file_target_map[f] = cols
            except ValueError:
                QMessageBox.warning(self.view, "경고", f"[{f.name}] 파일의 열 번호 형식이 잘못되었습니다. (예: 6,8,10)")
                return

        self.view.log("\n🚀 엑셀 -> JSON 변환 시작 (파일별 1:1 매칭 열 적용, 20KB 분할)")
        json_output_root = self.selected_dir.parent / f"{self.selected_dir.name}_json_output"

        for excel_file, target_cols in file_target_map.items():
            self.view.log(f"  └─ 처리 중: {excel_file.name} (적용 열: {target_cols})")
            file_json_dir = json_output_root / excel_file.stem
            self.model.process_excel_to_json(excel_file, target_cols, file_json_dir, min_length=min_len)

        self.view.log("🎉 모든 엑셀 파일의 JSON 변환 및 20KB 분할 저장이 완료되었습니다!")
        QMessageBox.information(self.view, "완료", "엑셀 -> JSON 변환이 완료되었습니다.")

    def run_json_to_word(self):
        if not self.selected_dir or not self.selected_dir.exists():
            QMessageBox.warning(self.view, "경고", "먼저 엑셀 폴더를 선택해주세요.")
            return

        json_output_root = self.selected_dir.parent / f"{self.selected_dir.name}_json_output"
        if not json_output_root.exists():
            QMessageBox.warning(self.view, "경고", "변환된 JSON 폴더가 없습니다. 먼저 1번 변환을 실행해주세요.")
            return

        word_output_root = self.selected_dir.parent / f"{self.selected_dir.name}_word_output"
        
        self.view.log("\n🚀 JSON -> Word 변환 및 분할 파일 병합 시작...")

        # Model의 메서드를 호출하여 그룹화, 정렬 및 병합을 일괄 처리 (MVC 원칙 준수)
        saved_files = self.model.convert_json_to_word_from_directory(
            json_output_root, 
            word_output_root, 
            logger_callback=self.view.log
        )

        if saved_files:
            self.view.log(f"🎉 모든 최종 Word 파일이 [{word_output_root.name}] 폴더에 모아져 저장되었습니다!")
            QMessageBox.information(self.view, "완료", f"Word 변환이 완료되었습니다.\n저장 폴더: {word_output_root.name}")