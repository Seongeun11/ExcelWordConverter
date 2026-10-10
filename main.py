#main.py
#pyinstaller --noconsole --onefile --hidden-import=pandas --hidden-import=docx --hidden-import=PySide6 --name="ExcelWordConverter" main.py
import sys
from PySide6.QtWidgets import QApplication
from model import ExcelModel
from view import ExcelWordView
from controller import ExcelWordController

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    model = ExcelModel()
    view = ExcelWordView()
    controller = ExcelWordController(model, view)
    
    view.show()
    sys.exit(app.exec())