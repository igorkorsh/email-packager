import sys
from pathlib import Path
from enum import Enum
from typing import Optional
from PyQt5.QtCore import QObject, pyqtSignal, Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLineEdit, QGroupBox, QHBoxLayout, QPushButton, QCheckBox, QFileDialog, QTreeWidget, QTreeWidgetItem, QProgressBar, QStatusBar
from PyQt5.QtGui import QIcon, QKeyEvent
from email_packager import EmailPackager
from constants import OPTIONS

class ProcessStatus(Enum):
    READY = "Ready"
    NO_INPUT_DIR = "No input directory"
    NO_OUTPUT_DIR = "No output directory"
    NO_HTML_FILES = "HTML files not found"
    IN_PROGRESS = "Processing..."
    COMPLETED = "Completed"
    ERROR = "Completed with errors. See error.log"


class StatusHandler(QObject):
    status_changed = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.status: Optional[ProcessStatus] = None
        
    def set_status(self, status: ProcessStatus):
        self.status = status
        self.status_changed.emit(status.value)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.app = EmailPackager()
        self.status_handler = StatusHandler()
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Email Packager 1.1.1")
        self.setMinimumSize(600, 400)
        widget = QWidget()
        layout = QVBoxLayout(widget)
        self.setCentralWidget(widget)
        
        self.setup_directory(layout)
        self.setup_options(layout)
        self.setup_table(layout)
        self.setup_progress(layout)
        self.setup_button(layout)
        self.setup_status()

    # Выбор папки
    def setup_directory(self, layout: QVBoxLayout):
        self.input_dir = QLineEdit()
        self.input_dir.setReadOnly(True)
        input_group = QGroupBox("Input directory")
        input_layout = QHBoxLayout(input_group)
        input_button = QPushButton("Select")
        input_button.clicked.connect(self.select_input)
        input_layout.addWidget(self.input_dir)
        input_layout.addWidget(input_button)
        layout.addWidget(input_group)
        
        self.output_dir = QLineEdit()
        self.output_dir.textChanged.connect(self.on_output_dir_changed)
        output_group = QGroupBox("Output directory")
        output_layout = QHBoxLayout(output_group)
        output_button = QPushButton("Select")
        output_button.clicked.connect(self.select_output)
        output_layout.addWidget(self.output_dir)
        output_layout.addWidget(output_button)
        layout.addWidget(output_group)

    def select_input(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Select input directory")
        if not directory:
            self.status_handler.set_status(ProcessStatus.NO_INPUT_DIR)
            return

        self.input_dir.setText(directory)
        self.app.input_dir = Path(directory)
        self.update_table()
        if len(self.app.data) == 0:
            self.status_handler.set_status(ProcessStatus.NO_HTML_FILES)
            return
        
        if not self.output_dir.text():
            output = (Path(directory) / 'out').as_posix()
            self.output_dir.setText(output)
            self.app.output_dir = Path(output)
        self.status_handler.set_status(ProcessStatus.READY)

    def select_output(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Select output directory")
        self.output_dir.setText(directory)
        self.app.output_dir = Path(directory)

    def on_output_dir_changed(self, text: str) -> None:
        if not text:
            self.button.setDisabled(True)
            self.status_handler.set_status(ProcessStatus.NO_OUTPUT_DIR)
        else:
            self.app.output_dir = Path(text)
            self.button.setDisabled(len(self.app.data) == 0)
            if not len(self.app.data) == 0:
                self.status_handler.set_status(ProcessStatus.READY) 

    # Опции
    def setup_options(self, layout: QVBoxLayout):
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout(options_group)

        for option, label in OPTIONS.items():
            checkbox = QCheckBox(label)
            checkbox.setChecked(self.app.options[option])
            checkbox.stateChanged.connect(lambda state, opt=option: self.option_changed(opt, state))
            options_layout.addWidget(checkbox)
        
        layout.addWidget(options_group)

    def option_changed(self, option: str, state: int) -> None:
        self.app.options[option] = bool(state)

    # Данные
    def setup_table(self, layout: QVBoxLayout):
        self.table = QTreeWidget()
        self.table.setHeaderLabels(["Filename", "Subject", "OFT", "EML", "PDF"])
        self.table.setAlternatingRowColors(True)
        self.table.setIndentation(0)
        # self.table.setStyleSheet("QTreeWidget::item { padding-top: 4px; padding-bottom: 4px; }")
        self.table.setColumnWidth(2, 50)
        self.table.setColumnWidth(3, 50)
        self.table.setColumnWidth(4, 50)
        self.table.header().setStretchLastSection(False)
        self.table.header().setSectionResizeMode(0, self.table.header().Stretch)
        self.table.header().setSectionResizeMode(1, self.table.header().Stretch)
        layout.addWidget(self.table)

    def update_table(self) -> None:
        self.table.clear()
        self.app.data.clear()
        self.app.get_data()
        self.button.setDisabled(len(self.app.data) == 0)
        self.table.itemChanged.connect(self.format_changed)

        for email in self.app.data:
            item = QTreeWidgetItem([email["path"].name, email["subject"]])
            item.setCheckState(2, Qt.Checked if email["is_oft"] else Qt.Unchecked)
            item.setCheckState(3, Qt.Checked if email["is_eml"] else Qt.Unchecked)
            item.setCheckState(4, Qt.Checked if email["is_pdf"] else Qt.Unchecked)
            item.setData(0, Qt.UserRole, email)
            self.table.addTopLevelItem(item)

    def format_changed(self, item: QTreeWidgetItem, column: int) -> None:
        row = self.table.indexOfTopLevelItem(item)
        is_checked = item.checkState(column) == Qt.Checked
         
        if column == 2:
            self.app.data[row]["is_oft"] = is_checked
        elif column == 3:
            self.app.data[row]["is_eml"] = is_checked
        elif column == 4:
            self.app.data[row]["is_pdf"] = is_checked

    def setup_progress(self, layout: QVBoxLayout):
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress)

    def setup_button(self, layout: QVBoxLayout):
        self.button = QPushButton("Package")
        self.button.clicked.connect(self.package)
        self.button.setDisabled(True)
        layout.addWidget(self.button)

    # Строка состояния
    def setup_status(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_handler.status_changed.connect(self.status_bar.showMessage)
        self.status_handler.set_status(ProcessStatus.READY)

    def package(self) -> None:
        self.status_handler.set_status(ProcessStatus.IN_PROGRESS)
        self.progress.setValue(0)
        self.app.output_dir.mkdir(exist_ok=True)
        
        for i, email in enumerate(self.app.data):
            self.app.process(email)
            self.progress.setValue(int((i + 1) / len(self.app.data) * 100))
            
        if self.app.error_handler.errors:
            self.status_handler.set_status(ProcessStatus.ERROR)
        else:
            self.status_handler.set_status(ProcessStatus.COMPLETED)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("src/package.ico"))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())