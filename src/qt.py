import sys
from packager import EmailPackager
from PyQt5.QtWidgets import QApplication, QMainWindow, QStatusBar, QWidget, QVBoxLayout, QLineEdit, \
    QPushButton, QFileDialog, QTableWidget, QCheckBox, QTableWidgetItem


class Window(QMainWindow):
    VERSION = "1.0.0"

    def __init__(self):
        super(Window, self).__init__()

        self.app = EmailPackager()
        self.files = []
        self.selected_elements = set()
        self.setWindowTitle(f"Email Packager {self.VERSION}")
        self.setMinimumSize(350, 400)

        self.widget = QWidget(self)
        self.setCentralWidget(self.widget)
        self.layout = QVBoxLayout(self.widget)

        # Создает строку состояния
        self.status = QStatusBar()
        self.setStatusBar(self.status)

        # Создает кнопку выбора папки
        self.button = QPushButton("Выбрать папку")
        self.button.clicked.connect(self.choose_dir)
        self.layout.addWidget(self.button)

        # Создает таблицу
        self.table = QTableWidget()
        self.setup_table()
        self.layout.addWidget(self.table)

        self.output = QLineEdit()
        self.output.setPlaceholderText("out")
        self.output.textChanged.connect(self.on_output_changed)
        self.layout.addWidget(self.output)

        self.package_button = QPushButton("Пересобрать")
        self.package_button.clicked.connect(self.package)
        self.package_button.setDisabled(True)
        self.layout.addWidget(self.package_button)

        self.delete_button = QPushButton("Удалить")
        self.delete_button.clicked.connect(self.remove_items)
        self.delete_button.setDisabled(True)
        self.layout.addWidget(self.delete_button)

    # Настраивает таблицу
    def setup_table(self):
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["", "Имя файла"])
        self.table.setColumnWidth(0, 0)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionMode(self.table.NoSelection)

    def choose_dir(self):
        self.status.clearMessage()
        path = QFileDialog.getExistingDirectory(self)

        if path:
            self.files = self.app.get_files(path)
            self.update_table()
        else:
            self.status.showMessage("Папка не выбрана")

    def update_table(self):
        self.table.setRowCount(0)

        if len(self.files) > 0:
            self.set_package_button_state()
            self.table.setRowCount(len(self.files))

            for row, file_path in enumerate(self.files):
                checkbox = QCheckBox()
                checkbox.stateChanged.connect(lambda state, index=row: self.set_delete_button_state(state, index))
                self.table.setCellWidget(row, 0, checkbox)
                self.table.setItem(row, 1, QTableWidgetItem(file_path.name))
        else:
            self.status.showMessage("В этой папке нет html-файлов")

    def set_package_button_state(self):
        self.package_button.setDisabled(not bool(len(self.files)))

    def set_delete_button_state(self, state: int, row: int):
        self.selected_elements.add(row) if state == 2 else self.selected_elements.discard(row)
        self.delete_button.setDisabled(not bool(len(self.selected_elements)))

    def remove_items(self):
        self.files = [file for i, file in enumerate(self.files) if i not in self.selected_elements]

        for row in sorted(self.selected_elements, reverse=True):
            self.table.removeRow(row)

        self.selected_elements.clear()
        self.delete_button.setDisabled(True)

    def on_output_changed(self):
        self.app.output = self.output.text() or "out"

    def package(self):
        for file in self.files:
            self.app.package(file)

        self.status.showMessage("Задача выполнена")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())