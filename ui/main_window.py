from __future__ import annotations

import traceback
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal, Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from services.excel_service import ExcelService, ProcessResult
from utils.logger import get_logger


class Worker(QObject):
    finished = Signal(object)
    failed = Signal(str)
    log = Signal(str)

    def __init__(self, service: ExcelService, bank_file: str, template_file: str, output_dir: str):
        super().__init__()
        self.service = service
        self.bank_file = bank_file
        self.template_file = template_file
        self.output_dir = output_dir

    def run(self):
        try:
            result = self.service.process(self.bank_file, self.template_file, self.output_dir, self.log.emit)
            self.finished.emit(result)
        except Exception as exc:
            self.service.logger.exception("处理失败")
            self.log.emit(traceback.format_exc())
            self.failed.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.logger = get_logger()
        self.excel_service = ExcelService(self.logger)
        self.thread = None
        self.worker = None
        self.output_file = ""

        self.setWindowTitle("银行流水自动生成手工日记账工具")
        self.resize(980, 680)
        self._center_window()
        self._init_ui()

    def _init_ui(self):
        root = QWidget()
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        title = QLabel("银行流水自动生成手工日记账工具")
        title.setFont(QFont("Microsoft YaHei", 15, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        file_group = QGroupBox("文件选择")
        file_layout = QGridLayout(file_group)
        file_layout.setHorizontalSpacing(10)
        file_layout.setVerticalSpacing(10)

        self.bank_input = QLineEdit()
        self.template_input = QLineEdit()
        self.output_input = QLineEdit()

        self.bank_input.setPlaceholderText("请选择银行流水文件（.xls/.xlsx）")
        self.template_input.setPlaceholderText("请选择模板文件（.xlsx）")
        self.output_input.setPlaceholderText("可选，默认输出到模板同目录")

        file_layout.addWidget(QLabel("银行流水文件"), 0, 0)
        file_layout.addWidget(self.bank_input, 0, 1)
        bank_btn = QPushButton("浏览")
        bank_btn.clicked.connect(self.select_bank_file)
        file_layout.addWidget(bank_btn, 0, 2)

        file_layout.addWidget(QLabel("模板文件"), 1, 0)
        file_layout.addWidget(self.template_input, 1, 1)
        tpl_btn = QPushButton("浏览")
        tpl_btn.clicked.connect(self.select_template_file)
        file_layout.addWidget(tpl_btn, 1, 2)

        file_layout.addWidget(QLabel("输出目录"), 2, 0)
        file_layout.addWidget(self.output_input, 2, 1)
        out_btn = QPushButton("浏览")
        out_btn.clicked.connect(self.select_output_dir)
        file_layout.addWidget(out_btn, 2, 2)
        layout.addWidget(file_group)

        ops_group = QGroupBox("操作")
        ops_layout = QHBoxLayout(ops_group)
        self.start_btn = QPushButton("开始处理")
        self.reset_btn = QPushButton("重置")
        self.exit_btn = QPushButton("退出")
        self.start_btn.clicked.connect(self.start_process)
        self.reset_btn.clicked.connect(self.reset_form)
        self.exit_btn.clicked.connect(self.close)
        ops_layout.addWidget(self.start_btn)
        ops_layout.addWidget(self.reset_btn)
        ops_layout.addWidget(self.exit_btn)
        ops_layout.addStretch(1)
        layout.addWidget(ops_group)

        status_group = QGroupBox("状态日志")
        status_layout = QVBoxLayout(status_group)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        status_layout.addWidget(self.log_box)
        layout.addWidget(status_group, 1)

        result_group = QGroupBox("结果")
        result_layout = QGridLayout(result_group)
        self.result_input = QLineEdit()
        self.result_input.setReadOnly(True)
        self.open_dir_btn = QPushButton("打开输出目录")
        self.open_dir_btn.clicked.connect(self.open_output_dir)
        result_layout.addWidget(QLabel("输出文件"), 0, 0)
        result_layout.addWidget(self.result_input, 0, 1)
        result_layout.addWidget(self.open_dir_btn, 0, 2)
        layout.addWidget(result_group)

        self.setStyleSheet(
            """
            QMainWindow { background: #f6f8fb; }
            QGroupBox { font-weight: bold; border: 1px solid #d9dce3; border-radius: 8px; margin-top: 8px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QLineEdit, QTextEdit { background: white; border: 1px solid #cfd5df; border-radius: 6px; padding: 6px; }
            QPushButton { background: #2377ff; color: white; border: 0; border-radius: 6px; padding: 8px 14px; }
            QPushButton:hover { background: #1b63d8; }
            QPushButton:disabled { background: #9bb7ea; }
            """
        )

    def _center_window(self):
        screen = self.screen().availableGeometry()
        geo = self.frameGeometry()
        geo.moveCenter(screen.center())
        self.move(geo.topLeft())

    def append_log(self, text: str):
        self.log_box.append(text)

    def select_bank_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择银行流水", "", "Excel 文件 (*.xls *.xlsx)")
        if path:
            self.bank_input.setText(path)
            self.append_log("已选择银行流水文件")

    def select_template_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择模板文件", "", "Excel 文件 (*.xlsx)")
        if path:
            self.template_input.setText(path)
            if not self.output_input.text().strip():
                self.output_input.setText(str(Path(path).parent))
            self.append_log("已选择模板文件")

    def select_output_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.output_input.setText(path)

    def reset_form(self):
        self.bank_input.clear()
        self.template_input.clear()
        self.output_input.clear()
        self.result_input.clear()
        self.output_file = ""
        self.log_box.clear()

    def _validate_input(self) -> bool:
        if not self.bank_input.text().strip():
            QMessageBox.warning(self, "提示", "未选择银行流水文件")
            return False
        if not self.template_input.text().strip():
            QMessageBox.warning(self, "提示", "未选择模板文件")
            return False
        if not self.output_input.text().strip():
            self.output_input.setText(str(Path(self.template_input.text().strip()).parent))
        return True

    def start_process(self):
        if not self._validate_input():
            return

        self.start_btn.setEnabled(False)
        self.append_log("正在处理，请稍候…")

        self.thread = QThread(self)
        self.worker = Worker(
            self.excel_service,
            self.bank_input.text().strip(),
            self.template_input.text().strip(),
            self.output_input.text().strip(),
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self._on_success)
        self.worker.failed.connect(self._on_failure)

        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def _on_success(self, result: ProcessResult):
        self.start_btn.setEnabled(True)
        self.output_file = result.output_path
        self.result_input.setText(result.output_path)
        QMessageBox.information(self, "处理成功", f"处理完成，已写入 {result.rows_written} 行。")

    def _on_failure(self, msg: str):
        self.start_btn.setEnabled(True)
        QMessageBox.critical(self, "处理失败", f"处理失败：{msg}")

    def open_output_dir(self):
        from PySide6.QtGui import QDesktopServices
        from PySide6.QtCore import QUrl

        path = self.output_file or self.output_input.text().strip()
        if not path:
            QMessageBox.information(self, "提示", "暂无可打开的输出目录")
            return
        target = Path(path)
        if target.is_file():
            target = target.parent
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
