"""
Tester Tool - GUI 界面模块
"""

import os
import subprocess
import platform
from pathlib import Path
from datetime import datetime, timedelta
import re

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QComboBox, QSpinBox,
    QMessageBox, QProgressBar, QGroupBox, QFormLayout, QLineEdit,
    QDoubleSpinBox, QScrollArea, QTableWidget, QTableWidgetItem, QTextEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QDragEnterEvent, QDropEvent

from tester.processor import DataProcessor
from tester.validator import ExcelValidator


class ProcessThread(QThread):
    """数据处理后台线程"""
    
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, str)  # message, output_path
    error = pyqtSignal(str)
    
    def __init__(self, csv_files, excel_file, ambient_cols, 
                 file_index_col, channel_index_col, time_interval, temp_threshold):
        super().__init__()
        self.csv_files = csv_files
        self.excel_file = excel_file
        self.ambient_cols = ambient_cols
        self.file_index_col = file_index_col
        self.channel_index_col = channel_index_col
        self.time_interval = time_interval
        self.temp_threshold = temp_threshold
    
    def run(self):
        try:
            self.progress.emit(10)
            
            processor = DataProcessor()
            self.progress.emit(30)
            
            result = processor.process(
                csv_files=self.csv_files,
                excel_file=self.excel_file,
                ambient_cols=self.ambient_cols,
                file_index_col=self.file_index_col,
                channel_index_col=self.channel_index_col,
                time_interval=self.time_interval,
                temp_threshold=self.temp_threshold
            )
            
            self.progress.emit(90)
            
            output_path = result.get('output_path', '')
            message = result.get('message', '处理完成')
            warnings = result.get('warnings', [])
            
            # 如果有警告，添加到消息中
            if warnings:
                message += "\n\n⚠️ 警告:\n" + "\n".join(warnings)
            
            self.progress.emit(100)
            self.finished.emit(message, output_path)
            
        except Exception as e:
            self.error.emit(str(e))


class TesterApp(QMainWindow):
    """试验数据处理工具主界面"""
    
    def __init__(self):
        super().__init__()
        self.csv_files = []
        self.excel_file = ""
        self.ambient_cols = []  # 所有数据行
        self.process_thread = None
        self.output_dir = ""
        
        self.init_ui()
    
    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle("试验数据处理工具 v1.0")
        self.setMinimumSize(800, 600)
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QFormLayout()
        
        # CSV 文件选择
        csv_layout = QHBoxLayout()
        self.csv_label = QLabel("未选择文件")
        self.csv_label.setStyleSheet("color: #666;")
        btn_csv = QPushButton("选择 CSV 文件")
        btn_csv.clicked.connect(self.select_csv_files)
        csv_layout.addWidget(self.csv_label, 1)
        csv_layout.addWidget(btn_csv)
        file_layout.addRow("CSV 数据:", csv_layout)
        
        # Excel 文件选择
        excel_layout = QHBoxLayout()
        self.excel_label = QLabel("未选择文件")
        self.excel_label.setStyleSheet("color: #666;")
        btn_excel = QPushButton("选择 Excel 文件")
        btn_excel.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(self.excel_label, 1)
        excel_layout.addWidget(btn_excel)
        file_layout.addRow("Excel 模板:", excel_layout)
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # 参数设置区域
        param_group = QGroupBox("参数设置")
        param_layout = QFormLayout()
        
        # 索引列选择
        index_layout = QHBoxLayout()
        
        # 文件索引列选择
        file_index_label = QLabel("文件索引列:")
        self.file_index_combo = QComboBox()
        self.file_index_combo.addItems(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"])
        self.file_index_combo.setCurrentText("D")
        self.file_index_combo.currentTextChanged.connect(self.on_index_column_changed)
        index_layout.addWidget(file_index_label)
        index_layout.addWidget(self.file_index_combo)
        
        # 通道索引列选择
        channel_index_label = QLabel("  通道索引列:")
        self.channel_index_combo = QComboBox()
        self.channel_index_combo.addItems(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"])
        self.channel_index_combo.setCurrentText("E")
        self.channel_index_combo.currentTextChanged.connect(self.on_index_column_changed)
        index_layout.addWidget(channel_index_label)
        index_layout.addWidget(self.channel_index_combo)
        
        param_layout.addRow("", index_layout)
        
        # 时间间隔 - 可输入，单位分钟（单位在右边）
        time_layout = QHBoxLayout()
        self.time_spin = QSpinBox()
        self.time_spin.setRange(1, 480)  # 1分钟到8小时
        self.time_spin.setValue(60)
        # 去掉后缀，用右侧标签显示单位
        time_layout.addWidget(self.time_spin)
        time_layout.addWidget(QLabel("分钟"))
        time_layout.addStretch()
        param_layout.addRow("稳定时间间隔:", time_layout)
        
        # 温差阈值 - 不带单位
        self.threshold_spin = QDoubleSpinBox()
        self.threshold_spin.setRange(0.1, 20.0)
        self.threshold_spin.setSingleStep(0.1)
        self.threshold_spin.setValue(2.0)
        param_layout.addRow("温差阈值:", self.threshold_spin)
        
        param_group.setLayout(param_layout)
        main_layout.addWidget(param_group)
        
        # 环境温度选择 - 从已选择的数据行中选
        self.ambient_group = QGroupBox("环境温度 (勾选作为环境温度的行)")
        ambient_layout = QVBoxLayout()
        self.ambient_table = QTableWidget()
        self.ambient_table.setColumnCount(5)
        self.ambient_table.setHorizontalHeaderLabels(["选择", "B列(名称)", "文件索引", "通道索引", "Limit"])
        self.ambient_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        ambient_layout.addWidget(self.ambient_table)
        self.ambient_group.setLayout(ambient_layout)
        main_layout.addWidget(self.ambient_group)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # 日志/消息显示区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet("""
            background: #1e1e1e; 
            color: #ffffff; 
            font-family: monospace; 
            font-size: 12px;
            padding: 8px;
        """)
        main_layout.addWidget(self.log_text)
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        
        self.btn_process = QPushButton("开始处理")
        self.btn_process.setEnabled(False)
        self.btn_process.clicked.connect(self.start_process)
        btn_layout.addWidget(self.btn_process)
        
        self.btn_open_dir = QPushButton("打开输出目录")
        self.btn_open_dir.setEnabled(False)
        self.btn_open_dir.clicked.connect(self.open_output_dir)
        btn_layout.addWidget(self.btn_open_dir)
        
        main_layout.addLayout(btn_layout)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
    
    def select_csv_files(self):
        """选择 CSV 文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 CSV 文件",
            "",
            "CSV 文件 (*.csv);;所有文件 (*.*)"
        )
        
        if files:
            self.csv_files = files
            self.csv_label.setText(f"已选择 {len(files)} 个文件")
            self.csv_label.setStyleSheet("color: #2e7d32;")
            self.check_ready()
    
    def select_excel_file(self):
        """选择 Excel 文件"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 文件",
            "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        
        if file:
            self.excel_file = file
            self.excel_label.setText(Path(file).name)
            self.excel_label.setStyleSheet("color: #2e7d32;")
            
            # 验证 Excel 文件
            self.validate_excel_file()
            self.check_ready()
    
    def get_time_interval_minutes(self):
        """获取时间间隔（分钟）"""
        return self.time_spin.value()
    
    def validate_excel_file(self):
        """验证 Excel 文件"""
        try:
            # 获取用户选择的索引列
            file_index_col = self.get_file_index_col()
            channel_index_col = self.get_channel_index_col()
            
            validator = ExcelValidator()
            result = validator.validate(
                self.excel_file, 
                file_index_col=file_index_col,
                channel_index_col=channel_index_col
            )
            
            if not result['valid']:
                QMessageBox.critical(
                    self,
                    "文件验证失败",
                    f"错误: {result['message']}"
                )
                self.btn_process.setEnabled(False)
                self.ambient_cols = []
                self.ambient_table.setRowCount(0)
                return False
            
            # 显示环境温度选择
            self.ambient_cols = result.get('ambient_rows', [])
            self.display_ambient_cols()
            
            return True
            
        except Exception as e:
            import traceback
            error_msg = f"验证文件时出错: {str(e)}\n\n详细信息:\n{traceback.format_exc()}"
            QMessageBox.critical(self, "错误", error_msg)
            return False
    
    def display_ambient_cols(self):
        """显示环境温度选择"""
        self.ambient_table.setRowCount(len(self.ambient_cols))
        
        for i, row_info in enumerate(self.ambient_cols):
            # 添加复选框（默认不选中）
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            checkbox.setCheckState(Qt.CheckState.Unchecked)
            self.ambient_table.setItem(i, 0, checkbox)
            # B列名字
            self.ambient_table.setItem(i, 1, QTableWidgetItem(str(row_info.get('b_value', ''))))
            # 文件索引
            self.ambient_table.setItem(i, 2, QTableWidgetItem(str(row_info.get('d_value', ''))))
            # 通道索引
            self.ambient_table.setItem(i, 3, QTableWidgetItem(str(row_info.get('e_value', ''))))
            # Limit列
            self.ambient_table.setItem(i, 4, QTableWidgetItem(str(row_info.get('limit', ''))))
    
    def column_letter_to_index(self, letter: str) -> int:
        """将列字母转换为数字（1-based）"""
        return ord(letter.upper()) - ord('A') + 1
    
    def on_index_column_changed(self):
        """当索引列改变时，重新验证Excel"""
        if self.excel_file:
            self.validate_excel_file()
    
    def get_file_index_col(self) -> int:
        """获取文件索引列号"""
        return self.column_letter_to_index(self.file_index_combo.currentText())
    
    def get_channel_index_col(self) -> int:
        """获取通道索引列号"""
        return self.column_letter_to_index(self.channel_index_combo.currentText())
    
    def check_ready(self):
        """检查是否准备好处理"""
        ready = bool(self.csv_files and self.excel_file)
        self.btn_process.setEnabled(ready)
    
    def start_process(self):
        """开始处理数据"""
        if not self.csv_files or not self.excel_file:
            return
        
        # 获取选中的环境温度行（根据checkbox）
        ambient_rows = []
        for i in range(self.ambient_table.rowCount()):
            item = self.ambient_table.item(i, 0)
            if item and item.checkState() == Qt.CheckState.Checked:
                if i < len(self.ambient_cols):
                    ambient_rows.append(self.ambient_cols[i])
        
        if not ambient_rows:
            QMessageBox.warning(self, "警告", "请选择环境温度列")
            return
        
        time_interval = self.get_time_interval_minutes()
        temp_threshold = self.threshold_spin.value()
        
        # 清空日志区域并显示开始信息
        self.log_text.clear()
        self.log_text.append(f"📋 开始处理...\nCSV文件: {len(self.csv_files)}个\nExcel文件: {Path(self.excel_file).name}\n稳定时间间隔: {time_interval}分钟\n温差阈值: {temp_threshold}°C\n环境温度: {len(ambient_rows)}行")
        
        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.btn_process.setEnabled(False)
        self.statusBar().showMessage("处理中...")
        
        # 启动后台线程
        self.process_thread = ProcessThread(
            csv_files=self.csv_files,
            excel_file=self.excel_file,
            ambient_cols=ambient_rows,
            file_index_col=self.get_file_index_col(),
            channel_index_col=self.get_channel_index_col(),
            time_interval=time_interval,
            temp_threshold=temp_threshold
        )
        
        self.process_thread.progress.connect(self.progress_bar.setValue)
        self.process_thread.finished.connect(self.process_finished)
        self.process_thread.error.connect(self.process_error)
        
        self.process_thread.start()
    
    def process_finished(self, message: str, output_path: str):
        """处理完成"""
        self.progress_bar.setVisible(False)
        self.btn_process.setEnabled(True)
        self.statusBar().showMessage("处理完成")
        
        # 在日志区域显示结果
        self.log_text.append(f"✅ 完成: {message}")
        
        if output_path:
            self.output_dir = str(Path(output_path).parent)
            self.btn_open_dir.setEnabled(True)
        
        QMessageBox.information(self, "完成", message)
    
    def process_error(self, error: str):
        """处理出错"""
        self.progress_bar.setVisible(False)
        self.btn_process.setEnabled(True)
        self.statusBar().showMessage("处理失败")
        
        # 在日志区域显示错误
        self.log_text.append(f"❌ 错误: {error}")
        
        QMessageBox.critical(self, "错误", f"处理失败:\n{error}")
    
    def open_output_dir(self):
        """打开输出目录"""
        if self.output_dir:
            system = platform.system()
            if system == "Windows":
                os.startfile(self.output_dir)
            elif system == "Darwin":
                subprocess.Popen(["open", self.output_dir])
            else:
                subprocess.Popen(["xdg-open", self.output_dir])


def main():
    """GUI 入口"""
    app = QApplication([])
    window = TesterApp()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()
