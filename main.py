# encoding=utf-8

import sys
import os
from PyQt5 import QtWidgets
from process import main as process
from process import get_except_keywords, set_except_keywords


class AppWindow(QtWidgets.QMainWindow):
    KEYWORDS_DUMP_FILE = 'keywords.dat'

    def __init__(self, parent=None):
        super(AppWindow, self).__init__(parent)
        if os.path.exists(self.KEYWORDS_DUMP_FILE):
            with open(self.KEYWORDS_DUMP_FILE, 'rb') as f:
                set_except_keywords(f.read())
        self.resize(500, 500)
        self.create_main_ui()

    # UI definition
    def create_main_ui(self):
        self.main_frame = QtWidgets.QWidget()
        setting_layout = self.create_setting_layout()
        file_list_layout = self.create_file_list_layout()
        output_layout = self.create_output_layout()
        v_layout = QtWidgets.QVBoxLayout()
        v_layout.addLayout(setting_layout)
        v_layout.addLayout(file_list_layout)
        v_layout.addLayout(output_layout)
        self.main_frame.setLayout(v_layout)
        self.setCentralWidget(self.main_frame)

    def create_setting_layout(self):
        self.keywords_input = QtWidgets.QTextEdit()
        self.keywords_input.setText(get_except_keywords())
        tip_text = QtWidgets.QLabel()
        tip_text.setText(u'非清华机构关键词（用于排除结果中的非清华机构，不区分大小写，每行一个）')
        v_layout = QtWidgets.QVBoxLayout()
        v_layout.addWidget(tip_text)
        v_layout.addWidget(self.keywords_input)
        return v_layout

    def create_file_list_layout(self):
        self.file_list = QtWidgets.QListWidget()
        self.file_list.resize(300, 400)
        self.add_file_button = QtWidgets.QPushButton('添加文件')
        self.remove_file_button = QtWidgets.QPushButton('移除文件')
        self.add_file_button.clicked.connect(self.on_add_file)
        self.remove_file_button.clicked.connect(self.on_remove_files)
        self.tip_text = QtWidgets.QLabel()
        self.tip_text.setText('已选择0个文件')
        h_button_layout = QtWidgets.QHBoxLayout()
        h_button_layout.addStretch()
        h_button_layout.addWidget(self.add_file_button)
        h_button_layout.addWidget(self.remove_file_button)
        v_layout = QtWidgets.QVBoxLayout()
        v_layout.addLayout(h_button_layout)
        v_layout.addWidget(self.file_list)
        v_layout.addWidget(self.tip_text)
        return v_layout

    def create_output_layout(self):
        self.save_button = QtWidgets.QPushButton('导出结果')
        self.save_button.clicked.connect(self.on_save)
        h_layout = QtWidgets.QHBoxLayout()
        h_layout.addStretch()
        h_layout.addWidget(self.save_button)
        return h_layout

    # Event handler
    def on_add_file(self):
        path_list, _ = QtWidgets.QFileDialog.getOpenFileNames(self, '选取文件', '', 'highly*.*;hot*.*')
        if not path_list:
            return
        for path in path_list:
            self.file_list.addItem(path)
        self.update_selected_count()

    def on_remove_files(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
        self.update_selected_count()

    def update_selected_count(self):
        self.tip_text.setText('已选择%d个文件' % self.file_list.count())

    def on_save(self):
        set_except_keywords(self.keywords_input.toPlainText())
        save_filename, _ = QtWidgets.QFileDialog.getSaveFileName(self, '导出结果', '', '')
        filenames = list()
        for i in xrange(self.file_list.count()):
            filename = self.file_list.item(i).text()
            filenames.append(filename)
        tsv, xls, dup = process(filenames, save_filename)
        with open(self.KEYWORDS_DUMP_FILE, 'wb') as f:
            f.write(get_except_keywords())
        message = u'''
            tsv格式结果文件：{tsv}\r\n
            excel 格式结果文件：{xls}\r\n
            按院系去重结果文件：{dup}
        '''.format(tsv=tsv, xls=xls, dup=dup)
        self.add_message_box(message, u'结果导出成功')

    def add_message_box(self, message, title):
        msg = QtWidgets.QMessageBox()
        msg.setText(message)
        msg.setWindowTitle(title)
        msg.exec_()


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = AppWindow()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
