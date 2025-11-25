import sys
import os
import pyodbc
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QWidget, QPushButton, QLabel, QLineEdit, QComboBox,
    QTableWidget, QTableWidgetItem, QTreeWidget, QTreeWidgetItem,
    QTabWidget, QMessageBox, QFileDialog, QSplitter, QHeaderView,
    QDialog, QFormLayout
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class EditDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("新增行")
        self.columns = columns
        self.values = {}
        self.init_ui()

    def init_ui(self):
        layout = QFormLayout(self)
        self.edit_fields = {}
        for col in self.columns:
            le = QLineEdit(self)
            layout.addRow(col, le)
            self.edit_fields[col] = le

        btn_layout = QHBoxLayout()
        ok = QPushButton("确定")
        cancel = QPushButton("取消")
        ok.clicked.connect(self.on_ok)
        cancel.clicked.connect(self.reject)
        btn_layout.addWidget(ok)
        btn_layout.addWidget(cancel)
        layout.addRow(btn_layout)

    def on_ok(self):
        for col, le in self.edit_fields.items():
            self.values[col] = le.text()
        self.accept()

class MDBEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_connection = None
        self.current_cursor = None
        self.edits = {}  # 存储编辑变更：键 (表名, row, col) -> 新文本
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('MDB 编辑器')
        self.setGeometry(100, 100, 1200, 800)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        splitter = QSplitter(Qt.Horizontal)

        # 左面板 — 连接 & 表结构
        left = QWidget()
        left_layout = QVBoxLayout(left)

        conn_grp = QWidget()
        conn_l = QVBoxLayout(conn_grp)
        file_h = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        browse = QPushButton("浏览")
        browse.clicked.connect(self.browse_mdb_file)
        file_h.addWidget(QLabel("MDB 文件:"))
        file_h.addWidget(self.file_path_edit)
        file_h.addWidget(browse)
        conn_l.addLayout(file_h)

        self.connect_btn = QPushButton("连接数据库")
        self.connect_btn.clicked.connect(self.connect_to_mdb)
        conn_l.addWidget(self.connect_btn)
        left_layout.addWidget(conn_grp)

        left_layout.addWidget(QLabel("数据库表:"))
        self.table_tree = QTreeWidget()
        self.table_tree.setHeaderLabel("表结构")
        self.table_tree.itemClicked.connect(self.on_table_selected)
        left_layout.addWidget(self.table_tree)

        splitter.addWidget(left)

        # 右面板 — 标签页内容
        right = QWidget()
        right_layout = QVBoxLayout(right)

        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        right_layout.addWidget(self.tabs)

        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        main_layout.addWidget(splitter)

        self.statusBar().showMessage("准备就绪")

    def browse_mdb_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 MDB/ACCDB 文件", "", "Access 数据库 (*.mdb *.accdb)")
        if path:
            self.file_path_edit.setText(path)

    def connect_to_mdb(self):
        path = self.file_path_edit.text()
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "错误", "请选择有效的 MDB 文件")
            return
        try:
            if self.current_cursor:
                self.current_cursor.close()
            if self.current_connection:
                self.current_connection.close()
            conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path}"
            self.current_connection = pyodbc.connect(conn_str)
            self.current_cursor = self.current_connection.cursor()

            # 获取所有表
            tables = [t.table_name for t in self.current_cursor.tables(tableType='TABLE')]
            self.table_tree.clear()
            for tn in tables:
                ti = QTreeWidgetItem([tn])
                self.table_tree.addTopLevelItem(ti)
                # 列
                for col in self.current_cursor.columns(table=tn):
                    ci = QTreeWidgetItem([f"{col.column_name} ({col.type_name})"])
                    ti.addChild(ci)

            self.statusBar().showMessage(f"已连接: {os.path.basename(path)}")

        except Exception as e:
            QMessageBox.critical(self, "连接失败", str(e))
            self.statusBar().showMessage("连接失败")

    def on_table_selected(self, item, col):
        if item.parent() is None:
            table_name = item.text(0)
            self.open_table_tab(table_name)

    def open_table_tab(self, table_name):
        # 如果已经打开 tab，就切换
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == table_name:
                self.tabs.setCurrentIndex(i)
                return

        try:
            # 读出 dataframe
            df = pd.read_sql(f"SELECT * FROM [{table_name}]", self.current_connection)

            table_w = QTableWidget()
            table_w.setRowCount(len(df))
            table_w.setColumnCount(len(df.columns))
            table_w.setHorizontalHeaderLabels(df.columns)

            # 填数据 + 使可编辑
            for r, row in df.iterrows():
                for c, col in enumerate(df.columns):
                    val = row[col]
                    item = QTableWidgetItem("" if pd.isna(val) else str(val))
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    table_w.setItem(r, c, item)

            table_w.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

            # 连接信号
            table_w.itemChanged.connect(lambda item, tn=table_name, tw=table_w: self.on_cell_changed(item, tn, tw))

            # 保存、插入、删除按钮
            btn_widget = QWidget()
            btn_l = QHBoxLayout(btn_widget)
            btn_save = QPushButton("保存修改")
            btn_new = QPushButton("新增行")
            btn_del = QPushButton("删除行")
            btn_l.addWidget(btn_save)
            btn_l.addWidget(btn_new)
            btn_l.addWidget(btn_del)

            btn_save.clicked.connect(lambda _, tn=table_name, tw=table_w: self.save_changes(tn, tw))
            btn_new.clicked.connect(lambda _, tn=table_name, tw=table_w: self.insert_row(tn, tw))
            btn_del.clicked.connect(lambda _, tn=table_name, tw=table_w: self.delete_row(tn, tw))

            # 容器
            container = QWidget()
            v = QVBoxLayout(container)
            v.addWidget(table_w)
            v.addWidget(btn_widget)

            self.tabs.addTab(container, table_name)
            self.tabs.setCurrentWidget(container)
            self.statusBar().showMessage(f"加载表 {table_name}，共 {len(df)} 行")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载表失败: {e}")

    def on_cell_changed(self, item: QTableWidgetItem, table_name, table_widget):
        # 记录变更
        r = item.row()
        c = item.column()
        new = item.text()
        self.edits[(table_name, r, c)] = new

    def save_changes(self, table_name, table_widget):
        if not self.edits:
            QMessageBox.information(self, "提示", "没有修改要保存")
            return

        cursor = self.current_connection.cursor()
        try:
            # 假设第一列是主键
            col_names = [table_widget.horizontalHeaderItem(c).text() for c in range(table_widget.columnCount())]

            # 从数据库 / GUI 里重新读取 DataFrame，以判断唯一列
            df = pd.read_sql(f"SELECT * FROM [{table_name}]", self.current_connection)

            # 找出所有唯一（无重复值）的列
            unique_cols = [c for c in col_names if c in df.columns and df[c].is_unique]

            if unique_cols:
                pk_col = unique_cols[0]  # 取第一个唯一列
            else:
                # 如果没有唯一列，就退回你之前的逻辑（比如第一列）
                pk_col = col_names[0]
                QMessageBox.warning(self,
                        "警告 — 未检测到唯一列",
                        f"表 “{table_name}” 未发现所有值都不同的列，已退回使用第一列 “{pk_col}” 作为主键定位，可能导致更新不准确。")
            print("主键：",pk_col)
            pk_index = col_names.index(pk_col)
            for (tn, r, c), new_value in list(self.edits.items()):
                if tn != table_name:
                    continue
                col = col_names[c]
                pk_item = table_widget.item(r, pk_index)
                if pk_item is None:
                    continue
                pk_value = pk_item.text()
                sql = f"UPDATE [{table_name}] SET [{col}] = ? WHERE [{pk_col}] = ?"
                print("SQL：，new,pk",sql, (new_value, pk_value))
                cursor.execute(sql, (new_value, pk_value))

            self.current_connection.commit()
            QMessageBox.information(self, "成功", "修改已保存")
            self.edits.clear()
            # 重新加载 tab
            self.reload_table_tab(table_name)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))
        finally:
            cursor.close()

    def insert_row(self, table_name, table_widget):
        # 弹对话框让用户输入值
        col_names = [table_widget.horizontalHeaderItem(c).text() for c in range(table_widget.columnCount())]
        dlg = EditDialog(col_names, self)
        if dlg.exec_() == QDialog.Accepted:
            vals = dlg.values  # dict col -> str

            # 构造 INSERT
            cols = ", ".join(f"[{c}]" for c in col_names)
            qmarks = ", ".join("?" for _ in col_names)
            sql = f"INSERT INTO [{table_name}] ({cols}) VALUES ({qmarks})"
            params = [vals[c] for c in col_names]

            try:
                cursor = self.current_connection.cursor()
                cursor.execute(sql, params)
                self.current_connection.commit()
                cursor.close()
                QMessageBox.information(self, "插入成功", "新行已插入")
                self.reload_table_tab(table_name)
            except Exception as e:
                QMessageBox.critical(self, "插入失败", str(e))

    def delete_row(self, table_name, table_widget):
        selected = table_widget.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "提示", "请先选中一行")
            return

        col_names = [table_widget.horizontalHeaderItem(c).text() for c in range(table_widget.columnCount())]
        pk_col = col_names[0]
        pk_item = table_widget.item(selected, 0)
        if pk_item is None:
            QMessageBox.warning(self, "错误", "无法获取主键值")
            return
        pk_value = pk_item.text()

        confirm = QMessageBox.question(self, "确认删除", f"确定删除主键 = {pk_value} 的行吗？")
        if confirm != QMessageBox.Yes:
            return

        try:
            cursor = self.current_connection.cursor()
            sql = f"DELETE FROM [{table_name}] WHERE [{pk_col}] = ?"
            cursor.execute(sql, (pk_value,))
            self.current_connection.commit()
            cursor.close()
            QMessageBox.information(self, "删除成功", "行已删除")
            self.reload_table_tab(table_name)
        except Exception as e:
            QMessageBox.critical(self, "删除失败", str(e))

    def reload_table_tab(self, table_name):
        # 找到 tab index，然后关闭 + 重新 open
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == table_name:
                self.tabs.removeTab(i)
                break
        self.open_table_tab(table_name)

    def close_tab(self, idx):
        self.tabs.removeTab(idx)

    def closeEvent(self, event):
        if self.current_cursor:
            self.current_cursor.close()
        if self.current_connection:
            self.current_connection.close()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setFont(QFont("Arial", 10))
    w = MDBEditor()
    w.show()
    sys.exit(app.exec_())
