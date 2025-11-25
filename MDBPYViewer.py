import sys
import os
import pyodbc
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QWidget, QPushButton, QLabel, QLineEdit, QTreeWidget,
    QTreeWidgetItem, QTabWidget, QMessageBox, QFileDialog,
    QSplitter, QTableWidget, QTableWidgetItem, QDialog, QFormLayout,
    QInputDialog
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


class ColumnEditDialog(QDialog):
    """用于整列修改：输入一个新值给某一列"""
    def __init__(self, col_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"修改整列：{col_name}")
        self.value = None
        self.init_ui(col_name)

    def init_ui(self, col_name):
        layout = QFormLayout(self)
        self.le = QLineEdit(self)
        layout.addRow(f"新值 ({col_name})", self.le)
        btn_layout = QHBoxLayout()
        ok = QPushButton("确定")
        cancel = QPushButton("取消")
        ok.clicked.connect(self.on_ok)
        cancel.clicked.connect(self.reject)
        btn_layout.addWidget(ok)
        btn_layout.addWidget(cancel)
        layout.addRow(btn_layout)

    def on_ok(self):
        self.value = self.le.text()
        self.accept()


class BulkEditDialog(QDialog):
    """批量修改对话框 - 可以选择修改方式"""
    def __init__(self, col_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"批量修改列：{col_name}")
        self.col_name = col_name
        self.method = "replace"  # 默认替换方式
        self.value = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # 修改方式选择
        layout.addWidget(QLabel("选择修改方式:"))
        
        self.replace_radio = QPushButton("替换所有值")
        self.replace_radio.setCheckable(True)
        self.replace_radio.setChecked(True)
        self.replace_radio.clicked.connect(lambda: self.set_method("replace"))
        
        self.prefix_radio = QPushButton("添加前缀")
        self.prefix_radio.setCheckable(True)
        self.prefix_radio.clicked.connect(lambda: self.set_method("prefix"))
        
        self.suffix_radio = QPushButton("添加后缀")
        self.suffix_radio.setCheckable(True)
        self.suffix_radio.clicked.connect(lambda: self.set_method("suffix"))
        
        layout.addWidget(self.replace_radio)
        layout.addWidget(self.prefix_radio)
        layout.addWidget(self.suffix_radio)
        
        # 值输入
        self.value_edit = QLineEdit(self)
        layout.addWidget(QLabel("输入值:"))
        layout.addWidget(self.value_edit)
        
        # 按钮
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("确定")
        cancel_btn = QPushButton("取消")
        ok_btn.clicked.connect(self.on_ok)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def set_method(self, method):
        self.method = method
        # 更新按钮状态
        self.replace_radio.setChecked(method == "replace")
        self.prefix_radio.setChecked(method == "prefix")
        self.suffix_radio.setChecked(method == "suffix")

    def on_ok(self):
        self.value = self.value_edit.text()
        if self.value == "" and self.method == "replace":
            QMessageBox.warning(self, "警告", "替换值不能为空！")
            return
        self.accept()


class MDBEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_connection = None
        self.current_cursor = None
        self.edits = {}  # 存储编辑变更：键 (表名, row, col) -> (old_value, new_value)
        self.pk_cache = {}  # 缓存每张表的主键列名字
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

    def get_primary_key(self, table_name):
        """
        尝试获取主键列。如果之前缓存过就用缓存，
        否则通过 ODBC statistics 方法读取 Access 的主键。
        """
        if table_name in self.pk_cache:
            return self.pk_cache[table_name]

        cursor = self.current_connection.cursor()
        pk_cols = {}
        # 使用 statistics 方法获取主键信息（Access ODBC 不支持 cursor.primaryKeys）
        for row in cursor.statistics(table_name):
            # row[5] 是 index type，PrimaryKey 表示主键索引
            if row[5] == "PrimaryKey":
                # row[7] 是序号，row[8] 是列名
                pk_cols[row[7]] = row[8]
        cursor.close()

        if not pk_cols:
            # 如果没找到，就退回到第一列（兜底）
            df = pd.read_sql(f"SELECT * FROM [{table_name}]", self.current_connection)
            pk = df.columns[0]
        else:
            # 按序号排序，拼接所有主键列（支持复合主键）
            pk = [pk_cols[i] for i in sorted(pk_cols.keys())]
            if len(pk) == 1:
                pk = pk[0]

        self.pk_cache[table_name] = pk
        return pk

    def open_table_tab(self, table_name):
        # 如果已经打开 tab，就切换
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == table_name:
                self.tabs.setCurrentIndex(i)
                return

        try:
            df = pd.read_sql(f"SELECT * FROM [{table_name}]", self.current_connection)
            table_w = QTableWidget()
            table_w.setRowCount(len(df))
            table_w.setColumnCount(len(df.columns))
            table_w.setHorizontalHeaderLabels(df.columns)

            # 找出主键列 (可能是列表，也可能是单个)
            pk = self.get_primary_key(table_name)
            # 记录 pk_index（如果是多个主键，这里暂只处理单个主键列情况）
            if isinstance(pk, list):
                # 目前我们只取第一个主键列处理
                pk_col = pk[0]
            else:
                pk_col = pk
            col_names = list(df.columns)
            try:
                pk_index = col_names.index(pk_col)
            except ValueError:
                pk_index = None

            # 填数据 + 设置 flags
            for r, row in df.iterrows():
                for c, col in enumerate(df.columns):
                    val = row[col]
                    item = QTableWidgetItem("" if pd.isna(val) else str(val))
                    # 允许所有列编辑
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    table_w.setItem(r, c, item)

            table_w.horizontalHeader().setSectionResizeMode(table_w.horizontalHeader().ResizeToContents)

            # 连接信号
            table_w.itemChanged.connect(lambda item, tn=table_name, tw=table_w, pk_i=pk_index: self.on_cell_changed(item, tn, tw, pk_i))

            # 保存、插入、删除、整列修改按钮
            btn_widget = QWidget()
            btn_l = QHBoxLayout(btn_widget)
            btn_save = QPushButton("保存修改")
            btn_new = QPushButton("新增行")
            btn_del = QPushButton("删除行")
            btn_bulk_edit = QPushButton("整列修改")
            btn_l.addWidget(btn_save)
            btn_l.addWidget(btn_new)
            btn_l.addWidget(btn_del)
            btn_l.addWidget(btn_bulk_edit)

            btn_save.clicked.connect(lambda _, tn=table_name, tw=table_w: self.save_changes(tn, tw))
            btn_new.clicked.connect(lambda _, tn=table_name, tw=table_w: self.insert_row(tn, tw))
            btn_del.clicked.connect(lambda _, tn=table_name, tw=table_w: self.delete_row(tn, tw))
            btn_bulk_edit.clicked.connect(lambda _, tn=table_name, tw=table_w: self.bulk_edit_column(tn, tw))

            container = QWidget()
            v = QVBoxLayout(container)
            v.addWidget(table_w)
            v.addWidget(btn_widget)

            self.tabs.addTab(container, table_name)
            self.tabs.setCurrentWidget(container)
            self.statusBar().showMessage(f"加载表 {table_name}，共 {len(df)} 行")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载表失败: {e}")

    def on_cell_changed(self, item: QTableWidgetItem, table_name, table_widget, pk_index):
        r = item.row()
        c = item.column()
        new = item.text()
        # 记录变更，同时保存 old 值
        key = (table_name, r, c)
        if key in self.edits:
            old, _ = self.edits[key]
            self.edits[key] = (old, new)
        else:
            # 获取 old 值
            df = pd.read_sql(f"SELECT * FROM [{table_name}]", self.current_connection)
            try:
                old = str(df.iat[r, c])
            except Exception:
                old = ""
            self.edits[key] = (old, new)

    def save_changes(self, table_name, table_widget):
        if not self.edits:
            QMessageBox.information(self, "提示", "没有修改要保存")
            return

        cursor = self.current_connection.cursor()
        
        try:
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
                        f"表 \"{table_name}\" 未发现所有值都不同的列，已退回使用第一列 \"{pk_col}\" 作为主键定位，可能导致更新不准确。")
            
            pk_index = col_names.index(pk_col)
            
            success_count = 0
            error_count = 0
            
            for (tn, r, c), (old_value, new_value) in list(self.edits.items()):
                if tn != table_name:
                    continue
                    
                try:
                    col = col_names[c]
                    # 如果是主键列
                    if c == pk_index:
                        sql = f"UPDATE [{table_name}] SET [{col}] = ? WHERE [{pk_col}] = ?"
                        cursor.execute(sql, (new_value, old_value))
                    else:
                        # 定位行：用原来的 old PK 值
                        pk_item = table_widget.item(r, pk_index)
                        if pk_item is None:
                            continue
                        pk_current = pk_item.text()
                        sql = f"UPDATE [{table_name}] SET [{col}] = ? WHERE [{pk_col}] = ?"
                        cursor.execute(sql, (new_value, pk_current))
                    
                    success_count += 1
                    
                except Exception as e:
                    error_count += 1
                    print(f"更新失败: {e}")

            self.current_connection.commit()
            
            if error_count == 0:
                QMessageBox.information(self, "成功", f"所有修改已保存 ({success_count} 条记录)")
            else:
                QMessageBox.warning(self, "部分成功", 
                                  f"成功保存 {success_count} 条记录，失败 {error_count} 条记录")
                
            self.edits.clear()
            self.reload_table_tab(table_name)
            
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))
            self.current_connection.rollback()
        finally:
            cursor.close()

    def bulk_edit_column(self, table_name, table_widget):
        """整列修改功能"""
        # 获取当前选中的列
        current_col = table_widget.currentColumn()
        if current_col < 0:
            QMessageBox.warning(self, "提示", "请先选中要修改的列")
            return
        
        col_names = [table_widget.horizontalHeaderItem(c).text() for c in range(table_widget.columnCount())]
        col_name = col_names[current_col]
        
        # 弹出批量修改对话框
        dlg = BulkEditDialog(col_name, self)
        if dlg.exec_() == QDialog.Accepted:
            method = dlg.method
            value = dlg.value
            
            if method == "replace" and value == "":
                QMessageBox.warning(self, "警告", "替换值不能为空！")
                return
            
            # 执行批量修改
            self.execute_bulk_edit(table_name, table_widget, current_col, col_name, method, value)

    def execute_bulk_edit(self, table_name, table_widget, col_index, col_name, method, value):
        """执行批量修改操作"""
        cursor = self.current_connection.cursor()
        
        try:
            # 获取主键信息
            pk = self.get_primary_key(table_name)
            if isinstance(pk, list):
                pk_col = pk[0]
            else:
                pk_col = pk
            
            col_names = [table_widget.horizontalHeaderItem(c).text() for c in range(table_widget.columnCount())]
            pk_index = col_names.index(pk_col)
            
            success_count = 0
            error_count = 0
            
            # 遍历所有行
            for row in range(table_widget.rowCount()):
                try:
                    # 获取主键值
                    pk_item = table_widget.item(row, pk_index)
                    if pk_item is None:
                        continue
                    
                    pk_value = pk_item.text()
                    
                    # 获取当前值
                    current_item = table_widget.item(row, col_index)
                    if current_item is None:
                        current_value = ""
                    else:
                        current_value = current_item.text()
                    
                    # 根据方法计算新值
                    if method == "replace":
                        new_value = value
                    elif method == "prefix":
                        new_value = value + current_value
                    elif method == "suffix":
                        new_value = current_value + value
                    else:
                        new_value = value
                    
                    # 执行更新
                    sql = f"UPDATE [{table_name}] SET [{col_name}] = ? WHERE [{pk_col}] = ?"
                    cursor.execute(sql, (new_value, pk_value))
                    
                    # 更新表格显示
                    if current_item:
                        current_item.setText(new_value)
                    
                    success_count += 1
                    
                except Exception as e:
                    error_count += 1
                    print(f"更新行 {row} 失败: {e}")
            
            self.current_connection.commit()
            
            if error_count == 0:
                QMessageBox.information(self, "成功", f"整列修改完成 ({success_count} 条记录)")
            else:
                QMessageBox.warning(self, "部分成功", 
                                  f"成功修改 {success_count} 条记录，失败 {error_count} 条记录")
                
        except Exception as e:
            QMessageBox.critical(self, "批量修改失败", str(e))
            self.current_connection.rollback()
        finally:
            cursor.close()

    def insert_row(self, table_name, table_widget):
        col_names = [table_widget.horizontalHeaderItem(c).text() for c in range(table_widget.columnCount())]
        dlg = EditDialog(col_names, self)
        if dlg.exec_() == QDialog.Accepted:
            vals = dlg.values
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
        pk = self.get_primary_key(table_name)
        if isinstance(pk, list):
            pk_col = pk[0]
        else:
            pk_col = pk
        pk_index = col_names.index(pk_col)
        pk_item = table_widget.item(selected, pk_index)
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