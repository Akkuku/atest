import sys
import datetime
import re
import pandas as pd
from pathlib import Path
from PySide6 import QtCore
from PySide6.QtWidgets import QApplication, QGridLayout, QWidget, QLabel, QPushButton, QLineEdit, QVBoxLayout, QCheckBox, QTableWidget, QStyledItemDelegate,QItemDelegate, QTableWidgetItem, QHeaderView
from PySide6.QtGui import QColor, QFont, QIcon, QRegularExpressionValidator, QDoubleValidator
from PySide6.QtCore import Qt
from atest import Atest
from sqlalchemy import create_engine
from sqlalchemy_utils import create_database, database_exists

atest_path = Path('./resources/atesty/')
docs_path = Path('./resources/faktury/')
db_path = Path('./resources/params/params.db')
icon_path = 'resources/icons/icon.png'
today_date = datetime.date.today()

class RegexDelegate(QStyledItemDelegate):
    def __init__(self, regex):
        super().__init__()
        self.regex = regex

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setValidator(QRegularExpressionValidator(self.regex))
        return editor

class FloatDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super().__init__()

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setValidator(QDoubleValidator())
        return editor

class TableWidget(QTableWidget):
    def __init__(self, df):
        super().__init__()
        self.df_original = df
        self.df = df
        self.df_editable = self.df.iloc[[2,3,4,8],1:-1]

        # set table dimension
        nRows, nColumns = self.df_editable.shape
        self.setColumnCount(nColumns)
        self.setRowCount(nRows)

        self.setHorizontalHeaderLabels(self.df.iloc[0,1:])
        self.horizontalHeader().setStyleSheet("::section {"'''
            background-color: lightblue; 
            font-family: Times;
            font-size: 14px;}''')
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.setVerticalHeaderLabels(self.df.iloc[[2,3,4,8],0])
        self.verticalHeader().setStyleSheet("::section {"'''
            background-color: lightblue; 
            font-family: Times;
            font-size: 14px;}''')
        self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # data insertion
        for i in range(self.rowCount()):
            for j in range(self.columnCount()):
                item = QTableWidgetItem(str(self.df_editable.iloc[i, j]))
                item.setTextAlignment(Qt.AlignCenter)
                item.setFont(QFont('Arial', 12))
                if j==0:
                    pass
                elif j<10 or j==15:
                    item.setBackground(QColor(255, 255, 212))
                elif j<13 or j==16:
                    item.setBackground(QColor(255, 230, 212))
                else:
                    item.setBackground(QColor('light cyan'))
                self.setItem(i, j, item)

        self.cellChanged[int, int].connect(self.updateDF)   

        # data editing validation
        float_delegate = FloatDelegate()
        self.setItemDelegateForRow(0, float_delegate)
        self.setItemDelegateForRow(1, float_delegate)
        self.setItemDelegateForRow(2, float_delegate)

    def updateDF(self, row, column):
        text = self.item(row, column).text()
        self.df_editable.iloc[row, column] = text
        self.df.iloc[2:5,1:-1] = self.df_editable.iloc[0:3,:]
        self.df.iloc[8,1:-1] = self.df_editable.iloc[3,:]

class Atest_window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Atest")
        self.resize(450, 200)
        self.setWindowIcon(QIcon(icon_path))
        layout = QGridLayout(self)

        last_edit_date = datetime.datetime.fromtimestamp(db_path.stat().st_mtime).date()
        days_since_edit = (today_date - last_edit_date).days

        self.document_id = ''

        self.accept_style = '''
            color: green; 
            font-family: Arial;
            font-size: 20px;'''
        self.rejected_style = '''
            color: red; 
            font-family: Arial;
            font-size: 20px;'''

        self.fs_checkbox = QCheckBox("Najnowsza faktura")
        self.fs_checkbox.setChecked(True)
        self.fs_checkbox.setFont(QFont('Arial', 12))
        layout.addWidget(self.fs_checkbox, 0,0)
        self.fs_checkbox.stateChanged.connect(self.check_valid)

        self.fs_text = QLabel("Nr faktury:")
        self.fs_text.setFont(QFont('Arial', 12))
        layout.addWidget(self.fs_text, 1, 0)
        self.fs_text.hide()

        self.fs_edit = QLineEdit(f"FS ###/{today_date.year}")
        self.fs_edit.setValidator(QRegularExpressionValidator(r"FS .*\/[0-9]{4}"))
        self.fs_edit.setFont(QFont('Arial', 12))
        layout.addWidget(self.fs_edit, 1, 1)
        self.fs_edit.textChanged.connect(self.change_doc_id)
        self.fs_edit.hide()
        
        edit_params_button = QPushButton("Edytuj parametry produktów")
        edit_params_button.setFont(QFont('Arial', 12))
        layout.addWidget(edit_params_button, 2, 0, 1, 2)
        edit_params_button.clicked.connect(self.edit_param)

        last_edit_text = QLabel(f'Dni od ostatniej edycji parametrów: {days_since_edit}')
        if days_since_edit >= 20:
            last_edit_text.setStyleSheet(self.rejected_style)
        layout.addWidget(last_edit_text, 3, 0)

        create_button = QPushButton("Wygeneruj atest")
        create_button.setFont(QFont('Arial', 12))
        layout.addWidget(create_button, 4, 0, 1, 2)
        create_button.clicked.connect(self.generate_atest)

        self.response_text = QLabel("Wygeneruj najnowszy", alignment=Qt.AlignBottom)
        self.response_text.setStyleSheet(self.accept_style)
        layout.addWidget(self.response_text, 5, 0, 1, 2)

    def validate_document(self, name) -> bool:
        if not re.match(r"FS [0-9]+\/[0-9]{4}", name):
            self.response_text.setText(f"Wpisz poprawnny nr faktury (np. FS 123/{today_date.year})")
            self.response_text.setStyleSheet(self.rejected_style)
            return False
        else:
            self.response_text.setText(f"Wygeneruj atest dla {self.document_id}")
            self.response_text.setStyleSheet(self.accept_style)
            return True
    
    @QtCore.Slot()
    def edit_param(self):
        self.w = Params_window()
        self.w.show()

    @QtCore.Slot()
    def generate_atest(self):
        done = False
        self.response_text.setText("Trwa generowanie atestu")
        if self.fs_checkbox.isChecked():
            atest = Atest(atest_path, docs_path, db_path)
            done = atest.make_pdf()
            if done:
                self.response_text.setText("Atest wygenerowany")
        elif self.validate_document(self.document_id):
            try:
                atest = Atest(atest_path, docs_path, db_path, self.document_id)
                done = atest.make_pdf()
                if done:
                    self.response_text.setText("Atest wygenerowany")
            except:
                self.response_text.setText("Taka faktura nie istnieje")
                self.response_text.setStyleSheet(self.rejected_style)
        else:
            self.response_text.setText(f"Wpisz poprawnny nr faktury (np. FS 123/{today_date.year})")
            self.response_text.setStyleSheet(self.rejected_style)
    
    @QtCore.Slot()
    def change_doc_id(self):
        self.document_id = self.fs_edit.text()
        if not self.fs_checkbox.isChecked():
            self.validate_document(self.fs_edit.text())

    @QtCore.Slot()
    def check_valid(self):
        if self.fs_checkbox.isChecked():
            self.response_text.setText("Wygeneruj najnowszy")
            self.response_text.setStyleSheet(self.accept_style)
            self.fs_text.hide()
            self.fs_edit.hide()
        else:
            self.fs_text.show()
            self.fs_edit.show()
            if self.validate_document(self.document_id):
                self.document_id = self.fs_edit.text()
                self.response_text.setText(f"Wygeneruj atest dla {self.document_id}")
                self.response_text.setStyleSheet(self.accept_style)

class Params_window(QWidget):
    engine = create_engine(f'sqlite:///{db_path}')
    params_df = pd.read_sql('params', con=engine, index_col='index')
    params_df.iloc[2] = params_df.iloc[2].apply(lambda x: str(x).replace('.', ','))
    params_df.iloc[3] = params_df.iloc[3].apply(lambda x: str(x).replace('.', ','))
    params_df.iloc[4] = params_df.iloc[4].apply(lambda x: str(x).replace('.', ','))

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Parametry")
        self.setGeometry(100, 300, 1700, 500)
        self.setWindowIcon(QIcon(icon_path))
        layout = QVBoxLayout(self)

        last_edit_date = datetime.datetime.fromtimestamp(db_path.stat().st_mtime).date()
        days_since_edit = (today_date - last_edit_date).days

        last_edit_text = QLabel(f'Dni od ostatniej edycji parametrów: {days_since_edit}')
        if days_since_edit >= 29:
            last_edit_text.setStyleSheet('color: red')
        layout.addWidget(last_edit_text, alignment=QtCore.Qt.AlignLeft)

        self.param_table = TableWidget(self.params_df)
        layout.addWidget(self.param_table)

        save_button = QPushButton('Zapisz')
        layout.addWidget(save_button)
        save_button.clicked.connect(self.save_params)
    
    @QtCore.Slot()
    def save_params(self):
        engine = create_engine(f'sqlite:///{db_path}')
        if not database_exists(engine.url):
            create_database(engine.url)
        else: 
            engine.connect()
        self.param_table.df.to_sql('params', con=engine, if_exists='replace')
        
        self.hide()
    

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Atest_window()
    window.show()
    sys.exit(app.exec())