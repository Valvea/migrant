import csv
import io
import time
import os
import sys
from threading import Thread
import pandas as pd
from PySide2 import QtCore, QtGui, QtWidgets
from PySide2.QtCore import Qt, Signal, QModelIndex
from patent_request import Check_patent
import ctypes
from numpy import nan



myappid = u'migrant+'  # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)



class CommandEdit(QtWidgets.QUndoCommand):
    def __init__(self, itemWidget, item, textBeforeEdit, row=None):
        super(CommandEdit, self).__init__()
        self.itemWidget = itemWidget
        self.row = row
        if self.row == 'delete row' or self.row == 'add_empty_row':
            if self.row == 'delete row':
                self.dataBeforeEdit = self.itemWidget.get_data()
                self.dataAfterEdit = self.dataBeforeEdit.iloc[self.dataBeforeEdit.index != item.row()]
                # print(self.dataBeforeEdit,self.dataAfterEdit)
            if self.row == 'add_empty_row':
                self.dataAfterEdit = self.itemWidget.get_data()
                self.dataBeforeEdit = self.dataAfterEdit.drop([item.row()])
        else:
            self.textBeforeEdit = textBeforeEdit
            self.role = Qt.DecorationRole if \
                self.itemWidget.headerData(item.column(), Qt.Horizontal, Qt.DisplayRole) \
                == 'Каптча' else Qt.DisplayRole
            self.textAfterEdit = item.data(self.role)
            self.index = item

    def redo(self):
        if self.row == 'delete row' or self.row == 'add_empty_row':
            self.itemWidget.beginResetModel()
            self.itemWidget._data = self.dataAfterEdit
            self.itemWidget.endResetModel()
        else:
            self.itemWidget.beginResetModel()
            self.itemWidget.blockSignals(True)
            self.itemWidget.setData(self.index, self.textAfterEdit, self.role)
            self.itemWidget.blockSignals(False)
            self.itemWidget.layoutChanged.emit()
            self.itemWidget.endResetModel()

    def undo(self):
        if self.row == 'delete row' or self.row == 'add_empty_row':
            self.itemWidget.beginResetModel()
            self.itemWidget._data = self.dataBeforeEdit
            self.itemWidget.endResetModel()
        else:
            self.itemWidget.beginResetModel()
            self.itemWidget.blockSignals(True)
            self.itemWidget.setData(self.index, self.textBeforeEdit, self.role)
            self.itemWidget.blockSignals(False)
            self.itemWidget.layoutChanged.emit()
            self.itemWidget.endResetModel()


class MyDelegate(QtWidgets.QItemDelegate):

    def setEditorData(self, editor, index):
        if index.model().headerData(index.column(), Qt.Horizontal, Qt.DisplayRole) != 'Каптча':
            text = index.data(Qt.EditRole) or index.data(Qt.DisplayRole)
            editor.setText(text)


class TableModel(QtCore.QAbstractTableModel):

    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

        for column in self._data.columns:
            if column in ['Патент план', 'Дата выдачи', 'Оплачен до','План регистрация','Регистрация']:
                self._data[column] = pd.to_datetime(self._data[column]).dt.date


        self._data['Патент план'] = self._data["Оплачен до"] + pd.Timedelta(days=20)
        self._data['Чек план'] = self._data["Оплачен до"] + pd.Timedelta(days=10)
        self._data['План регистрация']= self._data['Регистрация'] - pd.Timedelta(days=10)

        print(self._data.info())
        self.checks = {}
        self.inp_ = Communicate()
        self.copy_inx = 0

    def checkState(self, index):
        if index in self.checks.keys():
            return self.checks[index]
        else:
            return Qt.Unchecked

    def setData(self, index, value, role):
        # self.beginResetModel()
        if role == Qt.CheckStateRole:
            self.checks[QtCore.QPersistentModelIndex(index)] = value
            self.inp_.check_state.emit(value, index)
            return True
        if self.headerData(index.column(),Qt.Horizontal,Qt.DisplayRole)=='Регистрация':
            self._data.loc[index.row(),'План регистрация']=(pd.to_datetime(value)-pd.Timedelta(days=10)).date()
        self._data.iloc[index.row(), index.column()] = value
        self.dataChanged.emit(index, index)
        # self.endResetModel()

        return True

    def flags(self, index):
        if self.headerData(index.column(), Qt.Horizontal, Qt.DisplayRole) == 'Каптча':
            return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsSelectable | Qt.ItemIsUserCheckable
        return QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsSelectable | Qt.ItemIsDragEnabled |Qt.ItemIsDropEnabled

    def data(self, index, role):

        value = self._data.iloc[index.row(), index.column()]

        if role == Qt.DisplayRole and self.headerData \
                    (index.column(), Qt.Horizontal, Qt.DisplayRole) != 'Каптча':
            if pd.isna(value):
                return ' '
            return str(value)
        elif self.headerData \
                    (index.column(), Qt.Horizontal, Qt.DisplayRole) == 'Каптча':
            if role == Qt.CheckStateRole:
                return self.checkState(QtCore.QPersistentModelIndex(index))
            else:
                return value
        elif role == Qt.ForegroundRole:
            if value == "Патент не оплачен!" or value== 'Не верно введена каптча':
                return QtGui.QColor('red')
            elif value == 'Патент оплачен!':
                return QtGui.QColor('green')
            elif value =='Необходимо оформить новый патент!':
                return QtGui.QColor('blue')

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, index):
        return self._data.shape[1]

    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])
            if orientation == Qt.Vertical:
                return str(self._data.index[section])

    def add_empty_row(self):
        self.beginResetModel()
        try:
            self._data = self._data.append(
                pd.DataFrame(columns=self._data.columns, data=([[None] * len(self._data.columns)]),
                             index=[self._data.index[-1] + 1]))
        except IndexError:
            self._data = self._data.append(
                pd.DataFrame(columns=self._data.columns, data=([[None] * len(self._data.columns)]),
                             index=[0]))
        self.layoutChanged.emit()
        self.endResetModel()
        self.inp_.add_delete_row.emit(self.createIndex(self._data.index[-1], 0), 'add_empty_row')

    def get_data(self):
        return self._data

    def delete_row(self, indexs):
        self.inp_.add_delete_row.emit(indexs[0].row(), 'delete row')
        self._data.drop([index.row() for index in indexs ], inplace=True)
        self._data.reset_index(inplace=True, drop=True)
        self.layoutChanged.emit()

    def rename_column(self, index, new_name):
        self._data.rename(columns={self._data.columns[index]: new_name}, inplace=True)
        return True

    def add_column(self, index):
        print(index)
        self.beginResetModel()
        try:
            self._data.insert(index, f"Новый столбец {''}", value='')
        except ValueError:
            self._data.insert(index, f"Новый столбец {self.copy_inx + 1}", value='')
            self.copy_inx += 1
        self.endResetModel()

    def delete_column(self, index):
        self.beginResetModel()
        self._data.drop(self._data.columns[index], axis=1, inplace=True)
        self.endResetModel()

    def sorting(self, column, ascending=True):
        self.beginResetModel()
        self._data.sort_values(by=self._data.columns[column], ascending=ascending,inplace=True,ignore_index=True)
        self.endResetModel()



class Communicate(QtCore.QObject):
    closeApp = Signal()
    send_data = Signal(pd.Series)
    check_state = Signal(int, QtCore.QModelIndex)
    redy_to_send = Signal(str, int)
    add_delete_row = Signal(QtCore.QModelIndex, str)


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super().__init__()
        self.textBeforeEdit = ""
        self.index = None
        self.undoStack = QtWidgets.QUndoStack(self)
        self.setWindowTitle('Проверка патента на действительность')
        self.model = TableModel(pd.read_excel('данные.xlsx',dtype=object))

        self.table = QtWidgets.QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.openMenu)
        self.table.setFont(QtGui.QFont("Arial", 11))
        self.model.dataChanged.connect(self.changed)
        self.table.clicked[QtCore.QModelIndex].connect(self.clicked)
        self.c = Communicate()
        self.c.closeApp.connect(self.close)
        self.c.send_data.connect(self.checking)
        self.model.inp_.add_delete_row.connect(self.changed)
        self.model.inp_.check_state.connect(self.check_capcha_input)
        self.table.setSortingEnabled(True)

        self.table.setDropIndicatorShown(True)
        self.table.setDragDropOverwriteMode(False)
        self.table.installEventFilter(self)
        self.table.viewport().installEventFilter(self)

        self.setWindowIcon(QtGui.QIcon(os.getcwd() + r'\logo.png'))

        self.add_Tool_Menu_bar()
        self.setGeometry(100, 300, 1100, 500)

        self.header = self.table.horizontalHeader()
        self.table.resizeColumnsToContents()
        self.header.setSectionsMovable(True)
        self.header.setDragEnabled(True)
        self.header.setDragDropMode(self.table.InternalMove)
        self.row_header = self.table.verticalHeader()
        self.row_header.setSectionsMovable(True)
        self.row_header.setDragEnabled(True)
        self.row_header.setDragDropMode(self.table.InternalMove)

        self.delegate = MyDelegate()
        self.table.setItemDelegate(self.delegate)
        self.setCentralWidget(self.table)
        self.init_rc_menu()
        self.column_settings()
        self.set_line_pos(0)
        self.header.setContextMenuPolicy(Qt.CustomContextMenu)
        self.header.customContextMenuRequested.connect(self.header_pressed)



    def eventFilter(self, source, event):
        if event.type()==QtCore.QEvent.Drop:
            self.dragdropCells_Off()
            self.deleteSelection()

        if (event.type() == QtCore.QEvent.KeyPress and
                event.matches(QtGui.QKeySequence.Copy)):
            self.copySelection()
            return True
        elif (event.type() == QtCore.QEvent.KeyPress and
              event.matches(QtGui.QKeySequence.Paste)):
            self.pasteSelection()
            return True
        elif (event.type() == QtCore.QEvent.KeyPress and
              event.matches(QtGui.QKeySequence.Delete)):
            self.deleteSelection()
            return True

        return super(QtWidgets.QMainWindow, self).eventFilter(source, event)




    def column_manage(self, action):
        action = action.text()

        if action == 'Вставить справа':
            self.column_actions[action](self.column_selected + 1)
        elif action == 'Убыванию':
            self.column_actions[action](self.column_selected, ascending=False)
        elif action in ['целочисленный','дробный','дата','текстовый']:

            self.model._data.iloc[:, self.column_selected]=self.model._data.iloc[:, self.column_selected].fillna(0)
            self.model._data.iloc[:,self.column_selected]=self.column_actions[action](self.column_selected)
            print(self.model._data.info())
        else:
            self.column_actions[action](self.column_selected)

        self.table.resizeColumnsToContents()

    def column_settings(self):
        self.column_menu = QtWidgets.QMenu()
        sort_column = QtWidgets.QMenu("Cортировать по", self.column_menu)
        type_column=QtWidgets.QMenu("Установить тип данных", self.column_menu)
        past_left = QtWidgets.QAction("Вставить слева", self)
        past_right = QtWidgets.QAction("Вставить справа", self)
        delete_column = QtWidgets.QAction("Удалить столбец", self)
        rename_column = QtWidgets.QAction("Переименовать столбец", self)

        sort_column_up = QtWidgets.QAction("Возрастанию", self)
        sort_column_down = QtWidgets.QAction("Убыванию", self)
        sort_column.addActions([sort_column_up, sort_column_down])

        column_types=[QtWidgets.QAction(x, self)
                      for x in ['целочисленный','дробный','текстовый','дата',] ]
        type_column.addActions(column_types)

        self.column_menu.addMenu(sort_column)
        self.column_menu.addMenu(type_column)

        self.column_menu.addActions([past_left, past_right, delete_column, rename_column])
        self.column_menu.triggered[QtWidgets.QAction].connect(self.column_manage)

        self.line = QtWidgets.QLineEdit(parent=self.header.viewport())  # Create
        self.line.setAlignment(QtCore.Qt.AlignHCenter)
        self.line.setHidden(False)
        self.column_selected = None
        self.line.editingFinished.connect(self.rename_column)
        self.column_actions = {'Переименовать столбец': self.transfer_column_name,
                               'Вставить слева': self.model.add_column,
                               'Вставить справа': self.model.add_column,
                               'Удалить столбец': self.model.delete_column,
                               'Возрастанию': self.model.sorting,
                               'Убыванию': self.model.sorting,
                               }
        self.column_actions.update({k:v for k,v in zip(['целочисленный','дробный','текстовый','дата',],
                                    [lambda x:self.model._data.iloc[:,x].astype('int32',errors='ignore'),
                                     lambda x:self.model._data.iloc[:,x].astype('float',errors='ignore'),
                                     lambda x:self.model._data.iloc[:,x].astype('str',errors='ignore'),
                                     lambda x:pd.to_datetime(self.model._data.iloc[:,x]).dt.date] )})


    def set_line_pos(self, position, hide=True):
        edit_geometry = self.line.geometry()
        edit_geometry.setHeight(self.header.height())
        edit_geometry.setWidth(self.header.sectionSize(position))
        edit_geometry.moveLeft(self.header.sectionViewportPosition(position))
        self.line.setGeometry(edit_geometry)
        self.line.setText(self.model._data.columns[position])
        self.line.setHidden(hide)

    def transfer_column_name(self, index):
        self.set_line_pos(index, False)
        self.line.blockSignals(False)
        self.line.setFocus()
        self.line.selectAll()

    def rename_column(self):
        if self.model.rename_column(self.column_selected, self.line.text()):
            self.line.blockSignals(True)
            self.line.setHidden(True)
            self.line.setText('')

    def header_pressed(self, position):
        self.column_selected = self.header.logicalIndexAt(position)
        self.column_menu.exec_(self.table.viewport().mapToGlobal(position))

    def dragdropCells_On(self):
        print('drag')
        self.table.setDragEnabled(True)
        self.table.setDragDropMode(self.table.DragDrop)

    def dragdropCells_Off(self):
        self.table.setDragEnabled(False)

    def init_rc_menu(self):
        self.menu = QtWidgets.QMenu()
        copy = QtWidgets.QAction("Копировать", self)
        delete = QtWidgets.QAction("Удалить строку", self)
        check = QtWidgets.QAction("Проверить патент ", self)
        delete_strings = QtWidgets.QAction("Удалить данные", self)
        move_items = QtWidgets.QAction("Переместить элементы", self)
        self.menu.addAction(copy)
        self.menu.addAction(delete)
        self.menu.addAction(check)
        self.menu.addAction(delete_strings)
        self.menu.addAction(move_items)
        copy.triggered.connect(self.copySelection)
        delete.triggered.connect(self.delete_row)
        check.triggered.connect(self.get_data_for_request)
        delete_strings.triggered.connect(self.deleteSelection)
        move_items.triggered.connect(self.dragdropCells_On)

    def changed(self, item, row=None):

        command = CommandEdit(self.model, item, self.textBeforeEdit, row)
        self.undoStack.push(command)

    def delete_row(self):
        self.model.beginResetModel()
        self.model.delete_row(self.table.selectedIndexes())
        self.model.endResetModel()



    def deleteSelection(self):
        selection = self.table.selectedIndexes()
        if selection:
            selection = self.table.selectedIndexes()
            rows = sorted(set(index.row() for index in selection))
            columns = sorted(set(index.column() for index in selection))
            self.model.beginResetModel()
            self.model._data.iloc[rows[0]:rows[-1] + 1, columns[0]:columns[-1] + 1] = None
            self.model.endResetModel()

    def copySelection(self):
        selection = self.table.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                column = index.column() - columns[0]
                table[row][column] = index.data()
            stream = io.StringIO()
            buffer = QtWidgets.QApplication.clipboard().text()
            delimiter = ' ' if buffer.count(' ') > 0 else '\t'
            csv.writer(stream, delimiter=delimiter).writerows(table)
            QtGui.QGuiApplication.clipboard().setText(stream.getvalue())

    def pasteSelection(self):
        selection = self.table.selectedIndexes()
        if selection:
            model = self.model
            buffer = QtWidgets.QApplication.clipboard().text()
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            delimiter=' ' if buffer.count(' ')> 0 else '\t'
            reader = csv.reader(io.StringIO(buffer), delimiter=delimiter)
            if len(rows) == 1 and len(columns) == 1:
                for i, line in enumerate(reader):
                    for j, cell in enumerate(line):
                        model.setData(model.index(rows[0] + i, columns[0] + j), cell, Qt.DisplayRole)
            else:
                arr = [[cell for cell in row] for row in reader]
                for index in selection:
                    row = index.row() - rows[0]
                    column = index.column() - columns[0]
                    model.setData(model.index(index.row(), index.column()), arr[row][column], Qt.DisplayRole)
        return


    def get_data_for_request(self):

        for index in self.table.selectedIndexes():
            captcha_row = index.row()
            data = self.model.get_data().iloc[captcha_row,
                                              [index for index, cols in enumerate(self.model._data.columns)
                                               if cols in ['Номер патента', 'Серия патента', 'Дата выдачи']]]

            if sum(map(lambda x: False if not x else True, data.values)) < 3:
                QtWidgets.QMessageBox.critical(self, 'Неверно введены данные!',
                                               f"Проверьте правильность данных\nвведенных в ячейке\nCтрока:{captcha_row}",
                                               QtWidgets.QMessageBox.Ok)
            else:
                self.c.send_data.emit(data)
                time.sleep(0.5)

    def display_captcha(self, captcha, _captcha_row):

        self.model.beginResetModel()
        self.columns_table = {k: v for v, k in enumerate(self.model._data.columns)}
        self.captcha_index = self.model.createIndex(_captcha_row, self.columns_table['Каптча'])
        self.model.setData(self.captcha_index, captcha, Qt.DecorationRole)
        self.header.resizeSection(self.columns_table['Каптча'], 170)
        self.table.verticalHeader().resizeSection(_captcha_row, 35)
        self.model.endResetModel()

    def check_capcha_input(self, state, index):
        if state == 2 and isinstance(index.data(), str):
            self.c.redy_to_send.emit(index.data(), index.row())

    def response(self, answer):
        self.columns_index = {k: v for v, k in enumerate(self.model._data.columns)}

        if answer['resultCode'] == 'INVALID_CAPTCHA_ERROR' \
                or answer['resultCode'] == 'EMPTY_CAPTCHA_ERROR' \
                or answer['resultCode'] == 'NOT_FOUND_PATENT_ERROR':
            patent_index = self.model.createIndex(answer['captcha_row'], self.columns_index['Статус патента'])
            self.model.setData(patent_index,
                               'Не верно введена каптча' if answer['resultCode'] == 'INVALID_CAPTCHA_ERROR' or answer[
                                   'resultCode'] == 'EMPTY_CAPTCHA_ERROR'
                               else "Патент не найден!", Qt.DisplayRole)


            # self.check_patent.status = False


        if answer['resultCode'] == 'SUCCESS':
            patent_index = self.model.createIndex(answer['captcha_row'], self.columns_index['Статус патента'])
            date_patent_index = self.model.createIndex(answer['captcha_row'], self.columns_index['Оплачен до'])
            date_patent_index_to_pay = self.model.createIndex(answer['captcha_row'], self.columns_index['Патент план'])
            plan_check_index = self.model.createIndex(answer['captcha_row'], self.columns_index['Чек план'])

            if answer['status'] == 'PATENT_IS_NOT_PAID':
                self.model.setData(patent_index, 'Патент не оплачен!', Qt.DisplayRole)

            if answer['status'] == 'PATENT_VALID':
                self.model.setData(patent_index, 'Патент оплачен!', Qt.DisplayRole)

            if answer['status'] == 'PATENT_EXPIRED':
                self.model.setData(patent_index, 'Необходимо оформить новый патент!', Qt.DisplayRole)

            self.model.setData(date_patent_index,
                               pd.to_datetime(answer['statusDetail']['paidTillDate']).date(), Qt.DisplayRole)
            self.model.setData(date_patent_index_to_pay,
                               pd.to_datetime(answer['statusDetail']['paidTillDate']).date() + pd.Timedelta(days=20),
                               Qt.DisplayRole)

            self.model.setData(plan_check_index,
                               pd.to_datetime(answer['statusDetail']['paidTillDate']).date() + pd.Timedelta(
                                   days=10), Qt.DisplayRole)


                # self.model.itemData(patent_index)
            # self.check_patent.status = False

    def openMenu(self, position):

        self.menu.exec_(self.table.viewport().mapToGlobal(position))

    def closeEvent(self, event):
        result = QtWidgets.QMessageBox.question(self,
                                                "Подтвердите выход...",
                                                "Вы уверены что хотите выйти?\n"
                                                "Изменения не сохранятся!",
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        event.ignore()

        if result == QtWidgets.QMessageBox.Yes:
            event.accept()

    def add_empty_row(self):
        self.model.add_empty_row()

    def clicked(self, index):
        item = index.data()
        self.textBeforeEdit = item
        self.index = index

    def add_Tool_Menu_bar(self):

        layout = QtWidgets.QHBoxLayout()
        bar = self.menuBar()
        file = bar.addMenu("Файл")

        save = QtWidgets.QAction("Сохранить", self)
        save.setShortcut("Ctrl+S")
        file.addAction(save)

        saveFile = QtWidgets.QAction("Сохранить как", self)
        saveFile.setShortcut("Ctrl+B")
        file.addAction(saveFile)

        quit = QtWidgets.QAction("Выход", self)
        file.addAction(quit)
        file.triggered[QtWidgets.QAction].connect(self.processtrigger)
        self.setLayout(layout)

        tool = QtWidgets.QToolBar()
        self.empty_string = QtWidgets.QAction \
            (QtGui.QIcon(os.getcwd() + r'\icons\add.png'), 'Добавить пустую строку')
        self.empty_string.setShortcut("Ctrl+M")
        self.cancel = QtWidgets.QAction \
            (QtGui.QIcon(os.getcwd() + r'\icons\undo.png'), 'Отменить изменения')
        self.cancel.setShortcut("Ctrl+Z")
        self.redo = QtWidgets.QAction \
            (QtGui.QIcon(os.getcwd() + r'\icons\redo.png'), 'Вперед')

        self.empty_string.triggered.connect(self.add_empty_row)
        self.cancel.triggered.connect(self.undoStack.undo)
        self.redo.triggered.connect(self.undoStack.redo)

        tool.addAction(self.empty_string)
        tool.addAction(self.cancel)
        tool.addAction(self.redo)

        tool.setStyleSheet('QToolBar{spacing:15px;}')
        self.addToolBar(tool)

    def file_save(self):
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Сохранить таблицу', os.getcwd(), '.xlsx')
        self.model.get_data().to_excel(name[0] + name[1])

    def checking(self, data):
        self.check_patent = Check_patent()
        self.check_patent.signal.captcha.connect(self.display_captcha)
        self.check_patent.signal.response.connect(self.response)
        self.c.redy_to_send.connect(self.check_patent.unlock)
        self.request_thread = Thread(target=self.check_patent.get_patent_status, daemon=True, args=(data,))
        self.request_thread.start()

    def processtrigger(self, q):
        actions = {'Сохранить': self.model.get_data().to_excel("данные.xlsx", index=False),
                   'Выход': self.c.closeApp.emit,
                   'Сохранить как': self.file_save
                   }
        actions[q.text()]()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()