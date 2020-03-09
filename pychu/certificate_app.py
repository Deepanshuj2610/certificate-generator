#!/usr/bin/python3
import os
import sys

import pandas
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import QThread
from PyQt5.QtWidgets import QMessageBox, QLineEdit, QApplication, QMainWindow, QFileDialog


class FileEdit_spreadsheet(QLineEdit):
    def __init__(self, parent=None):
        super(FileEdit_spreadsheet, self).__init__(parent)

        self.setDragEnabled(True)

    def dragEnterEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dropEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            filepath = str(urls[0].path())[1:]
            filepath = '/' + filepath
            if filepath.endswith('.xlsx') or filepath.endswith('.csv'):
                self.setText(filepath)
            else:
                dialog = QMessageBox()
                dialog.setWindowTitle("Error: Invalid File")
                dialog.setText("Only Excel(.xlsx) and CSV files(.csv) files are accepted")
                dialog.setIcon(QMessageBox.Warning)
                dialog.exec_()


class FileEdit_template(QLineEdit):
    def __init__(self, parent=None):
        super(FileEdit_template, self).__init__(parent)
        self.setDragEnabled(True)

    def dragEnterEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            event.acceptProposedAction()

    def dropEvent(self, event):
        data = event.mimeData()
        urls = data.urls()
        if urls and urls[0].scheme() == 'file':
            filepath = str(urls[0].path())[1:]
            filepath = '/' + filepath
            if filepath.endswith('.jpg') or filepath.endswith('.jpeg'):
                self.setText(filepath)
            else:
                dialog = QMessageBox()
                dialog.setWindowTitle("Error: Invalid File")
                dialog.setText("Only JPEG (.jpeg) files are accepted")
                dialog.setIcon(QMessageBox.Warning)
                dialog.exec_()


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(836, 589)
        MainWindow.setMinimumSize(QtCore.QSize(760, 447))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.Info_Font = QtWidgets.QLabel(self.centralwidget)
        self.Info_Font.setMaximumSize(QtCore.QSize(16777215, 50))
        self.Info_Font.setObjectName("Info_Font")
        self.gridLayout.addWidget(self.Info_Font, 6, 0, 1, 1)
        self.FilePath_spreadsheet = FileEdit_spreadsheet(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.FilePath_spreadsheet.sizePolicy().hasHeightForWidth())
        self.FilePath_spreadsheet.setSizePolicy(sizePolicy)
        self.FilePath_spreadsheet.setMinimumSize(QtCore.QSize(0, 80))
        self.FilePath_spreadsheet.setObjectName("FilePath_spreadsheet")
        self.gridLayout.addWidget(self.FilePath_spreadsheet, 1, 0, 1, 1)
        self.combo_box_font = QtWidgets.QComboBox(self.centralwidget)
        self.combo_box_font.setObjectName("combo_box_font")
        self.gridLayout.addWidget(self.combo_box_font, 7, 0, 1, 1)
        self.FilePath_template = FileEdit_template(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.FilePath_template.sizePolicy().hasHeightForWidth())
        self.FilePath_template.setSizePolicy(sizePolicy)
        self.FilePath_template.setMinimumSize(QtCore.QSize(0, 80))
        self.FilePath_template.setObjectName("FilePath_template")
        self.gridLayout.addWidget(self.FilePath_template, 3, 0, 1, 1)
        self.Info_Folder = QtWidgets.QLabel(self.centralwidget)
        self.Info_Folder.setMaximumSize(QtCore.QSize(16777215, 50))
        self.Info_Folder.setObjectName("Info_Folder")
        self.gridLayout.addWidget(self.Info_Folder, 4, 0, 1, 1)
        self.Browse_output_folder = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.Browse_output_folder.sizePolicy().hasHeightForWidth())
        self.Browse_output_folder.setSizePolicy(sizePolicy)
        self.Browse_output_folder.setObjectName("Browse_output_folder")
        self.gridLayout.addWidget(self.Browse_output_folder, 4, 1, 2, 1)
        self.FolderPath = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.FolderPath.sizePolicy().hasHeightForWidth())
        self.FolderPath.setSizePolicy(sizePolicy)
        self.FolderPath.setMinimumSize(QtCore.QSize(0, 80))
        self.FolderPath.setObjectName("FolderPath")
        self.gridLayout.addWidget(self.FolderPath, 5, 0, 1, 1)
        self.Start = QtWidgets.QPushButton(self.centralwidget)
        self.Start.setMinimumSize(QtCore.QSize(0, 50))
        self.Start.setObjectName("Start")
        self.gridLayout.addWidget(self.Start, 6, 1, 2, 1)
        self.Browse_jpeg = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.Browse_jpeg.sizePolicy().hasHeightForWidth())
        self.Browse_jpeg.setSizePolicy(sizePolicy)
        self.Browse_jpeg.setObjectName("Browse_jpeg")
        self.gridLayout.addWidget(self.Browse_jpeg, 2, 1, 2, 1)
        self.Info_xlsx_csv = QtWidgets.QLabel(self.centralwidget)
        self.Info_xlsx_csv.setMaximumSize(QtCore.QSize(16777215, 50))
        self.Info_xlsx_csv.setObjectName("Info_xlsx_csv")
        self.gridLayout.addWidget(self.Info_xlsx_csv, 0, 0, 1, 1)
        self.Browse_xlsx_csv = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.Browse_xlsx_csv.sizePolicy().hasHeightForWidth())
        self.Browse_xlsx_csv.setSizePolicy(sizePolicy)
        self.Browse_xlsx_csv.setObjectName("Browse_xlsx_csv")
        self.gridLayout.addWidget(self.Browse_xlsx_csv, 0, 1, 2, 1)
        self.Info_template = QtWidgets.QLabel(self.centralwidget)
        self.Info_template.setMaximumSize(QtCore.QSize(16777215, 50))
        self.Info_template.setObjectName("Info_template")
        self.gridLayout.addWidget(self.Info_template, 2, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.Info_Font.setText(
            _translate("MainWindow", "Please select the font for the name you want to use in the certificate"))
        self.Info_Folder.setText(
            _translate("MainWindow", "Please select the output folder (No Drap and Drap supported)."))
        self.Browse_output_folder.setText(_translate("MainWindow", "Browse"))
        self.Start.setText(_translate("MainWindow", "Start"))
        self.Browse_jpeg.setText(_translate("MainWindow", "Browse"))
        self.Info_xlsx_csv.setText(
            _translate("MainWindow", "Please drag and drop the CSV file  below or use browse button."))
        self.Browse_xlsx_csv.setText(_translate("MainWindow", "Browse"))
        self.Info_template.setText(
            _translate("MainWindow", "Please select the template for the certificate in jpeg fomrat."))


# Main Worker thread for logic
class WorkerThread(QThread):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.filepath_template = ''
        self.filepath_spreadsheet = ''
        self.output_folder = ''
        self.font_name = ''
        self.selectFont = None  # font selection
        # initialize as necessary

    def generate(self, name):
        try:
            img = Image.open(self.filepath_template)  # template
        except IOError:
            pass
        draw = ImageDraw.Draw(img)
        draw.text((455, 200), name, font=self.selectFont,
                  fill=(110, 110, 110))  # 430,360 is the x,y co-ordinates 105,105,105 is the code for font colour
        img.save(self.output_folder + name + '.jpeg')  # certificate will be saved with the name of attendee

    def run(self):
        cars = pandas.read_csv(self.filepath_spreadsheet)  # txt file containing the names of the attendee
        df = cars.DataFrame(data, columns=['Name'])
        #cases = cars.Name.tolist()
        self.selectFont = ImageFont.truetype(os.getcwd() + '/text_design/' + self.font_name, size=25)  # font selection

        for i in df:
            i = i.replace("\n", "")
            self.generate(i)


class widget(QMainWindow):

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setMaximumSize(500, 200)
        self.status_mbox = QMessageBox()
        self.thread = WorkerThread(self)
        self.ui.FilePath_spreadsheet.returnPressed.connect(self.ui.Start.click)
        self.ui.FilePath_template.returnPressed.connect(self.ui.Start.click)
        self.ui.Start.clicked.connect(self.start)
        self.ui.Browse_jpeg.clicked.connect(self.browse_template)
        self.ui.Browse_xlsx_csv.clicked.connect(self.browse_spreadsheet)
        self.ui.Browse_output_folder.clicked.connect(self.browse_output_folder)
        self.thread.finished.connect(self.done)

        self.font_list = os.listdir(os.getcwd() + '/text_design/')
        self.ui.combo_box_font.addItems(self.font_list)

    def start(self):
        self.thread.filepath_template = self.ui.FilePath_template.text()
        self.thread.filepath_spreadsheet = self.ui.FilePath_spreadsheet.text()
        self.thread.output_folder = (self.ui.FolderPath.text() + '/').replace('file://', '')
        self.thread.font_name = self.font_list[self.ui.combo_box_font.currentIndex()]

        if not self.thread.filepath_template or not self.thread.filepath_spreadsheet:
            self.status_mbox.setText('No file selected')
            self.status_mbox.exec_()
        else:
            self.ui.Info_xlsx_csv.setText(
                '<h3>Processing Started! Please wait and don\'t close the app or touch any button</h3>')
            self.thread.start()

    def browse_template(self):
        fileName = QFileDialog.getOpenFileName(self, "Open Template", filter="(*.jpg),(*.jpeg)")
        print(type(fileName))
        self.ui.FilePath_template.setText(fileName[0])

    def browse_spreadsheet(self):
        fileName = QFileDialog.getOpenFileName(self, "Open CSV", filter="*.xlsx, *.csv")
        self.ui.FilePath_spreadsheet.setText(fileName[0])

    def browse_output_folder(self):
        folderName = QFileDialog.getExistingDirectory(self, "Select a Folder for Output certificates.")
        self.ui.FolderPath.setText(folderName)

    def done(self):
        self.status_mbox.setText('Operation completed ')
        self.status_mbox.exec_()
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget1 = widget()

    widget1.showMaximized()
    sys.exit(app.exec_())
