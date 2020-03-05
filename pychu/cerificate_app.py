#!/usr/bin/python3
import sys
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import QThread
from PyQt5.QtWidgets import QMessageBox, QLineEdit, QApplication, QMainWindow, QFileDialog



class FileEdit(QLineEdit):
    def __init__(self, parent=None):
        super(FileEdit, self).__init__(parent)

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
            if filepath.endswith('.xlsx'):
                self.setText(filepath)
            else:
                dialog = QMessageBox()
                dialog.setWindowTitle("Error: Invalid File")
                dialog.setText("Only Excel (.xlsx) files are accepted")
                dialog.setIcon(QMessageBox.Warning)
                dialog.exec_()

class combodemo(QLineEdit):
   def __init__(self, parent = None):
      super(combodemo, self).__init__(parent)
      
      layout = QHBoxLayout()
      self.cb = QComboBox()
      self.cb.addItem("C")
      self.cb.addItem("C++")
      self.cb.addItems(["Java", "C#", "Python"])
      self.cb.currentIndexChanged.connect(self.selectionchange)
		
      layout.addWidget(self.cb)
      self.setLayout(layout)
      self.setWindowTitle("combo box demo")

   def selectionchange(self,i):
      print ("Items in the list are :")
		
      for count in range(self.cb.count()):
         print (self.cb.itemText(count))
      print ("Current index",i,"selection changed ",self.cb.currentText())

def main():
   app = QApplication(sys.argv)
   ex = combodemo()
   ex.show()
   sys.exit(app.exec_())


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(723, 355)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.Info = QtWidgets.QLabel(self.centralwidget)
        self.Info.setMaximumSize(QtCore.QSize(16777215, 50))
        self.Info.setObjectName("Info")
        self.gridLayout.addWidget(self.Info, 0, 0, 1, 2)
        self.FilePath = FileEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.FilePath.sizePolicy().hasHeightForWidth())
        self.FilePath.setSizePolicy(sizePolicy)
        self.FilePath.setMinimumSize(QtCore.QSize(0, 80))
        self.FilePath.setObjectName("FilePath")
        self.gridLayout.addWidget(self.FilePath, 1, 0, 1, 1)
        self.Browse = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.Browse.sizePolicy().hasHeightForWidth())
        self.Browse.setSizePolicy(sizePolicy)
        self.Browse.setObjectName("Browse")
        self.gridLayout.addWidget(self.Browse, 1, 1, 1, 1)
        self.startButton = QtWidgets.QPushButton(self.centralwidget)
        self.startButton.setMinimumSize(QtCore.QSize(0, 50))
        self.startButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.startButton, 2, 0, 1, 2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.Info.setText(_translate("MainWindow",
                                     "<h3>Please drag and drop the Excel (*.xlsx) file in the text field below to use\
                                      button to browse for the file<h3>"))
        self.Browse.setText(_translate("MainWindow", "Browse"))
        self.startButton.setText(_translate("MainWindow", "Start"))


# Main Worker thread for logic
class WorkerThread(QThread):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.filepath = ''
        # initialize as necessary

    def run(self):
        pass
# working code goes here including any method calls


class widget(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setMaximumSize(500, 200)
        self.status_mbox = QMessageBox()
        self.thread = WorkerThread(self)
        self.ui.FilePath.returnPressed.connect(self.ui.startButton.click)
        self.ui.startButton.clicked.connect(self.start)
        self.ui.Browse.clicked.connect(self.browse)
        self.thread.finished.connect(self.done)

    def start(self):
        self.thread.filepath = self.ui.FilePath.text()
        if not self.thread.filepath:
            self.status_mbox.setText('No file selected')
            self.status_mbox.exec_()
        else:
            self.ui.Info.setText('<h3>Processing Started! Please wait and don\'t close the app or touch any button</h3>')
            self.thread.start()

    def browse(self):
        fileName = QFileDialog.getOpenFileName(self,
                                                "Open Image", "Excel Files(*.xlsx)")
        self.ui.FilePath.setText(fileName[0])

    def done(self):
        self.status_mbox.setText('Operation completed ')
        self.status_mbox.exec_()
        print(self.thread.filepath)
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget1 = widget()
    widget1.show()
    sys.exit(app.exec_())
    

