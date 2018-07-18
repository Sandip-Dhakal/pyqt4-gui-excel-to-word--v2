
# pyinstaller main.spec

import sys
from PyQt4 import QtGui, QtCore
from output import Ui_MainWindow
from PyQt4.QtGui import QMessageBox
from translateClass import Transformer
from sheetNames import SheetNames
from docx import Document
from functools import partial

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

class MainProgram(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self, parent = None):
        super(MainProgram, self).__init__(parent)
        self.setupUi(self)
        self.actionOpen.triggered.connect(self.file_open)
        self.pushButton.clicked.connect(self.converter)
        self.filename = 'no_file'
        self.document = Document()
        self.answer = Document()
        self.count = []
        # self.lineEdit.setPlaceholderText = '10'
        # self.lineEdit.setInputMask("999")
        self.actionExit.triggered.connect(self.closeWindow)
        self.lineEdits=[]
        # print(self.filename)

    def file_open(self):
        dialog = QtGui.QFileDialog(self)
        dialog.setWindowTitle('Open File')
        dialog.setNameFilter('Excel Files (*.xls *.xlsx *.csv)')
        dialog.setFileMode(QtGui.QFileDialog.ExistingFile)
        if dialog.exec_() == QtGui.QDialog.Accepted:
            self.filename = dialog.selectedFiles()[0]
            self.label_2.setText(self.filename.split('/')[-1])
            xsheets = SheetNames(self.filename).giveNames() 
            for x in xsheets:
                self.count.append(10)

            for idx, sheet in enumerate(xsheets):
                print(sheet)
                self.horizontalLayout_2 = QtGui.QHBoxLayout()
                self.horizontalLayout_2.setObjectName(_fromUtf8("horizontalLayout_2"))
                self.label = QtGui.QLabel(self.centralwidget)
                font = QtGui.QFont()
                font.setPointSize(10)
                font.setBold(True)
                font.setUnderline(False)
                font.setWeight(75)
                self.label.setFont(font)
                self.label.setObjectName(_fromUtf8("label"))
                self.horizontalLayout_2.addWidget(self.label)
                self.lineEdit= QtGui.QLineEdit(self.centralwidget)
                self.lineEdits.append(self.lineEdit)
                self.lineEdit.setMaximumSize(QtCore.QSize(30, 40))
                self.lineEdit.setFocusPolicy(QtCore.Qt.ClickFocus)
                self.lineEdit.setObjectName(_fromUtf8(sheet))
                self.lineEdit.setText("10")
                self.lineEdit.setInputMask("999")
                self.horizontalLayout_2.addWidget(self.lineEdit)
                spacerItem = QtGui.QSpacerItem(10, 20, QtGui.QSizePolicy.Fixed, QtGui.QSizePolicy.Minimum)
                self.horizontalLayout_2.addItem(spacerItem)
                self.label.setText("No. of Questions "+sheet+":")
                self.verticalLayout_3.addLayout(self.horizontalLayout_2)
                
                
    def converter(self):
        if (self.filename == 'no_file'):
            self.showdialog('Open excel file to process!')
        else:
            for x in range(len(self.count)):
                self.count[x] = int(self.lineEdits[x].text()) 
            self.document, self.answer = Transformer(self.filename,  self.count).initiate()
            files_types = "Word (*.docx *.doc)"
            name, filters = QtGui.QFileDialog.getSaveFileNameAndFilter(self, 'Save file', '', files_types)
            # print(name+"Sandip")
            # self.showdialog(name)
            try:
                sand = name.split(".")
                name2= sand[0]+"_answer."+sand[-1]
                self.document.save(name)
                self.answer.save(name2)
                self.showdialog("File created successfully!")
            except FileNotFoundError:
                pass
            
                
    def showdialog(self, messages='Information'):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText(messages)
        msg.setInformativeText("Click OK to continue..")
        msg.setWindowTitle("Info")
        # msg.setDetailedText("The details are as follows:")
        # msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.setStandardButtons(QMessageBox.Ok)
        retval = msg.exec_()

    def closeWindow(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("Exit")
        msg.setInformativeText("Click OK to Exit..")
        msg.setWindowTitle("Exit")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.buttonClicked.connect(self.exitProgram)

        retval = msg.exec_()
        
    def exitProgram(self,i):
        if(i.text()=='OK'):
            sys.exit()    
               

if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    demo = MainProgram()
    demo.show()
    sys.exit(app.exec_())