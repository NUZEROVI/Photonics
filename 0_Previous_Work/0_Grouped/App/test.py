# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'GUI_TEST.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import numpy as np
import pandas as pd

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(541, 100)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(10, 10, 521, 81))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.btn_upload = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.btn_upload.setObjectName("btn_upload")
        self.horizontalLayout.addWidget(self.btn_upload)
        self.label = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.lineEdit = QtWidgets.QLineEdit(self.horizontalLayoutWidget)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.btn_save = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.btn_save.setObjectName("btn_save")
        self.horizontalLayout.addWidget(self.btn_save)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.btn_upload.clicked.connect(self.uploadFile)
        self.btn_save.clicked.connect(self.saveFile)
        self.lineEdit.textChanged.connect(self.changeValue)
        self.inputValue = 0
        self.filename = ''

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.btn_upload.setText(_translate("MainWindow", "Upload"))
        self.label.setText(_translate("MainWindow", "Input Vgate :"))
        self.btn_save.setText(_translate("MainWindow", "Save"))

    def saveFile(self):
        dataset = pd.read_excel(self.filename, sheet_name='AUO')
        row, column = dataset.shape
        colList = dataset.columns.tolist() # All columns
        nanCol = [] # Declare null columms
        nanIndex = []
        for i in range(len(colList)):
            nan = ((dataset[colList[i]]).isnull() == True)[0] # nan is True or False
            if(nan): # if nan == true (isnull)
                nanCol.append(colList[i]) 
                nanIndex.append(i)
                
        nanIndex.append(column)

        # Divide into groups 
        for i in range(len(nanCol)):
            nanCol[i] = dataset.iloc[0:row, nanIndex[i]:nanIndex[i+1]]

        colName = nanCol[0].columns[1]

        value = self.inputValue   # search value
        newY = [] # new Y List
        # https://www.listendata.com/2020/12/how-to-use-variable-in-query-in-pandas.html
        # nanCol[3].query("`{0}` == -9".format(colName)) 

        if len( nanCol[0].query("`{0}` == @value".format(colName)) ) > 0:
            print('not null')
        
            # show result
            result = dataset.query("`{0}` == @value".format(colName)) 
            print(result)
            writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')
            result.to_excel(writer, sheet_name='Sheet1')
            writer.save()
            
        else :
            for i in range(len(nanCol)) :
                
                VgateName = nanCol[i].columns[1] # Vgate || Vgate.1 || Vgate.2 ... 
                ABSIDName = nanCol[i].columns[2] # ABSID || ABSID.1 || ABSID.3 ...
                vgateArr = nanCol[i][VgateName].to_numpy()
                arr = dataset[nanCol[3].columns[1]].to_numpy()

                above = vgateArr[np.searchsorted(arr,value,'left')-1]
                below = vgateArr[np.searchsorted(arr,value,'right')]
                print(above, below) 
                X = value
                x1 = min(above, below)
                x2 =  max(above, below)
                dd = nanCol[i].query("`{0}` == @above or `{0}` == @below".format(VgateName))
                dd = dd.sort_index(ascending=True)

                for j in range(2) :
                    y1 = dd.query("`{0}` == @x1".format(VgateName))[ABSIDName].to_numpy()[j]
                    y2 = dd.query("`{0}` == @x2".format(VgateName))[ABSIDName].to_numpy()[j]
            
                Y = y1 + (X-x1)*(y2-y1)/(x2-x1)
                print(Y)

                newY.extend((0, X, Y))
            print(newY)
            valList = dataset.columns.tolist()
            tf = pd.DataFrame(columns=valList)
            tf.loc[len(tf)] = newY
            writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')
            tf.to_excel(writer, sheet_name='Sheet1')
            writer.save()

    def changeValue(self, value):
        self.inputValue = value

    def uploadFile(self):

        self.filename = QFileDialog.getOpenFileName(filter="excel (*.*)")[0]
        if self.filename:
            self.fileName = self.filename.split("/")[-1] # Get image name
            print(self.fileName)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())