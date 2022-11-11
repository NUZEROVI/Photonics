# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'GUI.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QLineEdit
from PyQt5.QtCore import pyqtSlot 
from PyQt5.QtCore import pyqtSignal 

import numpy as np
import pandas as pd

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(743, 219)
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        MainWindow.setFont(font)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.statusbar = QtWidgets.QStatusBar(self.centralwidget)
        self.statusbar.setGeometry(QtCore.QRect(-10, 170, 751, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setItalic(True)
        self.statusbar.setFont(font)
        self.statusbar.setObjectName("statusbar")
        self.txt_file = QtWidgets.QLineEdit(self.centralwidget)
        self.txt_file.setEnabled(True)
        self.txt_file.setGeometry(QtCore.QRect(280, 23, 450, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.txt_file.setFont(font)
        self.txt_file.setObjectName("txt_file")
        self.btn_save = QtWidgets.QPushButton(self.centralwidget)
        self.btn_save.setEnabled(False)
        self.btn_save.setGeometry(QtCore.QRect(570, 90, 151, 71))
        self.btn_save.setStyleSheet("QPushButton{background-color: qradialgradient(spread:pad, cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5, stop:0 rgba(96,167,229,1), stop:0.80 rgba(96,167,229,1), stop:0.81 rgba(0, 0, 0,0));border:none;}"
                                     "QPushButton:pressed{background-color:qradialgradient(spread:pad, cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5, stop:0 rgba(0,91,146,1), stop:0.80 rgba(0,91,146,1), stop:0.81 rgba(0, 0, 0,0));border:none;}")
        
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.btn_save.setFont(font)
        self.btn_save.setObjectName("btn_save")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(11, 100, 250, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.txt_vgate = QtWidgets.QLineEdit(self.centralwidget)
        self.txt_vgate.setEnabled(False)
        self.txt_vgate.setGeometry(QtCore.QRect(280, 100, 280, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.txt_vgate.setFont(font)
        self.txt_vgate.setObjectName("txt_vgate")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(11, 25, 250, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setItalic(False)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label.setObjectName("label")
        MainWindow.setCentralWidget(self.centralwidget)

        self.txt_file.mousePressEvent = self.uploadFile
        self.btn_save.clicked.connect(self.saveFile)       

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Search System"))
        self.btn_save.setText(_translate("MainWindow", "Search"))
        self.label_3.setText(_translate("MainWindow", "Vgate :"))
        self.label.setText(_translate("MainWindow", "Choose Excel File :"))
    
    def runVgate(self, filename, sheetName, searchVgate):
        dataset = pd.read_excel(filename, sheet_name = sheetName) #default first sheet
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
        for i in range(len(nanCol)):
            nanCol[i] = dataset.iloc[0:row, nanIndex[i]:nanIndex[i+1]]
        
        colName = nanCol[0].columns[1]

        value = searchVgate # search value
        newY = []

        if len( nanCol[0].query("`{0}` == @value".format(colName)) ) > 0:
            for i in range(len(nanCol)) : #幾組   
                dd = dataset.query("`{0}` == @value".format(colName))
                #print(dd)
                for j in range(len(dd)):   
                    for p in range(2, nanIndex[1]):
                        ABSIDName = nanCol[i].columns[p] # ABSID || ABSID.1 || ABSID.3 ...
                        newY.append(dd[ABSIDName].to_numpy()[j]) 
            #print(newY)
        else :
            for i in range(len(nanCol)) : #幾組
                VgateName = nanCol[i].columns[1] # Vgate || Vgate.1 || Vgate.2 ... 
                ABSIDName = nanCol[i].columns[2] # ABSID || ABSID.1 || ABSID.3 ...
                vgateArr = nanCol[i][VgateName].to_numpy()
                arr = dataset[nanCol[3].columns[1]].to_numpy()

                above = vgateArr[np.searchsorted(arr,value,'left')-1]
                below = vgateArr[np.searchsorted(arr,value,'right')]
                #print(above, below) 
                X = value
                x1 = min(above, below)
                x2 =  max(above, below)
                dd = nanCol[i].query("`{0}` == @above or `{0}` == @below".format(VgateName))
                dd = dd.sort_index(ascending=True)
                #print(dd)
                
                for a in range(2):   
                    for p in range(2, nanIndex[1]):
                        ABSIDName = nanCol[i].columns[p] # ABSID || ABSID.1 || ABSID.3 ...
                        y1 = dd.query("`{0}` == @x1".format(VgateName))[ABSIDName].to_numpy()[a]
                        y2 = dd.query("`{0}` == @x2".format(VgateName))[ABSIDName].to_numpy()[a]
                        Y = y1 + (X-x1)*(y2-y1)/(x2-x1)
                        newY.append(Y)
                        
                #print(newY)

        tmp = []
        for k in range(0, len(newY), (nanIndex[1] - 2)*2) :
            lt = []
            
            lt.append("[" + nanCol[int(k/(nanIndex[1] - 2)/2)].columns[0] + "]")
            lt.append(value)
        
            for o in range((nanIndex[1] - 2)) :
                for u in range(2):
                    lt.append(newY[u*((nanIndex[1] - 2))+o+k])
            
            tmp.append(lt)


        newCol = [] 
        newCol.append('Cate')
        newCol.append('Vgate')
        for i in range(nanIndex[1]-2):
            newCol.append('pos')
            newCol.append('neg')

        #['Cate', 'Vgate', 'pos', 'neg', 'pos', 'neg', 'pos', 'neg']

        tf = pd.DataFrame(columns=newCol)   
        #tf = pd.DataFrame(columns=['Cate', 'Vgate', 'pos', 'neg', 'pos1', 'neg1', 'pos2', 'neg2'])
        #tf = pd.DataFrame(columns=['Cate', 'Vgate', 'pos', 'neg'])
        for w in range(len(tmp)):
            tf.loc[w] = tmp[w]
        #print(tf)

        # Save to excel
        writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')
        tf.to_excel(writer, sheet_name='Sheet1')
        writer.save()
        self.statusbar.showMessage('[Info] Save File Success')

    def saveFile(self):
        # Check txt_sheet & txt_vgate isNull ?
        """
        if not self.tet_sheet.text(): # isNull
            self.statusbar.showMessage('[Warning] Please input sheet.')
        else:
        """
        if not self.txt_vgate.text():
            self.statusbar.showMessage('[Warning] Please input vgate.')
        else:
            searchVgate = self.txt_vgate.text()
            self.runVgate(self.filename, 0, searchVgate)

            """
            # Check txt_sheet is exist ?
            chk_data = pd.read_excel(self.filename, sheet_name = None)
            sheetName = self.tet_sheet.text()
            searchVgate = self.txt_vgate.text()
            if sheetName in list(chk_data.keys()):
                self.runVgate(self.filename, sheetName, searchVgate)
            else:
                self.statusbar.showMessage('[Error] Not search sheet.')
            """


    def uploadFile(self, event):
        self.filename = QFileDialog.getOpenFileName(filter="excel (*.*)")[0]
        if self.filename:
            self.txtName = self.filename.split("/")[-1]
            self.txt_file.setText(self.txtName)
            self.statusbar.showMessage('[Info] Load File Success.')
            self.txt_vgate.setEnabled(True)
            self.btn_save.setEnabled(True)



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


