# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Task1_GUI.ui'
#
# Created by: PyQt5 UI code generator 5.12.3
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox

import pandas as pd
import numpy as np
import math
import xlsxwriter
from pathlib import Path

class initVal:
    def __init__(self):
        self._Freq = []
        self._W = []
        self._V = []
        self._Area = 1 * 1e-3

class pdGroups:
    def __init__(self):
        self.G1 = pd.DataFrame()
        self.G2 = pd.DataFrame()
        self.G3 = pd.DataFrame()

global init
global dfGroups


def indices(lst, item):
    return [i for i, x in enumerate(lst) if x == item]

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(758, 143)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.txt_file = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_file.setFont(font)
        self.txt_file.setObjectName("txt_file")
        self.horizontalLayout.addWidget(self.txt_file)
        self.btn_load = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_load.setFont(font)
        self.btn_load.setObjectName("btn_load")
        self.horizontalLayout.addWidget(self.btn_load)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.txt_v1 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_v1.setFont(font)
        self.txt_v1.setObjectName("txt_v1")
        self.horizontalLayout_2.addWidget(self.txt_v1)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.txt_v2 = QtWidgets.QLineEdit(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.txt_v2.setFont(font)
        self.txt_v2.setObjectName("txt_v2")
        self.horizontalLayout_2.addWidget(self.txt_v2)
        self.btn_cal = QtWidgets.QPushButton(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(18)
        self.btn_cal.setFont(font)
        self.btn_cal.setObjectName("btn_cal")
        self.horizontalLayout_2.addWidget(self.btn_cal)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Bind ui function
        self.btn_load.clicked.connect(self.loadFile)
        self.btn_cal.clicked.connect(self.calculate)

    def loadFile(self):
        self.excelPaths = QFileDialog.getOpenFileNames(filter="CSV Files (*.csv)")[0]

        self.filenames = ""
        if (self.excelPaths):
            for self.filename in self.excelPaths:  # 4
                self.filenames += (Path(self.filename).name + ",")

        self.txt_file.setText(self.filenames)

        # Multiple dataframes
        dfs = []
        global init
        init = initVal()
        for i in range(len(self.excelPaths)):
            dfs.append("df" + str(i))
            dfs[i] = pd.DataFrame()
            dfs[i] = pd.read_csv(self.excelPaths[i], encoding='utf-8', sep='\t')
            dfs[i] = dfs[i].iloc[:, 0][128:].str.split(',', expand=True)
            dfs[i] = dfs[i].rename(columns=dfs[i].iloc[0]).drop(dfs[i].index[0]).reset_index(drop=True).apply(
                pd.to_numeric, 1)  # Convert object to float64

            # Definite Value (_Freq, _V, _Area)
            init._Freq += sorted(list(set(dfs[i]["Freq"].tolist())))
            init._V = dfs[0].query("Freq == @init._Freq[0]")[
                "VBias"].tolist()  # all V #df1.iloc[2:154, 0].tolist()  # 152 (-15 ~ 15)
            init._Area = 1 * 1e-3

        # # 152 rows × 63 columns
        global dfGroups
        dfGroups = pdGroups()
        dfGroups.G1 = pd.DataFrame(columns=init._Freq)  # C
        dfGroups.G2 = pd.DataFrame(columns=init._Freq)  # G
        for i in range(len(init._Freq)):
            init._W.append(2 * (math.pi) * init._Freq[i])  # Definite Value (_W = 2*pi*_Freq)
            listC = []
            listG = []
            for j in range(len(self.excelPaths)):
                listC += dfs[j].query("Freq == @init._Freq[@i]")["C"].tolist()
                listG += dfs[j].query("Freq == @init._Freq[@i]")["G"].tolist()

            dfGroups.G1[init._Freq[i]] = listC
            dfGroups.G2[init._Freq[i]] = listG

        self.txt_v1.setText(str(init._V[0])) # default val (V = -15)
        self.txt_v2.setText(str(init._V[-1])) # default val (V = 15)

    def calculate(self):
        # 0. Update "_V" by windows Form (customized Value)
        # Duplicates element & Get minimum
        v_start = indices(init._V, float(self.txt_v1.text()))
        v_end = indices(init._V, float(self.txt_v2.text()))
        if ((float(self.txt_v1.text()) < 0) and (float(self.txt_v2.text()) > 0)):  # -5 ~ 5
            min_V1 = min(v_start)
            min_V2 = min(v_end)
        else:  # equal or v1 > v2 ..
            min_V1 = min(v_start)
            min_V2 = max(v_end)

        # update _V & dfGroups
        init._V = np.array(init._V)[min_V1: (min_V2+1)].tolist()
        dfGroups.G1 = pd.DataFrame(dfGroups.G1.to_numpy()[min_V1: (min_V2+1)])
        dfGroups.G2 = pd.DataFrame(dfGroups.G2.to_numpy()[min_V1: (min_V2+1)])

        # 1. "W" in colume[0]
        _each_W = []
        for j in range(len(init._Freq)):
            _each_W.append(float(init._W[j]))

        results = {
            "W": sorted(_each_W),
        }
        df = pd.DataFrame(results)

        # 2. cal "Gp_W" (append another column) (behind colume[0])
        for i in range(len(init._V)):
            Gp_W = []
            for j in range(len(init._Freq)):
                Cm = dfGroups.G1.iat[i, j] / init._Area
                Gm = dfGroups.G2.iat[i, j] / init._Area
                Cox = (dfGroups.G1.iloc[:, j].max()) / init._Area

                Gp_W.append(float((init._W[j] * (Cox ** 2) * Gm) / ((Gm ** 2) + (init._W[j] ** 2) * (((Cm - Cox) ** 2)))))

            df.insert(len(df.columns), str(init._V[i]), Gp_W, True)

        # 3. output results
        try:
            #df.to_csv('Gp_W_Results(with_Noise).csv', encoding='utf-8', index=False)
            writer = pd.ExcelWriter('Gp_W_Results(with_Noise).xlsx', engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.save()
            QMessageBox.information(None, 'Message', 'Finished')
        except:
            print("Output error")
            pass

        # 4. Exit()
        sys.exit(app.exec_())

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Task1 - Calulate \"Gp_W\" (with tolerance)"))
        self.label.setText(_translate("MainWindow", "CSV Files :"))
        self.btn_load.setText(_translate("MainWindow", "Load"))
        self.label_2.setText(_translate("MainWindow", "Volt :"))
        self.label_3.setText(_translate("MainWindow", " ~ "))
        self.btn_cal.setText(_translate("MainWindow", "Calculate"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())