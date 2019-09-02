import sys
import os
from PyQt5 import QtCore, QtWidgets, QtGui, uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import getpass
import pandas as pd
import time
import csv
import win32api
import sys
from QRGen import generateDoc
def openFileNameDialog():
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    fileName, _ = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()", "","All Files (*);;Python Files (*.py)", options=options)
    if fileName:
        win.dbPath.setText(fileName)
        print(win.dbPath.text()) # Development

def openDirectoryDialog():
    file = str(QFileDialog.getExistingDirectory(None, "Select Directory"))
    if file:
        win.attachmentPath.setText(file)
        print(win.attachmentPath.text())            # Development
def generate():
        if(len(win.dbPath.text()) == 0):
                print("You need to select a database file (Excel)")
        elif(len(win.attachmentPath.text()) == 0):
                print("You need to select the attachment folder path")
        else:
                print("Fuzzy: {}, Face Recognition: {}".format(win.fuzzy.isChecked(), win.facecrop.isChecked()))  # DEBUGGING PURPOSES
                generateDoc(win.dbPath.text(), win.attachmentPath.text(), win.fuzzy.isChecked(), win.facecrop.isChecked())




app = QtWidgets.QApplication([])
 
win = uic.loadUi("Draft.ui") #specify the location of your .ui file

win.dbBtn.clicked.connect(lambda: openFileNameDialog())
win.dirBtn.clicked.connect(lambda: openDirectoryDialog())
win.generate.clicked.connect(lambda: generate())
win.show()
sys.exit(app.exec())

 