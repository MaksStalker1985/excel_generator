# -*- coding: utf-8 -*-
"""
Created on Sat Jun  4 08:14:12 2022

@author: Maksym
"""

import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QMessageBox
import os
from pathlib import Path
import subprocess

df = 0
columns = 0
rows = 0
steps = 0
out_df = 0


def readExcel(file_name):
    global df
    global columns
    global rows
    try:
        if (os.path.isfile(file_name) == True):
            path = Path.absolute(file_name)
            df = pd.read_excel(path, header=None)
            rows = len(df)
            columns = len(df.columns)
            return True
    except:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("Error")
        msg.setInformativeText('Can not open file')
        msg.setWindowTitle("Error")
        msg.exec_()
        return False


def GenerateString():
    global rows
    global columns
    global df
    result = ''
    count = 0
    try:
        for count in range(0, columns):
            number = np.random.randint(0, rows)
            value = df.loc[number, count]
            result += value + ' '
            count += 1
        return result
    except:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("Error")
        msg.setInformativeText('Wrong quantities of rows and columns')
        msg.setWindowTitle("Error")
        msg.exec_()


def writeExcel():
    global steps
    strings = []
    i = 0
    while i in range(0, steps):
        strings.append(GenerateString())
        i += 1

    out_df = pd.DataFrame(strings)

    try:
        path = os.getcwd()

        i = path.find('main.app')
        new_path=''

        if i == -1 or i == 0:
            new_path = path
        else:
            path.split('.',1)[0]
            new_path = path[0:i-1:]

        new_path += "/output.xlsx"
        out_df.to_excel(new_path)
        file_to_show = new_path
        subprocess.call(["open", "-R", file_to_show])
    except:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("Error")
        msg.setInformativeText('Can not create ouptut file')
        msg.setWindowTitle("Error")
        msg.exec_()
