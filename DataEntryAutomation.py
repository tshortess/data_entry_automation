#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pynput import mouse
from pynput import keyboard
from pynput.mouse import Button, Controller as MouseController, Listener
from pynput.keyboard import Key, Controller as KeyboardController
import os.path
from glob import glob
import time
import pandas as pd
import datetime


# In[2]:


currentDirectory = os.getcwd()

csvExtension = ".csv"
xlsxExtension = ".xlsx"
xlsExtension = ".xls"
excelFilesInDirectory = glob('*{}'.format(csvExtension)) + glob('*{}'.format(xlsxExtension)) + glob('*{}'.format(xlsExtension))


# In[3]:


isRequestValid = False
if len(excelFilesInDirectory) == 0:
    print(f"You don't have any Excel files in the current directory of: {currentDirectory}")
    while not(isRequestValid):
        filePath = input("Please enter the full path (without quotes) to the Excel file you wish to automate. \nInput:")
        try:
            pathExists = os.path.exists(filePath)
        except:
            pathExists = False
        isExcelFile = filePath.endswith(csvExtension) or filePath.endswith(xlsxExtension) or filePath.endswith(xlsExtension)
        if pathExists and isExcelFile:
            try:
                dataFrame = pd.read_excel(filePath)
                isRequestValid = True
            except Exception as ex:
                print(ex)
                print("File is not valid. Try another file.")
                continue
else:
    print(f"You have the following Excel files in the current directory of: {currentDirectory}")
    for file in excelFilesInDirectory:
        print(f'{excelFilesInDirectory.index(file)} - {file}')
    while not(isRequestValid):
        requestedFile = input("Enter the number of the file you wish to automate. \nIf you want to enter an Excel file in a different path, enter that file's full path without quotes. \nInput: ")
        if requestedFile.isdigit():
            if int(requestedFile) > (len(excelFilesInDirectory) - 1) or int(requestedFile) < 0:
                continue
            else: 
                filePath = currentDirectory + "/" + excelFilesInDirectory[int(requestedFile)]
                try:
                    dataFrame = pd.read_excel(filePath)
                    isRequestValid = True
                except Exception as ex:
                    print(ex)
                    print("File is not valid. Try another file.")
                    continue
        else:
            try:
                pathExists = os.path.exists(requestedFile)
            except:
                pathExists = False
            isExcelFile = requestedFile.endswith(csvExtension) or requestedFile.endswith(xlsxExtension) or requestedFile.endswith(xlsExtension)
            if pathExists and isExcelFile:
                filePath = requestedFile
                try:
                    dataFrame = pd.read_excel(filePath)
                    isRequestValid = True
                except:
                    print("File is not valid. Try another file.")
                    continue


# In[ ]:


print("Move your cursor to the spot where you want to start entering data and then press the left Shift key at that spot.")

autonomous_mouse = MouseController()

with keyboard.Events() as events:
    for event in events:
        if event.key == keyboard.Key.shift:
            x_axis, y_axis = autonomous_mouse.position
            break


# In[ ]:


autonomous_mouse.position = (x_axis, y_axis)

autonomous_mouse.press(Button.left)
autonomous_mouse.release(Button.left)

autonomous_keyboard = KeyboardController()

print("Beginning the process of data entry automation...")
time.sleep(2)

for i, r in dataFrame.iterrows():
    for key in r:
        if isinstance(key, datetime.datetime):
            key = key.strftime("%m/%d/%Y")
        key = str(key)
        for x in key:
            autonomous_keyboard.press(x)
            autonomous_keyboard.release(x)
        time.sleep(.05)
        autonomous_keyboard.press(Key.tab)
        autonomous_keyboard.release(Key.tab)
    autonomous_keyboard.press(Key.enter)
    autonomous_keyboard.release(Key.enter)
    y_axis += 21
    autonomous_mouse.position = (x_axis, y_axis)
    time.sleep(.4)
    autonomous_mouse.press(Button.left)
    autonomous_mouse.release(Button.left)
    #time.sleep(1.25)


# # 
