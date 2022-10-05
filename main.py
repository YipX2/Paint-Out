import threading

import PySimpleGUI as sg

import pandas as pd

import time


#Opens Excel,  may need to change engine and import VBA script.
workbook = pd.read_excel(r"C:\Users\Forge Design\Desktop\PaintLineFiles\PaintLine.ods",engine = 'odf')


# Opens a new thread, calls function saveto, once saveto is complete kills thread.
def savefile():
    t = threading.Timer(1,saveto())
    t.daemon = True
    t.start()
    t.cancel()


# Change to save file to excel as workbook. May need to call out VBA script as well.
def saveto():
    print(workbook)



def add_one(column):
    workbook.at[0, (column)] += 1
    workbook.at[0, 'Tally'] += 1



def sub_one(column):
    workbook.at[0, (column)] -= 1
    workbook.at[0, 'Tally'] -=1



#Currently no counters to show users how many panels have been run. May need to add something.
layout = [
         [sg.Button("Add Tan Back", size=(15,1)),  sg.Button("Subtract Tan Back", size=(15,1))],
         [sg.Button("Add Tan Strike", size=(15,1)),sg.Button("Subtract Tan Strike", size=(15,1))],
         [sg.Button("Add Green Back", size=(15,1)),sg.Button("Subtract Green Back", size=(15,1))],
         [sg.Button("Add Green Strike", size=(15,1)),sg.Button("Subtract Green Strike", size=(15,1))],
         [sg.Button("Rework", size=(15,1)), sg.Button("Rerun", size=(15,1))],
         [sg.Button("Exit", size=(15,1)) ],
         ]

# Create the window
window = sg.Window(title="Paint Tally",layout=layout, margins=(150, 150))


# Create an event loop.
while True:
    event, values = window.read(timeout= 1*1000)
    if event == sg.TIMEOUT_KEY:
        savefile()
    if event == "Add Tan Back":
        add_one("Tan Backs In")
    if event == "Subtract Tan Back":
        sub_one("Tan Backs In")
    if event == "Add Tan Strike":
        add_one('Tan Strikes In')
    if event == "Subtract Tan Strike":
        sub_one('Tan Strikes In')
    if event == "Add Green Back":
        add_one('Green Backs In')
    if event == "Subtract Green Back":
        sub_one('Green Backs In')
    if event == "Add Green Strike":
        add_one('Green Strikes In')
    if event == "Subtract Green Strike":
        sub_one('Green Strikes In')
    if event == "Rework":
        add_one('Rework')
    if event =="Rerun":
        workbook.at[0, 'Rerun'] += 1
    if event == "Exit":
        saveto()
        break


window.close()