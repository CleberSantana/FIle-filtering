import PySimpleGUI as sg
from docx import Document
import os
import events
import tkinter as tk
from tkinter import filedialog

file_path = events.def_path()
#########################

#----interface building
i1 = [[sg.InputCombo(('AtlMi', 'AtlHi'), size=(13, 0), key = 'arch')]]
i2 = [[sg.InputCombo(('R1L', 'R1M', 'R1H'), size=(10, 0), key = 'variant')]]
i3 = [[sg.InputCombo(('LTM', 'ETM', 'RRM'), size=(10, 0), key = 'ECU')]]

fr1 = [[sg.Frame("Select the architeture", i1)]]
fr2 = [[sg.Frame("Select the variant", i2)]]
fr3 = [[sg.Frame("Select the ECU", i3)]]

layout = [[sg.Text('',key='fileaddress', size=(60,0), enable_events = True)],
            [sg.Frame('', fr1, relief="flat"), sg.Frame('', fr2, relief="flat"), sg.Frame('', fr3, relief="flat")],
            [sg.Text(" "*92),sg.Button('OK', key = 'ok')]]

window = sg.Window('CFTS FILTERING').Layout(layout)
event, values = window.Read(timeout=100)
window.Element('fileaddress').Update(value=file_path)

while True:
    event, values = window.Read()
    window.Element('fileaddress').Update(value=file_path)

    if event == 'ok':
        events.filtering(arch = values['arch'], variant = values['variant'], ECU = values['ECU'], file_path=file_path)

    if event == sg.WIN_CLOSED or event == 'Exit':
        break
        
    if event == 'fileaddress':
        file_path = events.def_path()
        window.Element('fileaddress').Update(value=file_path)
