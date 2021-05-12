import PySimpleGUI as sg
from docx import Document
import os
from components.events import def_path
from components.events import filtering
from components.events import filtering_orig
import tkinter as tk
from tkinter import filedialog

file_path = def_path()

#########################
def inter(file_path):
    #--------------------------------
    #---- building the interface ---- 
    #--------------------------------
    
    #----interface inputs
    i1 = [[sg.InputCombo(('arch1', 'arch2'), size=(13, 0), key = 'arch', default_value = "arch1")]] #list of all possible architetures
    i2 = [[sg.InputCombo(('var1', 'var2', 'var3'), size=(10, 0), key = 'variant', default_value = "var1")]] # list of all possible variants
    i3 = [[sg.InputCombo(('ecu1', 'ecu1', 'ecu3'), size=(10, 0), key = 'ECU', default_value = "ecu1")]] # list of all possible ECUs
    
    #intreface frames 
    fr1 = [[sg.Frame("Select the architeture", i1)]]
    fr2 = [[sg.Frame("Select the variant", i2)]]
    fr3 = [[sg.Frame("Select the ECU", i3)]]
    
    #layout building
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
            if file_path.endswith(".docm"):
                filtering(arch = values['arch'], variant = values['variant'], ECU = values['ECU'], file_path=file_path)
                sg.Popup('The file {} has been filtered'.format(file_path))    
            else:
                window.Element('fileaddress').Update(value=file_path)
                filtering_orig(arch = values['arch'], variant = values['variant'], ECU = values['ECU'])
                sg.Popup('All files has been filtered')
                

        if event == sg.WIN_CLOSED or event == 'Exit' or event == None:
            break
        
        if event == 'fileaddress':
            file_path = def_path()
            window.Element('fileaddress').Update(value=file_path)

def main():
    inter(file_path)