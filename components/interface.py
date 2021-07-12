import PySimpleGUI as sg
from docx import Document
import os
from components.events import def_path
from components.events import filtering
from components.events import filtering_orig
from components.events import ko, removevalue, clearlist
import tkinter as tk
from tkinter import filedialog

file_path = def_path()
lista = []

def inter(file_path):
    #--------------------------------
    #---- building the interface ---- 
    #--------------------------------
    
    #----interface inputs
    i1 = [[sg.InputCombo(('AtlMi', 'AtlHi'), size=(13, 0), key = 'arch', default_value = "AtlMi")]] #list of all possible architetures
    i2 = [[sg.InputCombo(('R1L', 'R1M', 'R1H'), size=(10, 0), key = 'variant', default_value = "R1L")]] # list of all possible variants
    i3 = [[sg.InputCombo(('LTM', 'ETM', 'RRM'), size=(10, 0), key = 'ECU', default_value = "LTM")]] # list of all possible ECUs
    
    #intreface frames 
    fr1 = [[sg.Frame("Select the architeture", i1)]] # 1st frame, with the architetures options
    fr2 = [[sg.Frame("Select the variant", i2)]] # 2nd frame, with the variants options
    fr3 = [[sg.Frame("Select the ECU", i3)]] # 3rd frame, with the ECUs options
    fr4 = [[sg.Button('<<', key = 'previous'),sg.Button('>>', key = 'next')], # 4th frame, with buttons, checkbox and listbox
           [sg.Text('KO'), sg.Checkbox('', key = 'KO', enable_events = True)],
           [sg.Text('% KO:'), sg.Text('000', key = 'porcko')],
           [sg.Listbox(values = lista, size=(10,5), key = 'listbox', enable_events=True)]]

    #layout building
    layout = [[sg.Text('',key='fileaddress', size=(60,0), enable_events = True)],
                [sg.Frame('', fr1, relief="flat"), sg.Frame('', fr2, relief="flat"), sg.Frame('', fr3, relief="flat"), sg.Button('OK', key = 'ok')],
                [sg.Multiline("", key = "text", size = (60, 15)),sg.Frame('', fr4, relief="flat")],
                [sg.Input("0", key = 'current', size=(3,0)), sg.Text("/"),sg.Text("000", key = 'number'), sg.Button('GO', key = 'go')],
                [sg.Text(" "*50, key = 'lista')]]

    window = sg.Window('CFTS FILTERING').Layout(layout) # window load
    event, values = window.Read(timeout=100) # update of the window
    window.Element('fileaddress').Update(value=file_path) # updating the 'fileaddress' element
    
    while True:
        event, values = window.Read(timeout=100)
        window.Element('fileaddress').Update(value=file_path)
        
        if event == 'ok':
            if file_path.endswith(".docm"):
                window['listbox'](values='')
                d = filtering(arch = values['arch'], variant = values['variant'], ECU = values['ECU'], file_path=file_path)
                window.Element('number').Update(value=len(d.keys()))
                window.Element('text').Update(value=d['1']['title'])
                #print(d)
                values['current'] = 1
                window.Element('current').Update(value=1)
                window.Element('porcko').Update(value='0')
                window.Element('KO').Update(value=False)
                list = ko ('ct','1')
                sg.Popup('The file {} has been filtered'.format(file_path))
            else:
                window.Element('fileaddress').Update(value=file_path)
                filtering_orig(arch = values['arch'], variant = values['variant'], ECU = values['ECU'])
                #window.Element('text').Update(value=d['1']['title'])
                #values['current'] = 1
                #window.Element('current').Update(value=1)
                sg.Popup('All files has been filtered')
               
        if event == sg.WIN_CLOSED or event == 'Exit' or event == None:
            break
        
        if event == 'fileaddress':
            file_path = def_path()
            window.Element('fileaddress').Update(value=file_path)
         
        if event == 'previous':
            ct = int(values['current'])
            if ct > 1:
                ct = int(values['current']) - 1
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[str(ct)]['title'])
                if d[str(ct)]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                else:
                    window.Element('KO').Update(value=False)

        if event == 'next':
            ct = int(values['current'])
            if ct < len(d.keys()):
                ct = int(values['current']) + 1
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[str(ct)]['title'])
                if d[str(ct)]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                elif d[str(ct)]['sts'] == 'OK':
                    window.Element('KO').Update(value=False)
        
        if event == 'go':
            try:
                ct = int(values['current'])
                window.Element('text').Update(value=d[str(ct)]['title'])
                if d[str(ct)]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                elif d[str(ct)]['sts'] == 'OK':
                    window.Element('KO').Update(value=False)
            except:
                pass

        if event == 'KO':
            ct = int(values['current']) # assign the current value to variable ct
            try:
                if d[str(ct)]['sts'] == 'OK': # comparing the value of the dict
                    d[str(ct)]['sts'] = 'KO' # changing the value of the dict for the specific spec
                    window.Element('KO').Update(value=True) # updating the element 'KO' (checkbox)
                    list = ko (ct, '0') # call the function 'list' to return the updated list of KOs
                    window.Element('porcko').Update(value=round(100*(len(list)/len(d.keys())))) # update the element 'porcko' with the percentage of the KOs 
                elif d[str(ct)]['sts'] == 'KO':
                    d[str(ct)]['sts'] = 'OK' 
                    window.Element('KO').Update(value=False)
                    list = removevalue(ct)
                    window.Element('porcko').Update(value=round(100*(len(list)/len(d.keys()))))
                #print(list)
                window['listbox'](values=list)
            except:
                pass
        
        if event == 'listbox':
            try:
                ct = int(str(values['listbox']).replace('[','').replace(']',''))
                values['current'] = ct
                window.Element('current').Update(value=values['current'])
                window.Element('text').Update(value=d[str(ct)]['title'])
                if d[str(ct)]['sts'] == 'KO':
                    window.Element('KO').Update(value=True)
                elif d[str(ct)]['sts'] == 'OK':
                    window.Element('KO').Update(value=False)
            except:
                pass
            
def main():
    inter(file_path)