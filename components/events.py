from docx import Document
import os
import time
import tkinter as tk
from tkinter import filedialog

#removing rows from tables (lines in the file)
def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)
    
#choose file 
def def_path():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select file", filetypes=[("CFTS Files", "*.docm")])
    if file_path.endswith(".docm"):
        return os.path.abspath(file_path)
    else:
        return 'All .docm files in root folder will be filtered'

#filtering and saving document
def filtering(arch, variant, ECU, file_path):
    file = (os.path.basename(file_path))
#if __name__ == "__main__":
    #for file in os.listdir(os.getcwd()):
    if file.endswith(".docm"):
        document = Document(file)         
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for index, paragraph in enumerate(cell.paragraphs):
                        if index == 0:                        
                            if ':]' in paragraph.text:
                                k=0
                                for c in paragraph.text:
                                    if ']' in c:
                                        k=k+1
                                if k == 1:                            
                                    if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) or variant in paragraph.text or ECU in paragraph.text or 'LATAM' in paragraph.text:
                                        pass
                                        #print('1- '+paragraph.text)                                    
                                    else:
                                        #print('1- delete - '+paragraph.text)
                                        remove_row(table, row)
                                if k == 2:                            
                                    if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and ECU in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and 'LATAM' in paragraph.text or variant in paragraph.text and ECU in paragraph.text or variant in paragraph.text and 'LATAM' in paragraph.text or ECU in paragraph.text and 'LATAM' in paragraph.text:
                                        pass
                                        #print('2 - '+paragraph.text)
                                    else:
                                        #print('2- delete - '+paragraph.text)
                                        remove_row(table, row)
                                if k == 3:                            
                                    if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text and ECU in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text and 'LATAM' in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and ECU in paragraph.text and 'LATAM' in paragraph.text or variant in paragraph.text and ECU in paragraph.text  and 'LATAM' in paragraph.text:
                                        pass
                                        #print('3 - '+paragraph.text)
                                    else:
                                        #print('3- delete - '+paragraph.text)
                                        remove_row(table, row)
                                if k == 4:                            
                                    if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text and ECU in paragraph.text and 'LATAM' in paragraph.text :
                                        pass
                                        #print('4 - '+paragraph.text)
                                    else:
                                        #print('4- delete - '+paragraph.text)
                                        remove_row(table, row)
        document.save('FILTRADO_'+file)

def filtering_orig(arch, variant, ECU):
    for file in os.listdir(os.getcwd()):
        if file.endswith(".docm"):
            document = Document(file)
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for index, paragraph in enumerate(cell.paragraphs):
                            if index == 0:                        
                                if ':]' in paragraph.text:
                                    k=0
                                    for c in paragraph.text:
                                        if ']' in c:
                                            k=k+1
                                    if k == 1:                            
                                        if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) or variant in paragraph.text or ECU in paragraph.text or 'LATAM' in paragraph.text:
                                            pass
                                            #print('1- '+paragraph.text)                                    
                                        else:
                                            #print('1- delete - '+paragraph.text)
                                            remove_row(table, row)
                                    if k == 2:                            
                                        if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and ECU in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and 'LATAM' in paragraph.text or variant in paragraph.text and ECU in paragraph.text or variant in paragraph.text and 'LATAM' in paragraph.text or ECU in paragraph.text and 'LATAM' in paragraph.text:
                                            pass
                                            #print('2 - '+paragraph.text)
                                        else:
                                            #print('2- delete - '+paragraph.text)
                                            remove_row(table, row)
                                    if k == 3:                            
                                        if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text and ECU in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text and 'LATAM' in paragraph.text or ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and ECU in paragraph.text and 'LATAM' in paragraph.text or variant in paragraph.text and ECU in paragraph.text  and 'LATAM' in paragraph.text:
                                            pass
                                            #print('3 - '+paragraph.text)
                                        else:
                                            #print('3- delete - '+paragraph.text)
                                            remove_row(table, row)
                                    if k == 4:                            
                                        if ('allSys: CTS1_2' in paragraph.text or arch in paragraph.text) and variant in paragraph.text and ECU in paragraph.text and 'LATAM' in paragraph.text :
                                            pass
                                            #print('4 - '+paragraph.text)
                                        else:
                                            #print('4- delete - '+paragraph.text)
                                            remove_row(table, row)
            document.save('FILTRADO_'+file)
    
