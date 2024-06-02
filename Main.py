from tkinter import *
from docx import Document
import docxedit
import DocData
import os
import subprocess

print("hehe haha")

# Global variables ˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇ
main_window = Tk()

template_path = "Base_document.docx"
output_path = 'out/Output_fun.docx'

doc = Document(template_path)

documentData = DocData.DocumentData()
# Global variables ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

def ShowDocument():
    SaveDocument()
    subprocess.Popen([output_path],shell=True)
    os.startfile(output_path)

def SaveDocument():
    doc = Document(template_path)
# Update the Docuemnt class with the data filled in the main window
    documentData.name = tk_name.get()
    documentData.address = tk_address.get()
    documentData.p_number = tk_phone.get()
    documentData.notes = tk_note.get()
    documentData.callsign = tk_callsign.get()
    documentData.type_data = tk_type.get()
    documentData.modell = tk_modell.get()
    documentData.description = tk_description.get()
    documentData.addons = tk_addon.get()
    documentData.diagnosis = tk_diagnosis.get()



# Update the template Docx file and replace the data
    documentData.UpdateData(doc)
    print(tk_name.get())
    print(documentData.name)
    doc.save(output_path)

#Main window layout definition
Label(main_window, text='Megrendelő Neve:').grid(row=0, column=0)
tk_name = Entry(main_window)
tk_name.grid(row=0, column=1)
Label(main_window, text='Cím').grid(row=1)
tk_address = Entry(main_window)
tk_address.grid(row=1,column=1)
Label(main_window, text='Telefon').grid(row=2)
tk_phone = Entry(main_window)
tk_phone.grid(row=2,column=1)
Label(main_window, text='Megjegyzés').grid(row=3)
tk_note = Entry(main_window)
tk_note.grid(row=3,column=1)
Label(main_window, text='Megnevezés').grid(row=0,column=2)
tk_callsign = Entry(main_window)
tk_callsign.grid(row=0,column=3)
Label(main_window, text='Típus').grid(row=1,column=2)
tk_type = Entry(main_window)
tk_type.grid(row=1,column=3)
Label(main_window, text='Modell').grid(row=2,column=2)
tk_modell = Entry(main_window)
tk_modell.grid(row=2,column=3)
# Blank line filler
Label(main_window, text='').grid(row=4,column=0)
Label(main_window, text='Hibajelenség').grid(row=5,column=0)
tk_description = Entry(main_window)
tk_description.grid(row=5,column=1)
Label(main_window, text='Tartozékok').grid(row=6,column=0)
tk_addon = Entry(main_window)
tk_addon.grid(row=6,column=1)
Label(main_window, text='Diagnózis').grid(row=5,column=2)
tk_diagnosis = Entry(main_window)
tk_diagnosis.grid(row=5,column=3)

Button(main_window, text='Save', width=15, command=SaveDocument).grid(row=20)
Button(main_window, text='Show', width=15, command=SaveDocument).grid(row=20, column=2)

#Main loop start here
main_window.mainloop()