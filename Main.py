from tkinter import *
from docx import Document
import docxedit
import DocData
import os

print("hehe haha")

# Global variables ˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇ
main_window = Tk()

template_path = "Base_document.docx"
output_path = 'out/Output_fun.docx'

doc = Document(template_path)

documentData = DocData.DocumentData()
# Global variables ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

def SaveDocument():
    documentData.name = tk_name.get()
    documentData.address = tk_address.get()
    documentData.p_number = tk_phone.get()
    documentData.notes = tk_note.get()
    documentData.UpdateData(doc)
    doc.save(output_path)


#Main window layout definition
Label(main_window, text='Név').grid(row=0)
tk_name = Entry(main_window)
tk_name.grid(row=0, column=2)
Label(main_window, text='Cím').grid(row=1)
tk_address = Entry(main_window)
tk_address.grid(row=1,column=2)
Label(main_window, text='Telefon').grid(row=2)
tk_phone = Entry(main_window)
tk_phone.grid(row=2,column=2)
Label(main_window, text='Megjegyzés').grid(row=3)
tk_note = Entry(main_window)
tk_note.grid(row=3,column=2)
Label(main_window, text='Név').grid(row=4)
Entry(main_window).grid(row=4,column=2)


Button(main_window, text='Save', width=25, command=SaveDocument).grid(row=20)
Button(main_window, text='Show', width=25, command=SaveDocument).grid(row=20, column=2)

#Main loop start here
main_window.mainloop()