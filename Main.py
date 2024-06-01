from tkinter import *
from docx import Document
import docxedit
import DocData

print("hehe haha")

# Global variables ˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇ
main_window = Tk()

template_path = "Base_document.docx"
output_path = 'out/Output_fun.docx'

doc = Document(template_path)

documentData = DocData.DocumentData()
# Global variables ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

def SaveDocument():
    documentData.UpdateData(doc)
    doc.save(output_path)


#Main window layout definition
Label(main_window, text='First Name').grid(row=0)
Button(main_window, text='Save', width=25, command=SaveDocument).grid(row=2)

#Main loop start here
main_window.mainloop()