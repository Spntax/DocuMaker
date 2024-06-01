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

documentData = DocData.DocumentData
# Global variables ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

#Iterating trough the document's tables cuz bitch ass docxedit can't replace strings in tables when parsing the whole doc
def SaveDocument():
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                docxedit.replace_string(cell,"{{NAME}}",documentData.name)
                docxedit.replace_string(cell,"{{ADDRESS}}",documentData.address)
                docxedit.replace_string(cell,"{{P_NUMBER}}",documentData.p_number)
                docxedit.replace_string(cell,"{{NOTES}}",documentData.notes)
                docxedit.replace_string(cell,"{{CALLSIGN}}",documentData.callsign)
                docxedit.replace_string(cell,"{{TYPE}}",documentData.type_data)
                docxedit.replace_string(cell,"{{MODELL}}",documentData.modell)
                docxedit.replace_string(cell,"{{DESCRIPTION}}",documentData.description)
                docxedit.replace_string(cell,"{{ADDONS}}",documentData.addons)
                docxedit.replace_string(cell,"{{DIAGNOSIS}}",documentData.diagnosis)
                docxedit.replace_string(cell,"{{YEAR}}",documentData.year)
                docxedit.replace_string(cell,"{{MONTH}}",documentData.month)
                docxedit.replace_string(cell,"{{DAY}}",documentData.day)
    doc.save(output_path)


#Main window layout definition
Label(main_window, text='First Name').grid(row=0)
Button(main_window, text='Save', width=25, command=SaveDocument).grid(row=2)

#Main loop start here
main_window.mainloop()