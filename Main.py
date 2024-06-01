import tkinter
from docx import Document
import docxedit

main_window = tkinter.Tk()
print("hehe haha")

template_path = "Base_document.docx"
output_path = 'out/Output_fun.docx'

doc = Document(template_path)
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            docxedit.replace_string(cell,"{{NAME}}","Masztur Bálint")
doc.save(output_path)

# main_window.mainloop()