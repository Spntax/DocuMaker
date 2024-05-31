import tkinter
import textract
import aspose.words as aw
main_window = tkinter.Tk()
print("hehe haha")

template_path = "Base_document.doc"
output_path = 'out/Output_fun.doc'

doc = aw.Document(template_path)
doc.save(output_path)
#text = textract.process("Base_document.doc")

# main_window.mainloop()