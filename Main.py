from tkinter import *
from tkinter import ttk 
from docx import Document
import docxedit
import DocData
import os
import subprocess
import DB
import Table
import time

print("hehe haha")

# Global variables ˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇ
main_window = Tk()

template_path = "Base_document.docx"
output_path = 'out/Output_fun.docx'

doc = Document(template_path)

documentData = DocData.DocumentData()
usersDB = DB.UsersDB()

list_data = usersDB.GetWorkTable("")

docSaved = False

# Global variables ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

def ShowDocument():
    SaveDocument()
    subprocess.Popen([output_path],shell=True)
    os.startfile(output_path)

def SaveDocument(tk_name,tk_address,tk_phone,tk_note,tk_callsign,tk_type,tk_modell,tk_description,tk_addon,tk_diagnosis):
    list_data = usersDB.GetWorkTable("")
    global docSaved
    doc = Document(template_path)
# Update the Docuemnt class with the data filled in the main window
    documentData.documentID = usersDB.GetDocumentID()
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

#Add Record to Database
    usersDB.InsertRecord(tk_name.get(),tk_address.get(),tk_phone.get(),tk_note.get(),tk_callsign.get(),tk_type.get(),tk_modell.get(),tk_description.get(),tk_addon.get(),tk_diagnosis.get())
#    bbb = Entry(main_window,width=20)
#    bbb.grid(row=main_window.grid_size()[1])
#    bbb.insert(END, "fkjghdfkl")

    docSaved = True
    print(docSaved)
    
def CreateWorkItemTable(root,list_data):
    total_rows = len(list_data)
    total_columns = len(list_data[0])
    total_columns = total_columns-1
    # code for creating table
    for j in range(total_columns):
        for i in range(total_rows):
            #If item is not "removed", green is not removed it is just job complete
            if(list_data[i][4]!=3):
                m = Menu(root, tearoff=0)
                m.add_command(label="Komment")
                m.add_command(label="Állapot 1")
                m.add_command(label="Állapot 2")
                e = Entry(root, width=20, fg='blue',
                        font=('Arial',11))
                if(list_data[i][4]==1):
                    e.configure({"disabledbackground": "yellow"})
                if(list_data[i][4]==2):
                    e.configure({"disabledbackground": "lawngreen"})

                e.grid(row=i+23, column=j)
                e.insert(END, list_data[i][j])
                e.configure(state=DISABLED)


#Main window layout definition
#selected = StringVar()
#dasCOmbobox = ttk.Combobox(main_window,textvariable=selected)
#dasCOmbobox.grid(row='100',column='100')
def DrawMainWindow():
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

    Button(main_window, text='Mentés', width=15, command=lambda: SaveDocument(tk_name,tk_address,tk_phone,tk_note,tk_callsign,tk_type,tk_modell,tk_description,tk_addon,tk_diagnosis)).grid(row=19,columnspan=4,pady=10)


    # ----------------------------- Workitem list --------------------------------
    sep = ttk.Separator(main_window, orient="horizontal").grid(row=21, columnspan=999,pady=15,padx=20,sticky="ew")
    Label(main_window,text="Név").grid(row=22,column=0)
    Label(main_window,text="Lakhely").grid(row=22,column=1)
    Label(main_window,text="Eszköz").grid(row=22,column=2)
    Label(main_window,text="Komment").grid(row=22,column=3)

    CreateWorkItemTable(main_window,list_data)

DrawMainWindow()
#Main loop start here
#main_window.mainloop()
while True:
    main_window.update()
    main_window.update_idletasks()
    #print(docSaved)
    if docSaved:
        docSaved = False
        main_window.destroy()
        main_window = Tk()
        list_data = usersDB.GetWorkTable("")
        print("yippie")
        DrawMainWindow()
        main_window.update()