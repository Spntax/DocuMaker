from docx import Document
import docxedit
from datetime import datetime
from spire.doc import *
from spire.doc.common import *
class DocumentData:
    def __init__(self):
        self.documentID = "0"
        self.name = "Name Template"
        self.address = "9999 Hupikékfalva, Törp utca 69"
        self.p_number = "0630-123-4567"
        self.notes = "Nincs megjegyzés"
        self.callsign = "Nincs megnevezés"
        self.type_data = "Nincs típus"
        self.modell = "Nincs Modell"
        self.description = "Nincs leírás"
        self.addons = "Nincsenek tartozékok"
        self.diagnosis = "Nincs diagnózis"
        self.year = datetime.today().strftime('%Y')
        self.month = datetime.today().strftime('%m')
        self.day = datetime.today().strftime('%d')

    def UpdateData(self,doc):
#Iterating trough the document's tables cuz bitch ass docxedit can't replace strings in tables when parsing the whole doc
        docxedit.replace_string(doc,"{{DOC_ID}}",self.documentID)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    docxedit.replace_string(cell,"{{NAME}}",self.name)
                    docxedit.replace_string(cell,"{{ADDRESS}}",self.address)
                    docxedit.replace_string(cell,"{{P_NUMBER}}",self.p_number)
                    docxedit.replace_string(cell,"{{NOTES}}",self.notes)
                    docxedit.replace_string(cell,"{{CALLSIGN}}",self.callsign)
                    docxedit.replace_string(cell,"{{TYPE}}",self.type_data)
                    docxedit.replace_string(cell,"{{MODELL}}",self.modell)
                    docxedit.replace_string(cell,"{{DESCRIPTION}}",self.description)
                    docxedit.replace_string(cell,"{{ADDONS}}",self.addons)
                    docxedit.replace_string(cell,"{{DIAGNOSIS}}",self.diagnosis)
                    docxedit.replace_string(cell,"{{YEAR}}",self.year)
                    docxedit.replace_string(cell,"{{MONTH}}",self.month)
                    docxedit.replace_string(cell,"{{DAY}}",self.day)
    
    def ExtractDataFromDB(self,dbList,docID):
        self.documentID = docID
        tmp = dbList[1]
        self.name = tmp
        tmp = dbList[2]
        self.address = tmp
        tmp = dbList[3]
        self.p_number = tmp
        tmp = dbList[4]
        self.callsign = tmp
        tmp = dbList[7]
        self.notes = tmp
        tmp = dbList[8]
        self.type_data = tmp
        tmp = dbList[9]
        self.modell = tmp
        tmp = dbList[10]
        self.description = tmp
        tmp = dbList[11]
        self.addons = tmp
        tmp = dbList[12]
        self.diagnosis = tmp

    def ConvertToPdf(self):
        # Create a Document object
        document = Document()
        # Load a Word DOCX file
        document.LoadFromFile("bin/Output_fun.docx")
        # Or load a Word DOC file
        #document.LoadFromFile("Sample.doc")

        # Save the file to a PDF file
        document.SaveToFile("bin/WordToPdf.pdf", FileFormat.PDF)
        document.Close()