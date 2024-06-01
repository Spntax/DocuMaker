from docx import Document
import docxedit
from datetime import datetime
class DocumentData:
    def __init__(self):
        self.name = "Masztur Bálint"
        self.address = "9700 Szombathely Szopok boré utca. 69/B"
        self.p_number = "0630-911/4200"
        self.notes = "Anyjádat xxxd"
        self.callsign = "az ott"
        self.type_data = "Mekláren"
        self.modell = "gyors"
        self.description = "nincs"
        self.addons = "az összes"
        self.diagnosis = "rosz"
        self.year = datetime.today().strftime('%Y')
        self.month = datetime.today().strftime('%m')
        self.day = datetime.today().strftime('%d')

    def UpdateData(self,doc):
#Iterating trough the document's tables cuz bitch ass docxedit can't replace strings in tables when parsing the whole doc
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