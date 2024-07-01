import sqlite3

class UsersDB:
    def __init__(self):    
        self.db =sqlite3.connect("bin/users_db.db")
        self.cur = self.db.cursor()
        results = self.cur.execute("SELECT name FROM Users")
        self.allNames = results.fetchall()
        self.filteredNames = self.allNames
    def SearchByName(self,nameToSearch):
        if(nameToSearch!=""):
            results = self.cur.execute("SELECT * FROM Users where name='"+nameToSearch+"'")
            return results.fetchall()
        else:
            results = self.cur.execute("SELECT * FROM Users")
            return results.fetchall()
    def GetWorkTable(self,nameToSearch):
        if(nameToSearch!=""):
            results = self.cur.execute("SELECT name,address FROM Users where name='"+nameToSearch+"'")
            return results.fetchall()
        else:
            results = self.cur.execute("SELECT name,Address,Callsign,Comment,JobState FROM Users")
            return results.fetchall()
#Comment, JobState
    def InsertRecord(self,tk_name,tk_address,tk_phone,tk_note,tk_callsign,tk_type,tk_modell,tk_description,tk_addon,tk_diagnosis):
        self.cur.execute("""INSERT into Users
                         (name,Address,Pnumber,Callsign,Comment,JobState,Notes,TypeData,Modell,Description,Addons,Diagnosis) 
                         VALUES ('"""+tk_name+"""','"""+tk_address+"""','"""+tk_phone+"""','"""+tk_callsign+"""','','','"""+tk_note+"""',
                         '"""+tk_type+"""','"""+tk_modell+"""','"""+tk_description+"""','"""+tk_addon+"""','"""+tk_diagnosis+"""');""")
        self.db.commit()
    def GetDocumentID(self):
        result = self.cur.execute("SELECT DocuID FROM DocumentID")
        return result.fetchone()[0]
    def GetName(self,name_):
        if(name_!=""):
            results = self.cur.execute("SELECT name FROM Users where name='"+name_+"'")
            return results.fetchall()
        else:
            results = self.cur.execute("SELECT name FROM Users")
            return results.fetchall()


#thing = UsersDB()
#thing.InsertRecord("Másik Nber","Hajléktalan","06308775959")