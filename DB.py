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
            returnVal = results.fetchall()
            if returnVal==[]:
                results = self.cur.execute("SELECT * FROM Users where name LIKE '%"+nameToSearch+"%'")
                return results.fetchall()
            else:
                return returnVal
        else:
            results = self.cur.execute("SELECT * FROM Users")
            return results.fetchall()
        
    def SearchByDocID(self,DocID_):
        results = self.cur.execute("SELECT * FROM Users where DocID= '"+str(DocID_)+"'")
        return results.fetchall()
    
    def GetByUserID(self,id_):
            results = self.cur.execute("SELECT * FROM Users where ID='"+str(id_)+"'")
            return results.fetchone()
    
    def GetWorkTable(self,nameToSearch):
        if(nameToSearch!=""):
            results = self.cur.execute("SELECT name,address FROM Users where name='"+nameToSearch+"'")
            return results.fetchall()
        else:
            results = self.cur.execute("SELECT ID,name,Address,Pnumber,TypeData,Comment,JobState FROM Users")
            return results.fetchall()
#Comment, JobState

    def GetDocumentID(self,docCounter):
        result = self.cur.execute("SELECT DocuID FROM DocumentID")
        currentDocID = result.fetchone()[0]
        if docCounter:
            self.cur.execute("UPDATE DocumentID SET DocuID = "+str((currentDocID+1))+" WHERE ID = 1")
            self.db.commit()
        return currentDocID
    
    def GetName(self,name_):
        if(name_!=""):
            results = self.cur.execute("SELECT name FROM Users WHERE name LIKE '%"+name_+"%'")
            return results.fetchone()[0]
        else:
            results = self.cur.execute("SELECT name FROM Users")
            return results.fetchone()[0]
        
    def GetComment(self,ID_):
            results = self.cur.execute("SELECT Comment FROM Users WHERE ID = "+str(ID_)+"")
            return results.fetchone()[0]
        
    def InsertRecord(self,tk_name,tk_address,tk_phone,tk_note,tk_callsign,tk_type,tk_modell,tk_description,tk_addon,tk_diagnosis):
        self.cur.execute("""INSERT into Users
                         (name,Address,Pnumber,Callsign,Comment,JobState,Notes,TypeData,Modell,Description,Addons,Diagnosis,DocID) 
                         VALUES ('"""+tk_name+"""','"""+tk_address+"""','"""+tk_phone+"""','"""+tk_callsign+"""','','0','"""+tk_note+"""',
                         '"""+tk_type+"""','"""+tk_modell+"""','"""+tk_description+"""','"""+tk_addon+"""','"""+tk_diagnosis+"""','0');""")
        self.db.commit()

    def UpdateState(self,ID_,newState_):
        #newState = str(newState_)
        self.cur.execute("UPDATE Users SET JobState = "+str(newState_)+" WHERE ID = "+str(ID_)+"")
        self.db.commit()

    def UpdateComment(self,ID_,newComment_):
        #newState = str(newState_)
        self.cur.execute("UPDATE Users SET Comment = '"+str(newComment_)+"' WHERE ID = "+str(ID_)+"")
        self.db.commit()

    def UpdateDocID(self,newID_):
        #newState = str(newState_)
        self.cur.execute("UPDATE DocumentID SET DocuID = "+str(newID_)+" WHERE ID = 1")
        self.db.commit()
    def UpdateRecordDocID(self,UserID_,DocID_):
        #newState = str(newState_)
        self.cur.execute("UPDATE Users SET DocID = "+str(DocID_)+" WHERE ID = "+str(UserID_)+"")
        self.db.commit()

#thing = UsersDB()
#thing.InsertRecord("Másik Nber","Hajléktalan","06308775959")