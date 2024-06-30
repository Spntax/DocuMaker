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
    def InsertRecord(self,name_,address_,pnumber_):
        self.cur.execute("INSERT into Users (name,Address,Pnumber) VALUES ('"+name_+"','"+address_+"','"+pnumber_+"');")
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