import sqlite3

class UsersDB:
    def __init__(self):    
        self.db =sqlite3.connect("bin/users_db.db")
        self.cur = self.db.cursor()
    def SearchByName(self,nameToSearch):
        results = self.cur.execute("SELECT * FROM Users where name="+nameToSearch+"")
        return results.fetchall()
    def InsertRecord(self,name_,address_,pnumber_):
        self.cur.execute("INSERT into Users (name,Address,Pnumber) VALUES ('"+name_+"','"+address_+"','"+pnumber_+"');")
        self.db.commit()
    def GetDocumentID(self):
        result = self.cur.execute("SELECT DocuID FROM DocumentID")
        return result.fetchone()

thing = UsersDB()
thing.InsertRecord("Másik Nber","Hajléktalan","06308775959")