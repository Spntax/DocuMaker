# Python program to create a table

from tkinter import *

class Table:
	def __init__(self,root,list_data):
		total_rows = len(list_data)
		total_columns = len(list_data[0])
		total_columns = total_columns-1
		# code for creating table
		for j in range(total_columns):
			for i in range(total_rows):
				#If item is not "removed", green is not removed it is just job complete
				if(list_data[i][4]!=3):
					self.m = Menu(root, tearoff=0)
					self.m.add_command(Label="Komment")
					self.m.add_command(Label="Állapot 1")
					self.m.add_command(Label="Állapot 2")
					self.e = Entry(root, width=20, fg='blue',
							font=('Arial',11))
					if(list_data[i][4]==1):
						self.e.configure({"disabledbackground": "yellow"})
					if(list_data[i][4]==2):
						self.e.configure({"disabledbackground": "lawngreen"})

					self.e.grid(row=i+23, column=j)
					self.e.insert(END, list_data[i][j])
					self.e.configure(state=DISABLED)

	def Update(self,root,list_data):
		total_rows = len(list_data)
		total_columns = len(list_data[0])
		total_columns = total_columns-1
		if root.grid_size()[0]<total_rows+23:
		# code for creating table
			for j in range(total_columns):
				for i in range(total_rows):
					#If item is not "removed", green is not removed it is just job complete
					if(list_data[i][4]!=3):
						self.e = Entry(root, width=20, fg='blue',
								font=('Arial',11))
						if(list_data[i][4]==1):
							self.e.configure({"disabledbackground": "yellow"})
						if(list_data[i][4]==2):
							self.e.configure({"disabledbackground": "lawngreen"})

						self.e.grid(row=i+23, column=j)
						self.e.insert(END, list_data[i][j])
						self.e.configure(state=DISABLED)
