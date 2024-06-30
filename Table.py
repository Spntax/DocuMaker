# Python program to create a table

from tkinter import *


class Table:
	def __init__(self,root,list_data):
		total_rows = len(list_data)
		total_columns = len(list_data[0])
		total_columns = total_columns-2
		# code for creating table
		for j in range(total_columns):
			for i in range(total_rows):
				self.e = Entry(root, width=20, fg='blue',
						font=('Arial',10,'bold'))
				if(list_data[i][6]==1):
					self.e.configure({"disabledbackground": "yellow"})
				if(list_data[i][6]==2):
					self.e.configure({"disabledbackground": "lawngreen"})

				self.e.grid(row=i+23, column=j)
				self.e.insert(END, list_data[i][j+1])
				self.e.configure(state=DISABLED)