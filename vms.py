
from tkinter import *
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import xlsxwriter
import os

def selectFiles(event=None):
			all_data = pd.DataFrame()
	  	#open window to select multiple vms files from SharePoint/R4
			filenames =  filedialog.askopenfilenames(initialdir = "C:/Users/"+userID+"/Documents/SharePoint/R4 SESD FSB Ecology Section - Documents/Staff Weekly Reports/",title = "Select files",filetypes = (("xlsx","*.xlsx"),("all files","*.*"))) 
			
			#if cancel clicked do nothing
			if not filenames:
				return
							
			
			for f in filenames:
				df = pd.read_excel(f)
				all_data = all_data.append(df)
				all_data.iloc[:, 6:9] = all_data.iloc[:, 6:9].apply(pd.to_datetime, errors='coerce')
		
			
			
			#prompt for save location and name
			vmsFilename = filedialog.asksaveasfilename(initialdir = "C:/Users/"+userID+"/Documents/SharePoint/R4 SESD FSB Ecology Section - Documents/Section Weekly Report/",title = "Save as",filetypes = (("xlsx","*.xlsx"),("all files","*.*")))
			
			#if cancel clicked do nothing
			if not vmsFilename:
				return
			
			#create the writer for excel
			writer = pd.ExcelWriter(vmsFilename+".xlsx", engine='xlsxwriter')
			#make the excel file
			all_data.to_excel(writer, sheet_name='ESS Section')
			writer.save()
			messagebox.showinfo("Complete", "vms saved: " +vmsFilename+".xlsx")
			
			
			
			
			
		
root = Tk()
root.title("VMS Magic Box5000")

w = 300 # width for the Tk root
h = 50 # height for the Tk root

# get screen width and height
ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

# set the dimensions of the screen 
# and where it is placed
root.geometry('%dx%d+%d+%d' % (w, h, x, y))

userID = os.getlogin()

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

#
ttk.Button(mainframe, text="Select Files", command=selectFiles).grid(column=1, row=2, sticky=W)


for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)



root.bind('<Return>', selectFiles)

root.mainloop()
