import tkinter
from tkinter import *
from tkinter import filedialog
import xlrd
import threading
from threading import Thread

media_filename = ''

def selectInputFile():
	global media_filename
	media_filename = ''
	media_filename = filedialog.askopenfilename(title = "Select Excel File",filetypes = [("Excel files", "*.xlsx *.xls")])
	if(len(media_filename)>0):
		fileSplitArray = media_filename.split('/')
		statusText.configure(text='Input file selected: '+str(fileSplitArray[len(fileSplitArray)-1]))

def execution(inputFile, outputFile, table):
	statusText.configure(text='Process Started')
	numberarray = []
	myfile = xlrd.open_workbook(inputFile)
	mysheet = myfile.sheet_by_index(0)
	headcounter=0
	headvalue=0
	for heading in mysheet.row_values(0):
		heading = heading.replace(' ','').replace('\t','').replace('\n','')
		if(heading==table):
			headvalue = headcounter
		headcounter = headcounter+1
	for rownum in range(mysheet.nrows):
		fullrow = (mysheet.row_values(rownum))
		x = fullrow[headvalue]
		try:
			x = int(x)
		except ValueError:
			x = x
		# x = str(x).replace(' ','')
		x = str(x).replace(' ','')
		numberarray.append(x)
	with open(outputFile, 'w') as writer:
		for data in numberarray:
			writer.write(str(data)+'\n')
	statusText.configure(text='Process completed')

def startProgram():
	outputName = outputFileInput.get('1.0', END)
	outputName = str(outputName).strip('\n')
	tableName = tableNameInput.get('1.0', END)
	tableName = str(tableName).strip('\n')
	if(len(media_filename)>0):
		if(len(outputName)>0):
			if(outputName[-3:]=='txt'):
				outputName = outputName
			else:
				outputName = str(outputName)+'.txt'
			if(len(tableName)>0):
				Thread(target = execution, args = (media_filename, outputName, tableName)).start()
			else:
				statusText.configure(text='Table name not defined')
		else:
			statusText.configure(text='Output file not defined')
	else:
		statusText.configure(text='Input file not selected')

App = tkinter.Tk()
App.resizable(False, False)
App.title('Excel to Text')
App.configure(background='#FFFFFF')

headerCanvas = Canvas(App,bg='#FFFFFF',bd=0,highlightthickness=0,relief='ridge')
headerCanvas.grid(row=0,column=0,columnspan=3, rowspan=1)
headerText = Label(headerCanvas, text='Excel to Text', font='Quicksand 13 bold', bg='#FFF', fg='#444')
headerText.pack(side=TOP, anchor=N, fill=BOTH, expand=1, ipady=10, pady=(5,10))



inputCanvas = Canvas(App,bg='#FFFFFF',bd=0,highlightthickness=0,relief='ridge')
inputCanvas.grid(row=1,column=0,rowspan=2, columnspan=1)
inputButton = Button(inputCanvas, text='input', bg='#006699', height=1, width=4, bd=0, font='Quicksand 9 bold', highlightthickness=0, relief='ridge', justify='center', activebackground='#E1E1E1', highlightbackground='#CDCDCD', fg='#FFF', highlightcolor='#006699', command=selectInputFile)
inputButton.pack(fill=BOTH, expand=1, ipady=3, padx=20, pady=(23,0))


outputFileTextCanvas = Canvas(App,bg='#FFFFFF',bd=0,highlightthickness=0,relief='ridge')
outputFileTextCanvas.grid(row=1,column=1,columnspan=2, sticky=S)
outputFileText = Label(outputFileTextCanvas, text='Enter output file name', font='Quicksand 9 bold', bg='#FFF', fg='#888')
outputFileText.pack(side=BOTTOM, anchor=S, pady=(0,4))

outputFileInputCanvas = Canvas(App,bg='#FFFFFF',bd=0,highlightthickness=0,relief='ridge')
outputFileInputCanvas.grid(row=2,column=1,columnspan=2, rowspan=1)
outputFileInput = Text(outputFileInputCanvas, font='Quicksand 9 bold', bg='#FFF', height=1, width=33, bd=0, relief='ridge', highlightthickness=1, fg='#333', highlightcolor='#069', selectbackground='#1b8a73', selectforeground='#FFFFFF', highlightbackground='#096', spacing1=5, spacing3=5)
outputFileInput.pack(fill=BOTH, padx=10, ipady=0)

tableNameInputTextCanvas = Canvas(App,bg='#FFFFFF',bd=0,highlightthickness=0,relief='ridge')
tableNameInputTextCanvas.grid(row=3,column=0,columnspan=2, rowspan=1, sticky=S)
tableNameInputText = Label(tableNameInputTextCanvas, text='Enter table name', font='Quicksand 9 bold', bg='#FFF', fg='#888')
tableNameInputText.pack(side=BOTTOM, anchor=S, fill=BOTH, expand=1, pady=(5,0))



tableNameInputCanvas = Canvas(App,bg='#FFFFFF',bd=0,highlightthickness=0,relief='ridge')
tableNameInputCanvas.grid(row=4,column=0,columnspan=2, rowspan=1)
tableNameInput = Text(tableNameInputCanvas, font='Quicksand 9 bold', bg='#FFF', height=1, width=33, bd=0, relief='ridge', highlightthickness=1, fg='#333', highlightcolor='#069', selectbackground='#1b8a73', selectforeground='#FFFFFF', highlightbackground='#096', spacing1=5, spacing3=5)
tableNameInput.pack(fill=BOTH, padx=10, ipady=0)

runButtonCanvas = Canvas(App,bg='#FFFFFF',bd=0,highlightthickness=0,relief='ridge')
runButtonCanvas.grid(row=3,column=2,columnspan=1, rowspan=2, pady=0)
runButton = Button(runButtonCanvas, text='run', bg='#009966', height=1, width=4, bd=0, font='Quicksand 9 bold', highlightthickness=0, relief='ridge', justify='center', activebackground='#E1E1E1', highlightbackground='#CDCDCD', fg='#FFF', highlightcolor='#009966', command=startProgram)
runButton.pack(fill=BOTH, expand=1, ipady=3, anchor=S, side=BOTTOM, pady=(35,7))



statusCanvas = Canvas(App,bg='#F05555',bd=0,highlightthickness=0,relief='ridge')
statusCanvas.grid(row=5,column=0,columnspan=3, rowspan=1, pady=(10,0), sticky=W, ipadx=7)
statusText = Label(statusCanvas, text='start shifting', font='Quicksand 9 bold', bg='#F05555', fg='#FFFFFF', width=53, anchor=W)
statusText.pack(ipady=5)


App.mainloop()