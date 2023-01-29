# Program created for DN company used to create batches for terminal update in BP env.

from tkinter import *
from pandas import DataFrame, ExcelWriter, read_excel
from tkinter import filedialog
import datetime
import calendar
from tkcalendar import DateEntry

window = Tk()
window.title("Update Generator!")
window.geometry("295x150")
window.resizable(False, False)

teminalList = []
terminalIdLength = 0
HOURSOFUPDATE1 = ['00:30', '2:00', '3:30', '5:00', '6:30']
HOURSOFUPDATE2 = ['20:00', '21:30', '00:30', '2:00', '3:30', '5:00', '6:30']
HOURSOFUPDATE3 = ['00:30', '1:30', '2:30', '3:30', '4:30', '5:30', '6:30']
HOURSOFUPDATE4 = ['20:00', '21:00', '22:00', '00:30', '1:30', '2:30', '3:30', '4:30', '5:30', '6:30']
COLOURS = [ "red", "blue", "green", "yellow", "purple", "orange", "white", "black", "pink", "grey" ]

def openFile():
    filepath = filedialog.askopenfilename(initialdir='c:\\', title="find terminal connected report")
    try:
        df = read_excel(filepath)
        terminalIdLength = len(df["terminal_id"])
        labelBatchResults.config(text=terminalIdLength)
        for terminalid in df["terminal_id"]:
            teminalList.append(terminalid)
        button.config(text='Generate', state=ACTIVE)
    except Exception as e:
        labelBatchResults.config(text="Cannot recognize a file")
    
def doMagic():
    
    amountOfBatches = 0
    amoutOfRows = 0
    rows = 0
    
    if(var1.get() == 1):
        amountOfBatches = int(len(teminalList)/250+1)
        print(amountOfBatches)
        amoutOfRows = 250
        rows  = list(range(1,amoutOfRows+1))
    else:
        amountOfBatches = int(entry.get())
        amoutOfRows = int(len(teminalList)/amountOfBatches)+1
        rows = list(range(1,amoutOfRows+1))
    
    
    columns = []

    
    #creating of columns
    
    def appendColumns(hoursofupdate, positions, mark):
        k = 0
        startOfDeployment = cal.get_date()
        for i in range(amountOfBatches):
            if(startOfDeployment.weekday() == calendar.THURSDAY and k==positions):
                if(mark == "ND"):
                    k=0
                    startOfDeployment += datetime.timedelta(days=3)  
                    columns.append(str(startOfDeployment) +" " +hoursofupdate[k])
                    k+=1
                else:
                    k=0
                    startOfDeployment += datetime.timedelta(days=4)  
                    columns.append(str(startOfDeployment) +" " +hoursofupdate[k])
                    k+=1
            elif(k == positions):
                if(mark == "ND" or "PN"):
                    k=0
                    columns.append(str(startOfDeployment) +" " +hoursofupdate[k])
                    k+=1
                else:
                    k=0
                    startOfDeployment += datetime.timedelta(days=1)
                    columns.append(str(startOfDeployment) +" " +hoursofupdate[k])
                    k+=1
            elif(hoursofupdate[k] == '22:00' or hoursofupdate[k] == '21:30'):
                columns.append(str(startOfDeployment) +" " +hoursofupdate[k])
                startOfDeployment += datetime.timedelta(days=1)
                k+=1                
            else:
                columns.append(str(startOfDeployment) +" " +hoursofupdate[k])
                k+=1       
                 
    if(variable.get() == "00:30 - 6:30 (1,5h)"):
        appendColumns(HOURSOFUPDATE1, len(HOURSOFUPDATE1), "norm")     
    elif(variable.get() == "20:00 - 6:30 (1,5h) - PN"):
        appendColumns(HOURSOFUPDATE2, len(HOURSOFUPDATE2), "PN")
    elif(variable.get() == "00:30 - 6:30 (1h)"):
        appendColumns(HOURSOFUPDATE3, len(HOURSOFUPDATE3), "norm")
    elif(variable.get() == "20:00 - 6:30 (1h) - PN"):
        appendColumns(HOURSOFUPDATE4, len(HOURSOFUPDATE4), "PN")  
    elif(variable.get() == "20:00 - 6:30 (1h) - ND"):
        appendColumns(HOURSOFUPDATE4, len(HOURSOFUPDATE4), "ND") 
    elif(variable.get() == "20:00 - 6:30 (1,5h) - ND"):
        appendColumns(HOURSOFUPDATE2, len(HOURSOFUPDATE2), "ND") 
        
    #creating lists and adding it to worksheet
    listSplited = [teminalList[x:x+amountOfBatches] for x in range(0, len(teminalList), amountOfBatches)]
    dk = DataFrame(listSplited, columns=columns, index=rows)       
       
    file = filedialog.asksaveasfilename(defaultextension=".xlsx")
    dk.to_excel(file, sheet_name='Deployment')

    #File Operations:
    writer = ExcelWriter(file) 
    dk.to_excel(writer, sheet_name='Deployment')
    
    #autoadjustcolumns
    for column in dk:
        column_length = max(dk[column].astype(str).map(len).max(), len(column))
        col_idx = dk.columns.get_loc(column)+1
        writer.sheets['Deployment'].set_column(col_idx, col_idx, column_length)
            
    writer.save()    

#GUI
menubar = Menu(window)
window.config(menu=menubar)

fileMenu = Menu(menubar, tearoff=0)
menubar.add_cascade(label='File', menu=fileMenu)
fileMenu.add_command(label='open', command=openFile, compound='left')

labelterminalhNo = Label(text="Terminal numbers:")
labelterminalhNo.grid(row=1, column=1)

labelBatchResults = Label()
labelBatchResults.grid(row=1, column=2)

LabelHowMany = Label(text='How many batches:')
LabelHowMany.grid(row=2, column=1)

entry = Entry()
entry.grid(row=2, column=2)

labelColour = Label(text="Max 250 terminals:")
labelColour.grid(row=3, column=1)

var1 = IntVar()
checkboxColour = Checkbutton(window, variable=var1, onvalue=1, offvalue=0)
checkboxColour.grid(row=3, column=2)

labelTimeFram = Label(text="Timeframe:")
labelTimeFram.grid(row=4, column=1)

variable = StringVar(window)
variable.set("00:30 - 6:30")
timeFrameMenu = OptionMenu(window, variable, "00:30 - 6:30 (1,5h)", "00:30 - 6:30 (1h)","20:00 - 6:30 (1,5h) - PN", "20:00 - 6:30 (1h) - PN", "20:00 - 6:30 (1h) - ND", "20:00 - 6:30 (1,5h) - ND")
timeFrameMenu.grid(row=4, column=2)

labelDeploymentStart = Label(text='First day of deplyment:')
labelDeploymentStart.grid(row=5, column=1)

cal = DateEntry(window, width=12, year=2022, month=6, day=1, background='darkblue', foreground='white', borderwidth=2)
cal.grid(row=5, column=2)

button = Button(text="Import Report", state=DISABLED, command=doMagic)
button.grid(row=6, column=1)

window.mainloop()