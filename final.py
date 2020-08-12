# -*- coding: utf-8 -*-
"""
Created on Tue Nov 27 08:45:12 2018

@author: mariat
"""
"""
The code below is creating a Gui window with ability to key in 4 numbers of Client ID
and pull up information from 4 files in one excel for that Client ID

"""


from tkinter import *


def clientid():
    #load libraries
    import pandas as pd
    import numpy as np
    import os
    #assign entry value to ci
    ci=entry_1.get()
    #load files
    ws=r'ws_stores.csv'
    df_ws=pd.read_csv(ws, encoding = "ISO-8859-1")
    datafile_today=r'fd_rc.csv'
    df_today=pd.read_csv(datafile_today,encoding = "ISO-8859-1")    
    app=r'app_file.csv'
    df_app=pd.read_csv(app,encoding="ISO-8859-1")
    nn=pd.read_csv(r'nn.csv', encoding="ISO-8859-1")    

     
    #turn ci to integer as the values in ClientID columns are integers
    ci=int(ci)
    
    #filter files by ci
    df_app=df_app[df_app['ClientID']==ci]    
    df_today=df_today[df_today['ClientID']==ci]
    df_ws=df_ws[df_ws['ClientID']==ci]
    df_north=nn[nn['ClientID']==ci]


    #reset index and replace by a new starting from 0
    df_app.reset_index(drop=True, inplace=True)
    df_today.reset_index(drop=True,  inplace=True)    
    df_ws.reset_index(drop=True,  inplace=True)
    df_north.reset_index(drop=True, inplace=True)
    
    df_north.dropna(axis=1, how='all', inplace=True)
 
    #load libraries for excel file creation
    from pandas import ExcelWriter
    import xlsxwriter
    import xlwt
    import xlwt.Workbook 
    #create excel file with the same name as ci
    writer=pd.ExcelWriter(r"{}.xls".format(ci), engine='xlsxwriter')
    #write each dataframe in a separate excel sheet in one file
    df_app.to_excel(writer,'AppProcessing',index=False)
    df_ws.to_excel(writer,'WS',index=False)
    df_today.to_excel(writer,'FD_RC',index=False)
    df_north.to_excel(writer,'North',index=False)

    #save excel file and start it
    writer.save()
    file=r'{}.xls'.format(ci)
    try:
        os.startfile(file)
    except AttributeError:
        os.system('open %s' % file)
        
    #delete content of Client Id field
    entry_1.delete(0,END)
     
    
def north():
    #oad libraries
    import pandas as pd
    import numpy as np
    import os
    #assign value keyed into the files in GPU to ci
    ci=entry_1.get()
    
    #turn ci to integer as values in ClientID are integers
    ci=int(ci)
    
    #filter file by ci value and replace index with a new starting from 0
    nn=pd.read_csv(r'nn.csv', encoding="ISO-8859-1")    
    df_north=nn[nn['ClientID']==ci]
    df_north.reset_index(drop=True, inplace=True)
    
    #drop all rows that have na values
    df_north.dropna(axis=1, inplace=True)
    #save dataframe to excel
    df_north.to_excel(r'{}.xlsx'.format(ci), index=False)
    file=r'{}.xlsx'.format(ci)
    
    #start file
    try:
        os.startfile(file)
    except AttributeError:
        os.system('open %s' % file)
        
    #delete content of Client Id fiels
    entry_1.delete(0,END)
    


def close():
    #close all excel files open on the desktop
    #possible only on Windows
    
    try:
    
        import win32com.client
        xl = win32com.client.Dispatch("Excel.Application")
        xl.DisplayAlerts = False 
        xl.Quit() 
    except ModuleNotFoundError:
        print("Cannot close all excel files. Please, do it manually.")


 #create window frame with a title  
my_window=Tk()
frmMain = Frame(my_window,bg = 'lightgray',relief=RIDGE,borderwidth=4) 
my_window.title(" Search Widget")

#create buttons and fields

label_1=Label(my_window,text='Client Id ',fg = "white", bg = "purple", borderwidth=4, width=10, height=1,relief=GROOVE,font='Arial 10')

entry_1=Entry(my_window,font='Arial 10')

#start execution by pressing Enter on the keyboard
entry_1.bind('<Return>',(lambda event: clientid()))

label_2=Label(my_window, text="North ID", fg = "white", bg = "purple", borderwidth=4, width=10, height=1, relief=GROOVE,font='Arial 10')
label_2_1=Label(my_window, text='same as client id', fg='purple', borderwidth=4, width=12, height=1,relief=GROOVE,font='Arial 10')

button_1=Button(my_window,text='Submit', command=clientid, fg='purple', borderwidth=4, width=10, height=1,font='Arial 10')


button_2=Button(my_window,text=" North Submit", command=north,fg='purple', borderwidth=4,width=10, height=1,font='Arial 10')

button_5=Button(my_window, text='Close All Files',command=close,fg='purple', borderwidth=4,width=10, height=1,font='Arial 10' )

#buttons, labels and fields situation in the window
label_1.grid(row=0,column=0,columnspan = 1, rowspan=1)
entry_1.grid(row=0,column=1,columnspan = 1, rowspan=1)
button_1.grid(row=0,column=2,columnspan = 1, rowspan=1)
label_2.grid(row=1,column=0,columnspan = 1, rowspan=1)
button_2.grid(row=1, column=2,columnspan = 1, rowspan=1)
label_2_1.grid(row=1, column=1,columnspan = 1, rowspan=1)
button_5.grid(row=3,column=2,columnspan=1,rowspan=1)

    


my_window.mainloop()


