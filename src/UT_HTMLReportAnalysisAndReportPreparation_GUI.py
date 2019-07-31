from tkinter import Label, Button, Entry
import tkinter as tk
from tkinter import messagebox, filedialog, StringVar, SUNKEN, W, X, BOTTOM
from os import path
import threading, sys
import os
import UT_HTMLReportAnalysisAndReportPreparation

class GUI_COntroller:
    '''
       This class initialize the required controls for TkInter GUI
    '''
    def __init__(self,TkObject):

        #Load company image
        Imageloc=tk.PhotoImage(file='../Images/alstom_logo.gif')        
        label3=Label(image=Imageloc,)
        label3.image = Imageloc        
        label3.place(x=200,y=100)

        global TkObject_ref, entryText_UTResPath, AnalyseDirRunBatchButton
        
        #1. select LDRA Tool suite directory
        TkObject_ref = TkObject
        LDRAToolsuitePath=Button(TkObject_ref,activebackground='green',borderwidth=3, anchor="w", text='Select UT test results path:',width=30, command=lambda:GUI_COntroller.selectResDirectory("SourceFilePath"), cursor="hand2")
        LDRAToolsuitePath.place(x=30,y=200)
        LDRAToolsuitePath.config(font=('helvetica',10,'bold'))    

        #1. This is text box where LDRA tool suite directory will be shown to user
        entryText_UTResPath = tk.StringVar()        
        Entry_LDRAToolSuitePath= Entry(TkObject_ref, width=78, textvariable=entryText_UTResPath, bd=1)
        Entry_LDRAToolSuitePath.place(x=290,y=205)                    
        Entry_LDRAToolSuitePath.config(font=('helvetica',10), state="readonly")    
        
        #Exit Window        
        closeButton=Button(TkObject_ref,activebackground='green',borderwidth=4, text='Close Window', command=GUI_COntroller.exitWindow)
        closeButton.place(x=570,y=300)    
        closeButton.config(font=('helvetica',11,'bold'))    

                
        #select sequence files directory
        AnalyseDirRunBatchButton=Button(TkObject_ref,activebackground='green',borderwidth=4, text='Generate UT Report',width=25, command=GUI_COntroller.RunTest)
        AnalyseDirRunBatchButton.place(x=200,y=300)
        AnalyseDirRunBatchButton.config(font=('helvetica',11,'bold'))            

    def selectResDirectory(dirSelectionType): 
        global entryText_UTResPath
        currdir = os.getcwd()
        
        if dirSelectionType == "SourceFilePath":
            selectedDir_res = filedialog.askdirectory(initialdir=currdir, title='Please select UT Results directory')
            
            if len(selectedDir_res)> 0:
                if not path.isdir(selectedDir_res):
                    entryText_UTResPath.set("")            
                    messagebox.showerror('Error','Please select a valid directory!')
                else:
                    entryText_UTResPath.set(str(selectedDir_res))
        
    def exitWindow():
            TkObject_ref.destroy()
            

    def RunTest():

        if len(entryText_UTResPath.get()) > 0:
            ProjectDirAnalysis.RunAnalysis()
        else:
            messagebox.showerror('Error','Please select LDRA tool path!') 
               
            
        
class ProjectDirAnalysis:
    def RunAnalysis(): 
    
        global statusBarText
        AnalyseDirRunBatchButton.config(state="disabled")

        statusBarText = StringVar()        
        StatusLabel = Label(TkObject_ref, textvariable=statusBarText, fg="green", bd=1,relief=SUNKEN,anchor=W) 
        StatusLabel.config(font=('helvetica',11,'bold'))
        StatusLabel.pack(side=BOTTOM, fill=X)
        
        thread = threading.Thread(target=UT_HTMLReportAnalysisAndReportPreparation.script_exe, args = (entryText_UTResPath.get(), TkObject_ref, statusBarText))
        
        thread.start()    

if __name__ == '__main__':    
    
    root = tk.Tk()
       
    #Change the background window color
    root.configure(background='#b7bbc7')     
       
    #Set window parameters
    root.geometry('850x680')
    root.title('UT Results analysis')
       
    #Removes the maximizing option
    root.resizable(0,0)
       
    ObjController = GUI_COntroller(root)
       
    #keep the main window is running
    root.mainloop()
    sys.exit()
