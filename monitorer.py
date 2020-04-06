import win32gui
import win32process
import win32console

import psutil

import openpyxl

# Load the Pandas libraries with alias 'pd' 

import csv
import json
import time
import os
from pynput.keyboard import Key, Listener

import datetime
#win = win32console.GetConsoleWindow() 
#win32gui.ShowWindow(win, 0) 


  

class IdleClassifier(object):

    lastpos = (0,0)
    lastkeytime = time.time()
    idlethreshold = 10 #seconds


    def UpdateKeyTime(self,key):
        self.lastkeytime=time.time()

    def UpdateLastMouse(self):
       _,_,self.lastpos = win32gui.GetCursorInfo()
       



    def Classify(self):

        _, _, (x,y) = win32gui.GetCursorInfo()

        idle = self.lastpos==(x,y) and time.time()-self.lastkeytime> self.idlethreshold

        self.lastpos=(x,y)

        if idle:
            return "IDLE"
        else:
            return "AWAKE"

        
def ClassClassifier(string, classfile):

    activeclass=""

    #for all classes
    for key in classfile:
        #chkec if there is key in window name
        for lol in classfile[key]:


            #if there is exit
            if lol in string:
                activeclass=key
                break

        #if there is exit
        if activeclass is not "":
            break

    return activeclass

 

def GetWindowName():
    return win32gui.GetWindowText(win32gui.GetForegroundWindow())

def WriteExcel(ws,index,arr,isRow=True):
    if isRow:
        for idx, val in enumerate(arr, start=1):
            ws.cell(row=index, column=idx).value = val

def Monitor(interval):


    directory="Outputs/"

    filepath = directory + str(datetime.date.today().strftime("%d-%m-%Y")) + ".xlsx"
    

    fields = ['Subcategory','WindowName','Time','MouseState']




    if not os.path.isfile(filepath):  # False
        wb = openpyxl.Workbook()
        wb.save(filepath)

        print("Created New File")
    else:
        wb = openpyxl.load_workbook(filepath)
        print("loading existing")


    ws1 = wb.active
    ws1.title = "Log"

    WriteExcel(ws1,1,fields)

    wb.save(filepath)

    #row to start writing on
    count=2

    #json with basic categorization information
    with open('info.json', 'r') as f:
        classes = json.load(f)

    print("Started Monitoring Your Work")

    mouseclassifier = IdleClassifier()

    with Listener(on_press=mouseclassifier.UpdateKeyTime) as listener:
        #listener.join()

        #listen for keyboard interrupt
        

            
        #main loop
        while(True):
            try:
                time.sleep(interval)

                string = GetWindowName().lower()

                dt_string = datetime.datetime.now().strftime("%H:%M:%S")

                curclass = ClassClassifier(string, classes)

                output = [ curclass, string,dt_string,mouseclassifier.Classify()]

                WriteExcel(ws1,count,output)

            
                print(output)
            except :
                print("Bonjour")
                wb.save(filepath)
                print("Enregistre")
                break



        #listener.join()
        

def main():
    dirName='Outputs'
    try:
    # Create target Directory
        os.mkdir(dirName)
        print("Directory " , dirName ,  " Created ") 
    except FileExistsError:
        print("Directory " , dirName ,  " already exists")

    
    Monitor(interval=10)

if __name__ == "__main__":
    main()