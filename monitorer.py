import  win32gui

from  openpyxl import load_workbook , Workbook

from pandas import DataFrame, to_datetime
# Load the Pandas libraries with alias 'pd' 

import json
import time
import os
import sys
from pynput.keyboard import Listener

import datetime

from win10toast import ToastNotifier

from infi.systray import SysTrayIcon

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

def FindFirstEmptyRow(ws,limit=10000):
    for idx in range(1,limit):

        if ws.cell(row=idx, column=2).value is None:
            return idx

    return -1


def DisplayDailyStats(ws1,interval):

    #Calculate them here:
    data = ws1.values
    # Get the first line in file as a header line
    columns = next(data)[0:]
    # Create a DataFrame based on the second and subsequent lines of data
    df = DataFrame(data, columns=columns)

    #make sure time is in datetime
    df['Time'] = to_datetime(df['Time'],dayfirst=True)  


    #set time range for the current day
    start_date= datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = start_date+datetime.timedelta(days=1) 

    #set time range mask
    mask = (df['Time'] >start_date ) & (df['Time'] <= end_date )

    #get relevant events
    df1 = df.loc[mask]

    #count ocurrences for each categorey
    stats = df1.groupby(['Category'])['Time'].count()


    strvals=""

    #for each category
    for idx,val in stats.items():
        
        #obtain real val
        realval=val*interval
        
        #obtain in hours minutes format
        timestring=""
        hours = datetime.datetime.strptime(str(datetime.timedelta(seconds=realval)),'%H:%M:%S').strftime("%H")
        minutes = datetime.datetime.strptime(str(datetime.timedelta(seconds=realval)),'%H:%M:%S').strftime("%M")

        #take care of strings
        if(hours[0]=='0'):
            hours=hours[1:]
            
        if(minutes[0]=='0'):
            minutes=minutes[1:]
            
        if(hours=="0"):
            timestring = minutes + "m"
        else:
            timestring = hours + "h"+minutes+"m"
        
        strvals = strvals + f"{idx}: {timestring}\n"

    print(strvals)
    #ToastNotifier().show_toast("Daily Stats",strvals,icon_path="eye.ico",duration=10,threaded=True)


def ExitProgram(systry):
    systry.shutdown()
    sys.exit()

def Monitor(interval):

    with open('info.json', 'r') as f:
        classes = json.load(f)

    directory="Outputs/"

    filepath = directory + str(datetime.date.today().strftime("%m-%Y")) + ".xlsx"
    

    fields = ['Category','WindowName','Time','MouseState']




    if not os.path.isfile(filepath):  # False
        wb = Workbook()
        wb.save(filepath)

        print("Created New File")
    else:
        wb = load_workbook(filepath)
        #notificationthread.raise_exception()
        print("loading existing")





    ws1 = wb.active
    ws1.title = "Log"
    
    
    print(ws1)

    #daily stats
    #df = pd.DataFrame(ws.values)


    count = FindFirstEmptyRow(ws1)
    
    WriteExcel(ws1,1,fields)


    wb.save(filepath)


    mouseclassifier = IdleClassifier()

    with Listener(on_press=mouseclassifier.UpdateKeyTime) as listener:
    #listener.join()

        
        while(True):

            time.sleep(interval)

            string = GetWindowName().lower()

            dt_string = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            curclass = ClassClassifier(string, classes)

            output = [ curclass, string,dt_string,mouseclassifier.Classify()]

            WriteExcel(ws1,count,output)
            
            print(count , output) 
            count = count + 1 

            wb.save(filepath)


    #listener.join()

            

def main():
    dirName='Outputs'
    try:
    # Create target Directory
        os.mkdir(dirName)
        print("Directory " , dirName ,  " Created ") 
    except FileExistsError:
        print("Directory " , dirName ,  " already exists")

    #toaster = ToastNotifier()
    #toaster.show_toast("Work Monitoring",    "Monitor Process has been started",    icon_path="eye.ico",    duration=5,    threaded=True)


    # Wait for threaded notification to finish
    #while toaster.notification_active():
    #    time.sleep(0.1)
    
    Monitor(interval=10)

if __name__ == "__main__":
    main()