import win32gui
import win32process
import win32console

import psutil

# Load the Pandas libraries with alias 'pd' 

import csv
import json
import time
from datetime import datetime
import os
from pynput.keyboard import Key, Listener

win = win32console.GetConsoleWindow() 
win32gui.ShowWindow(win, 0) 

interval=10
filename = "outputs.csv"
  

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


with open('info.json', 'r') as f:
    classes = json.load(f)

print("Started Monitoring Your Work")

mouseclassifier = IdleClassifier()

with Listener(on_press=mouseclassifier.UpdateKeyTime) as listener:
    #listener.join()

    create = not os.path.isfile(filename)    # False


    with open(filename, "a", newline='') as csvfile:
        outputcsv = csv.writer(csvfile, delimiter=';')

        if create:            
            outputcsv.writerow(["Class","String","Date","Mouse"])
        
        
        while(True):

            time.sleep(interval)

            string = GetWindowName().lower()

            dt_string = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            curclass = ClassClassifier(string, classes)

            output = [ curclass, string,dt_string,mouseclassifier.Classify()]

            outputcsv.writerow(output)
            print(output)

    listener.join()

'''
GET PROCESS EXERCUTABLE PATH
p = psutil.Process(7055)
>>> p.name()
'python'
>>> p.exe()
_, found_pid = win32process.GetWindowThreadProcessId(hwnd)
if found_pid == pid:
hwnds.append(hwnd)
'''

'''
monitor web https://medium.com/@manivannan_data/get-browser-history-chrome-firefox-using-python-in-ubuntu-16-04-fb1c1f7ab546
https://geekswipe.net/technology/computing/analyze-chromes-browsing-history-with-python/


create pics
'''