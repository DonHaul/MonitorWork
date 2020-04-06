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

#win = win32console.GetConsoleWindow() 
#win32gui.ShowWindow(win, 0) 

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



print("Started Monitoring Your Work")



import time

while True:
    time.sleep(1)
    print(    time.time())