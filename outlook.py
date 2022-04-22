from pywinauto import Application, Desktop
import time
import os 
import psutil
import io
import sys
import re

def killProcess(process):
    ti = 0
    name = process
    print('The process I am seeking to kill is: ' + str(name))
    for proc in psutil.process_iter():
    #check whether the process name matches
        if proc.name() == str(name):
            proc.kill()
            print('I have found, and killed ' + str(name))
        ti += 1
        if ti == 0:
            print('There are no running instances of ' + str(name))

def processRunCheck(processname):
    for proc in psutil.process_iter():
        try:
            if processname.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    return False;

def marriottEmailAttachment():
    
    app = Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK')
    time.sleep(30) #Need time for the activate office dialogue to appear. 
    mainDLG = app['Outlook Today - Outlook']
    productDLG = mainDLG.child_window(title="Enter your product key", control_type="Window")
    productDLG.child_window(title="Close", control_type="Button").click_input()
    time.sleep(5)
    eulaDLG = mainDLG.child_window(title="Accept the license agreement", control_type="Window")
    eulaDLG.child_window(title="Accept and start Outlook", control_type="Button").click()
    time.sleep(10)
    killProcess('OUTLOOK.EXE')

if __name__ == '__main__':
   
        killProcess('WINWORD.EXE')

        if (processRunCheck('outlook.exe') == False):
          
            time.sleep(10)
            marriottEmailAttachment()
            time.sleep(30)
            print('Outlook is currently running.')
            marriottMacro()
        
        else:
            print('I am now going to kill the Outlook Process.')
            killProcess('OUTLOOK.EXE')
            time.sleep(5)
            print('Outlook will now start.')
            marriottEmailAttachment()
            time.sleep(5)
            marriottMacro()
