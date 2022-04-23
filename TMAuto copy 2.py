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
    print(app.windows())
    mainDLG = app['Outlook Today - Outlook']
    #productDLG = mainDLG.child_window(title="Enter your product key", control_type="Window")
    #productDLG.child_window(title="Close", control_type="Button").click_input()
    #time.sleep(5)
    #eulaDLG = mainDLG.child_window(title="Accept the license agreement", control_type="Window")
    #eulaDLG.child_window(title="Accept and start Outlook", control_type="Button").click()
    activateDLG = mainDLG.child_window(title="Microsoft Office Activation Wizard", control_type="Window")
    activateDLG = mainDLG.child_window(title="Cancel", control_type="Button").click()

    time.sleep(5)
    mainDLG= app['Outlook Today - Outlook'] #Main application is presented as 'Outlook Today - Outlook'
    mainDLG.sophosTreeItem.click_input(double=True)#Expand the Sophos profile tree
    time.sleep(5)
    sophosDLG = app['Sophos - Outlook']
    tmDLG=sophosDLG['z.TM AUTO'].click_input(double=True)
    time.sleep(5)
    marriottDLG=app['z.TM AUTO - Sophos - Outlook']
    marriottDLG.child_window(title="With Attachments, Subject 100,000 BONUS Marriott Elite Points Offer, Received 1/31/2022, Size 914 KB, Flag Status Unflagged, ", control_type="DataItem").click_input()
    time.sleep(30)
    marriottDLG['BonusPointsOffer.docm776 KB1 of 1 attachmentsUse alt + down arrow to open the options menu'].click_input(double=True) #opens the attachment
    time.sleep(10)	
	
#TO ENABLE THE MACRO
def marriottMacro():
    app = Application(backend="uia").connect(title_re="^BonusPointsOffer")
    print(app.windows())
    time.sleep(20)
    mainDLG = app.window(title_re="^BonusPointsOffer")
    #productKeyDLG = mainDLG.child_window(title="Enter your product key", control_type="Window")
    #productKeyDLG.CloseButton.click_input()
    time.sleep(20)
    mainDLG.child_window(title="Enable Editing", control_type="Button").click_input(double=True)
    time.sleep(20)
    mainDLG.child_window(title="Enable Content", control_type="Button").click_input(double=True)

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
