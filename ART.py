from pywinauto import Application, Desktop, keyboard
import time
def Outlook():
    app = Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK')
    time.sleep(30) #Need time for the activate office dialogue to appear. 
    mainDLG = app['Outlook Today - Outlook']
    wizardDLG=mainDLG['Microsoft Office Activation Wizard2'] #there are two handles with the same child window name
    wizardDLG.child_window(title="Cancel", control_type="Button").click() #closes hardware change dialogue
    time.sleep(5)
    mainDLG= app['Outlook Today - Outlook'] #Main application is presented as 'Outlook Today - Outlook'
    mainDLG.sophosTreeItem.click_input(double=True)#Expand the Sophos profile tree
    time.sleep(5)
    sophosDLG = app['Sophos - Outlook']
    artDLG=sophosDLG['Atomic Red Team'].click_input(double=True)
    time.sleep(5)
    bankDLG=app['Atomic Red Team - Sophos - Outlook']
    bankDLG.child_window(title="With Attachments, Subject Suspicious account activity detected, URGENT action needed., Received 1/31/2022, Size 288 KB, Flag Status Unflagged, ", control_type="DataItem").click_input()
    time.sleep(20)
    bankDLG['FraudNotification.docm156 KB1 of 1 attachmentsUse alt + down arrow to open the options menu'].click_input(double=True) #opens the attachment
    time.sleep(10)	
def Word():
    app = Application(backend="uia").connect(title_re="^FraudNotification")
    time.sleep(20)
    mainDLG = app.window(title_re="^FraudNotification")
    wizardDLG = mainDLG['Microsoft Office Activation Wizard2'] #there are two handles with the same child window name
    wizardDLG.child_window(title="Cancel", control_type="Button").click() #closes hardware change dialogue
    time.sleep(20)
    mainDLG.child_window(title="Enable Editing", control_type="Button").click_input(double=True)
    time.sleep(20)
    mainDLG.child_window(title="Enable Content", control_type="Button").click_input(double=True)
def ART():
    app = Application(backend="uia").connect(title_re="Administrator: Windows PowerShell")
    mainDLG = app['Administrator: Windows PowerShell']
    mainDLG.type_keys('Y')
    mainDLG.type_keys('{ENTER}')
if __name__ == '__main__':
    Outlook()
    time.sleep(20)
    Word()
    time.sleep(20)
    ART()