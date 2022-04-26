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
    tmseedDLG=sophosDLG['Threat Menu'].click_input(double=True)
    time.sleep(5)
    seedDLG=app['Threat Menu - Sophos - Outlook']
    seedDLG.child_window(title="With Attachments, Subject Exclusive Suite Upgrade offer from Marriott International, Received 2/4/2022, Size 1 MB, Flag Status Unflagged, ", control_type="DataItem").click_input()
    time.sleep(30)
    seedDLG['FreeSuiteUpgade.docm 1 MB 1 of 1 attachments'].click_input(double=True) #opens the attachment
    time.sleep(10)	
def Word():
    app = Application(backend="uia").connect(title_re="^Free")
    time.sleep(20)
    mainDLG = app.window(title_re="^Free")
    time.sleep(20)
    wizardDLG=mainDLG['Microsoft Office Activation Wizard2'] #there are two handles with the same child window name
    wizardDLG.child_window(title="Cancel", control_type="Button").click() #closes hardware change dialogue
    time.sleep(20)
    mainDLG.child_window(title="Enable Editing", control_type="Button").click_input(double=True)
    time.sleep(20)
    mainDLG.child_window(title="Enable Content", control_type="Button").click_input(double=True)
def TMSeed():
    app = Application().connect(path=r"c:\threat\threatmenu.exe")
    mainDLG = app['C:\threat\threatmenu.exe']
    mainDLG.type_keys('5')
    mainDLG.type_keys('{ENTER}')
    time.sleep(10)
    mainDLG.type_keys('1')
    mainDLG.type_keys('{ENTER}')
    time.sleep(10)
    mainDLG.type_keys('2')
    mainDLG.type_keys('{ENTER}')
    time.sleep(10)
    mainDLG.type_keys('0')
    mainDLG.type_keys('{ENTER}')
    time.sleep(10)
    mainDLG.type_keys('1')
    mainDLG.type_keys('{ENTER}')
    time.sleep(10)
    mainDLG.type_keys('2')
    mainDLG.type_keys('{ENTER}')


if __name__ == '__main__':
    Outlook()
    time.sleep(20)
    Word()
    time.sleep(120)
    TMSeed()