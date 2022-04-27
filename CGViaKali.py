from pywinauto import Application, Desktop
import time
def Outlook():
    app=Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK')
    time.sleep(30)
    mainDLG=app['Outlook Today - Outlook']
    wizardDLG=mainDLG['Microsoft Office Activation Wizard2'] #there are two handles with the same child window name
    wizardDLG.child_window(title="Cancel", control_type="Button").click() #closes hardware change dialogue
    time.sleep(5)
    mainDLG.sophosTreeItem.click_input(double=True)#Expand the Sophos profile tree
    sophosDLG = app['Sophos - Outlook']
    bthDLG = sophosDLG['BTH2TreeItem']
    bthDLG.click_input(double=True)#expand the BTH2 menu tree
    time.sleep(5)
    bthDLG = app['BTH2 - Sophos - Outlook']
    bthDLG['2TreeItem'].click_input()
    time.sleep(5)
    twoDLG = app['2 - Sophos - Outlook'] #New window name
    email = twoDLG.child_window(title="With Attachments, Subject Your flight has been successfully booked!, Received 1/31/2022, Size 136 KB, Flag Status Unflagged, ", control_type="DataItem")
    email.click_input()
    time.sleep(20)
    twoDLG['DeltaFlightItinerary.docm85 KB1 of 1 attachmentsUse alt + down arrow to open the options menu'].click_input(double=True) #opens the attachment
def Word():
    app = Application(backend="uia").connect(title_re="^DeltaFlightItinerary")
    time.sleep(20)
    mainDLG = app.window(title_re="^DeltaFlightItinerary")
    wizardDLG = mainDLG['Microsoft Office Activation Wizard2'] #there are two handles with the same child window name
    wizardDLG.child_window(title="Cancel", control_type="Button").click() #closes hardware change dialogue
    time.sleep(20)
    mainDLG.child_window(title="Enable Editing", control_type="Button").click_input(double=True)
    time.sleep(20)
    mainDLG.child_window(title="Enable Content", control_type="Button").click_input(double=True)
if __name__ == '__main__':
    time.sleep(30)
    Outlook()
    time.sleep(20)
    Word()