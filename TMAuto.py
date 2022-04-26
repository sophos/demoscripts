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
    time.sleep(5)
    sophosDLG = app['Sophos - Outlook']
    tmDLG=sophosDLG['z.TM AUTO'].click_input(double=True)
    time.sleep(5)
    marriottDLG=app['z.TM AUTO - Sophos - Outlook']
    marriottDLG.child_window(title="With Attachments, Subject 100,000 BONUS Marriott Elite Points Offer, Received 1/31/2022, Size 914 KB, Flag Status Unflagged, ", control_type="DataItem").click_input()
    time.sleep(30)
    marriottDLG['BonusPointsOffer.docm776 KB1 of 1 attachmentsUse alt + down arrow to open the options menu'].click_input(double=True) #opens the attachment
    time.sleep(10)
def Word():
    app = Application(backend="uia").connect(title_re="^BonusPointsOffer")
    time.sleep(20)
    mainDLG = app.window(title_re="^BonusPointsOffer")
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