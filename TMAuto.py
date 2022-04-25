from pywinauto import Application, Desktop
import time

# Outlook function
def Outlook():
    app = Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK')
    time.sleep(30) #Need time for the activate office dialogue to appear. 
    mainDLG = app['Outlook Today - Outlook']
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

# Word function
def Word():
    app = Application(backend="uia").connect(title_re="^BonusPointsOffer")
    time.sleep(20)
    mainDLG = app.window(title_re="^BonusPointsOffer")
    time.sleep(20)
    mainDLG.child_window(title="Enable Editing", control_type="Button").click_input(double=True)
    time.sleep(20)
    mainDLG.child_window(title="Enable Content", control_type="Button").click_input(double=True)

if __name__ == '__main__':
    Outlook()
    time.sleep(20)
    Word()