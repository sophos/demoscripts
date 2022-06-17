import psutil
from pywinauto import Application, Desktop, keyboard
import time
import os
import winreg
import sys

log_file = "C:\\Windows\\Temp\\pywin_log.txt"
sys.stdout = open(log_file,'w',encoding="utf8")

def disableHitmanPro(subkey,value):
    reg_key = winreg.OpenKey(
        winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\HitmanPro.Alert",
        0, winreg.KEY_SET_VALUE)
    winreg.SetValueEx(reg_key,subkey,0,winreg.REG_SZ, value)

def toggleHitMan():
    os.system("net stop hmpalertsvc")
    time.sleep(5)
    os.system("net start hmpalertsvc")

def returnPID(process):
    process_name = process
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            processID = proc.pid
            return (processID)

def processRunCheck(processname):
    for proc in psutil.process_iter():
        try:
            if processname.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    return False;

def tamperCheck():
    app = Application(backend="uia").connect(process=sophosUI, visible_only=False)
    dlg = app['Sophos Endpoint Agent']
    checkbox = dlg.child_window(auto_id="OverrideSettingsCheckBox", control_type="CheckBox").wrapper_object()
    value = checkbox.get_toggle_state()
    return value

sophosUI = returnPID("Sophos UI.exe")

def SophosUI():
    app = Application(backend="uia").connect(path="explorer.exe")
    sys_tray = app.window(class_name="Shell_TrayWnd")
    sys_tray.child_window(title='Sophos Endpoint Agent').click_input(double=True)
    app = Application(backend="uia").connect(process=sophosUI, visible_only=False)
    dlg = app['Sophos Endpoint Agent']

    dlg['SettingsRadioButton'].click_input()
    time.sleep(5)
    if (tamperCheck() == 1):
        dlg.child_window(auto_id="OverrideSettingsCheckBox", control_type="CheckBox").click_input()
        dlg.child_window(title="Close", auto_id="NavBarCloseButton", control_type="Button").click_input()
    else:
        dlg.child_window(auto_id="OverrideSettingsCheckBox", control_type="CheckBox").click_input()
        dlg.child_window(auto_id="Files", control_type="Button").click_input()
        dlg.child_window(auto_id="Enable Deep Learning", control_type="Button").click_input()
        dlg.child_window(auto_id="Ransomware Detection", control_type="Button").click_input()
        dlg.child_window(auto_id="Exploit Mitigation", control_type="Button").click_input()
        dlg.child_window(auto_id="Malicious Behavior Detection", control_type="Button").click_input()
        dlg.child_window(auto_id="AMSI Protection", control_type="Button").click_input()
        dlg.child_window(title="Close", auto_id="NavBarCloseButton", control_type="Button").click_input()
def Outlook():
    app = Application(backend="uia").start(r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK')
    time.sleep(60) #Need time for the activate office dialogue to appear. 
    mainDLG = app['Outlook Today - Outlook']
    #wizardDLG=mainDLG['Microsoft Office Activation Wizard2'] #there are two handles with the same child window name
    #wizardDLG.child_window(title="Cancel", control_type="Button").click() #closes hardware change dialogue
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
    mainDLG = app.window(title_re="^Free")
    time.sleep(20)
    #wizardDLG=mainDLG['Microsoft Office Activation Wizard2'] #there are two handles with the same child window name
    #wizardDLG.child_window(title="Cancel", control_type="Button").click() #closes hardware change dialogue
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
    mainDLG = app['Administrator: Windows PowerShell']
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
    SophosUI()
    time.sleep(20)
    disableHitmanPro('HeapHeapHooray','off')
    time.sleep(3)
    toggleHitMan()
    time.sleep(3)
    Outlook()
    time.sleep(20)
    Word()
    time.sleep(60)
    TMSeed()