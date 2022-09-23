from pywinauto import Application, Desktop, keyboard, mouse
import time
import logging
import sys
import os
import psutil
import wmi


#file locations for executables and log files
outlook_path = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
check_file = "C:\\Windows\\Temp\\test.txt"


#These are the dialogues during the initial setup of Microsoft Outlook
productKey = "Enter your product key"
respectPrivacy = "Microsoft respects your privacy"
betterTogether = "Getting better together"
poweringExperience = "Powering your experiences"

#These are the names of each email that kicks off a different threat graph
tmAuto = "With Attachments, Subject 100,000 BONUS Marriott Elite Points Offer, Received 1/31/2022, Size 914 KB, Flag Status Unflagged, "

#Code to establish logging
logger = logging.getLogger(__name__)
stream_handler = logging.StreamHandler(sys.stdout)
file_handler = logging.FileHandler("tmAuto.log")
logger.addHandler(file_handler)
logger.addHandler(stream_handler)
formatter = logging.Formatter("[%(asctime)s] %(levelname)s:%(name)s:%(message)s")
stream_handler.setFormatter(formatter)
file_handler.setFormatter(formatter)
logger.setLevel(logging.DEBUG)

def killProcess(process):
    ti = 0
    name = process
    logger.debug('The process I am seeking to kill is: ' + str(name))
    f = wmi.WMI()

    for process in f.Win32_Process():
        if process.name == name:
            process.Terminate()
            logger.debug('I have found, and killed ' + str(name))
            ti +=1
    if ti == 0:
        logger.debug('There are no running instances of ' + str(name))


def processRunCheck(processname):
    for proc in psutil.process_iter():
        try:
            if processname.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    return False;

#This function searchs a text file for a string in the dialogue object
def currentWindow(app):
    string = str(app.windows()[0])
    remove_first = string[25:]
    window = remove_first[:-9]
    return window

def check_id(filepath,string,dialogue_object):
    temp = sys.stdout
    sys.stdout = open(filepath,'w',encoding="utf8")
    print(dialogue_object.print_control_identifiers())
    sys.stdout = temp
    with open(filepath, 'r') as file:
        content = file.read()
        if string in content:
            logger.debug(f'Dialogue: {string} in {dialogue_object} located.')
            return True
        else:
            logger.debug(f'Dialogue: {string} in {dialogue_object} not found.')
            return False

def revert_tree(app,check_file):

    dlg = app[currentWindow(app)]
    if check_id(check_file,"Junk Email",dlg) == True:
        dlg = app[currentWindow(app)]
        dlg.child_window(title="Sophos", control_type="TreeItem").collapse()
        logger.debug("Tree is reset")
    else:
        logger.debug("Tree is already collapsed")
        pass

def tmAutoOutlook():
    logger.debug("Starting Outlook Function")
    time.sleep(3)
    logger.debug("Outlook is starting...")
    app=Application(backend="uia").start(outlook_path)
    time.sleep(10)
    mainDLG=app['Outlook Today - Outlook']
    logger.debug("Pywinauto is connected to Outlook.")
    logger.debug("Clicking through product key dialogue.")
    mainDLG.child_window(title="Enter your product key", control_type="Window").child_window(title="Close", control_type="Button").click_input()
    time.sleep(3)
    if (check_id(check_file,respectPrivacy,mainDLG) == True):
        logger.debug("Privacy dialogue found, clicking through the rest of initial setup prompts.")
        mainDLG['Microsoft respects your privacyDialog2'].child_window(title="Next", control_type="Button").click_input()
        time.sleep(2)
        mainDLG['Getting better together'].child_window(title="Don't send optional data", control_type="Button").click_input()
        time.sleep(3)
        mainDLG['Powering your experiences2'].child_window(title="Done", control_type="Button").click_input()
    revert_tree(app,check_file)
    logger.debug("Reverted Email Inbox Tree back to initial state..")
    mainDLG.child_window(title="Sophos", control_type="TreeItem").click_input(double=True)
    mainDLG.child_window(title="Inbox", control_type="TreeItem").click_input()
    mainDLG = app['Inbox - Sophos - Outlook (Unlicensed Product)']
    mainDLG[tmAuto].click_input()
    logger.debug("Opening the attachment...")
    mainDLG.child_window(title="Attachment options", control_type="Button").click_input()
    mainDLG.ContextMenu.Open.click_input()
    logger.debug("Attachment opened...")	


def returnPID(process):
    process_name = process
    processID = None
    for proc in psutil.process_iter():
        if process_name in proc.name():
            processID = proc.pid
            return (processID)

def tmAutoWord():
    logger.debug("Word process is starting...")
    
    word = returnPID("WINWORD.EXE")

    if (isinstance(word,int) == False):
        logger.debug("Word PID not detected!")
    
    app = Application(backend="uia").connect(process=word, visible_only=False)

    try:
        try:
            mainDLG = app['Microsoft Word']
            mainDLG.No.click_input()
        except:
            logger.debug("Word is not in safe mode")
        finally:
            mainDLG = app.window(title_re="^DeltaFlightItinerary")
            mainDLG.child_window(title="Enter your product key", control_type="Window").child_window(title="Close", control_type="Button").click_input()
    except:
        logger.debug("Sophos has detected the malicious file.")

if __name__ == '__main__':
    logger.debug("Killing any Microsoft Word Processes...")
    killProcess('WINWORD.EXE')
    if (processRunCheck('outlook.exe') == False):
            logger.debug("Outlook process not found, starting function...")
            logger.debug("Sleeping for 30 seconds...")
            time.sleep(30)
            tmAutoOutlook()
            logger.debug("Sleeping for 20 seconds...")
            time.sleep(20)
            tmAutoWord()
            logger.debug("tmAuto Script Complete.")
        
    else:
            logger.debug("Killing existing Outlook process...")
            killProcess('OUTLOOK.EXE')
            logger.debug("Outlook processed killed.")
            logger.debug("Sleeping for 30 seconds...")
            time.sleep(30)
            tmAutoOutlook()
            logger.debug("Sleeping for 20 seconds...")
            time.sleep(20)
            tmAutoWord()
            logger.debug("tmAuto Script Complete.")
    
