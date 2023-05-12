from pypresence import Presence
import time
import ctypes
import logging

dev_level = logging.INFO
onLaunch = True
EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
titles = []
storedTitle = ""

logging.basicConfig(level=dev_level)


def foreach_window(hwnd, lParam):
    if IsWindowVisible(hwnd):
        length = GetWindowTextLength(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        GetWindowText(hwnd, buff, length + 1)
        titles.append(buff.value)
    return True


def outputDebug():
    logging.info("StoredTitle = {} // rpcActive = {}".format(storedTitle, rpcActive))


def checkForUpdate():
    global rpcActive, phrase
    global storedTitle
    global titles
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    aa = [name for name in titles if "FlowStone 64Bit" in name]  # checks to see if any process with the window name
    # flowstone exists
    searchstring = ''.join(aa)
    titles = []
    for_check = searchstring.split("- FlowStone 64Bit")[0]
    for_check = for_check.split("*")[0]
    for_check = for_check.strip()
    if for_check == "FlowStone 64Bit":
        details = "Looking at Main Screen"
        phrase = "Doing nothing..."
    else:
        details = "Schematic: {}".format(for_check)
        phrase = "Editing schematic"
    if for_check == "" and rpcActive:
        RPC.clear()  # kills the RPC when there is no FlowStone found to be running and it is stated as currently active
        rpcActive = False  # inform everything else that the RPC is closed
        storedTitle = ""

    if rpcActive and for_check != storedTitle:
        storedTitle = for_check
        RPC.update(large_image="largeimage", state=phrase, details=details, start=time.time())
    elif for_check != "" and not rpcActive:
        RPC.update(large_image="largeimage", state=phrase, details=details, start=time.time())
        storedTitle = for_check
        rpcActive = True
    outputDebug()


logging.info("If you get an error stating that the RPC handshake failed, Discord is probably not open")
while True:
    if onLaunch:
        RPC = Presence("1105502852836761663")  # discord application ID
        try:
            RPC.connect()
        except Exception as e:  # TODO: fix generic exception
            onLaunch = True
            logging.warning("RPC handshake failed... trying again in 15 seconds")
        else:
            onLaunch = False
        rpcActive = False
    checkForUpdate()
    time.sleep(15)  # blocking statement is ok in this case
