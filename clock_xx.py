import win32com.client
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
"""
p = PyMouse()
shell = win32com.client.Dispatch("WScript.Shell")
k = pykeyboard.PyKeyboard()
shell.AppActivate("sfdcclt"); time.sleep(1); p.click(100,100); k.tap_key(k.numpad_keys[4], n=1)

"""

#clocking inn-off automation:

p = PyMouse()
k = PyKeyboard()
shell = win32com.client.Dispatch("WScript.Shell")
def cOff(complete = False):
    #fuction to clock off the job
    shell.AppActivate("sfdcclt") #activate clock program window
    time.sleep(0.1)
    p.click(100,100) #click in area to fully activate
    time.sleep(0.1)
    k.tap_key(k.numpad_keys[4], n=1)
    k.tap_key(k.numpad_keys[2], n=1)
    k.tap_key(k.numpad_keys[0], n=1)
    k.tap_key(k.numpad_keys[4], n=1)
    k.tap_key(k.enter_key, n=1)
    if complete:
        k.tap_key(k.numpad_keys[1], n=1)
    else:
        k.tap_key(k.numpad_keys[1], n=1)
    k.tap_key(k.enter_key, n=1)
    if complete:
        k.tap_key(k.numpad_keys['Add'], n=1)
    #k.tap_key(k.enter_key, n=1)
        
def cOn(wo,op="05"):
    #fuction to clock off the job
    shell.AppActivate("sfdcclt") #activate clock program window
    time.sleep(0.1)
    p.click(100,100) #click in area to fully activate
    time.sleep(0.1)
    k.tap_key(k.numpad_keys[3], n=1)
    k.tap_key(k.numpad_keys[2], n=1)
    k.tap_key(k.numpad_keys[0], n=1)
    k.tap_key(k.numpad_keys[4], n=1)
    k.tap_key(k.enter_key, n=1)
    shell.SendKeys(wo, 0)
    k.tap_key(k.enter_key, n=1)
    shell.SendKeys(op, 0)
    k.tap_key(k.enter_key, n=1)
    k.tap_key(k.enter_key, n=1)
    k.tap_key(k.enter_key, n=1)
    shell.SendKeys("204", 0)
