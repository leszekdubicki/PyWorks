# -*- coding: cp1250 -*-
from prochecker import *
import win32api
import os
#class PPart(part):
class PPart(draw):
    def exploreLocal(self, prefix = "C:\\Users\\leszekd\\Documents\\tasks", nestLevel = 5):
        #go up by number of levels (skipping some folders) in folders as specified by nestLevel variable:
        l = 0
        wholePath = ''
        P = self['path']
        while l<nestLevel:
            P = os.path.dirname(P)
            if not os.path.basename(P) in ['SolidWorks', '3. Design']:
                wholePath = os.path.basename(P) + '\\' + wholePath
            l+=1
            
        wholePath = prefix  + '\\' + wholePath
        #print wholePath
        win32api.WinExec('explorer '+wholePath.replace('/','\\'))
    def desProp(self):
        swx.RunMacro('C:\Users\leszekd\Desktop\MACRO\Designer Custom Properties.swp','Print_A3_Landscape','main')
    def createLocalFolder(self, prefix = "C:\\Users\\leszekd\\Documents\\tasks"):
        #creates folder on local drive to put temporary  and not-official files there 
        P = self['full_path']
        P2 = P.split(os.sep)
        i = len(P2) - 1
        
        customerFolder = prefix + "\\" + P2[-5][0] + "\\" + P2[-5]
        formatFolder = prefix + "\\" + P2[-5][0] + "\\" + P2[-5] + "\\" + P2[-4]
        
        if not os.path.isdir(customerFolder):
            os.mkdir(customerFolder)
        if not os.path.isdir(formatFolder):
            os.mkdir(formatFolder)

def printList(list):
    #function to print all docs from list using PPC standard printing macro (need to add sheet format recognition)
    for part in list:
        part.open()
        part.activate()
        s = part.round_sheetsize()
        if s in SIZES:
            if SIZES[s][0] == 'A3':
                swx.RunMacro('C:\\Users\\leszekd\\Desktop\\MACRO\\Print A3 - Landscape.swp','Print_A3_Landscape','main')
            elif SIZES[s][0] == 'A4':
                swx.RunMacro('C:\\Users\\leszekd\\Desktop\\MACRO\\Print A4.swp','Print_A41','main')
            else:
                print('Sheet size not recognized')

def addJF():
    #function to create job form and fill some of the fields from the open SW assembly
    #get currently open doc:
    P = PPart()
    
    
