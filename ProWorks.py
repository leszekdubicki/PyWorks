# -*- coding: cp1250 -*-
from prochecker import *
import win32api
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