# -*- coding: cp1250 -*-

import os

#słownik mapujący parametry:

import sys
pathname = os.path.dirname(sys.argv[0])        
PATH = os.path.abspath(pathname)	#sciezka do katalogu gdzie jest skrypt
DRWNUM = "drwnum.py"
print DRWNUM

#os.popen('c:\\windows\\system32\\cmd.exe')
os.system('c:\\windows\\system32\\cmd.exe /K "cd '+PATH+' & "'+DRWNUM)

