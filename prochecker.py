# -*- coding: cp1250 -*-
#Version - 1.18 (2012-08-08)
#Dodana funkcja ustawiania jednostek (set_units) na mm
#Dodana funkcja "round_sheetsize" do drukowania nie-A formatow arkusza w printerze
#Dodana funkcja Rebuild
#Zmieniona funkcja "check_mand_infos"
from win32com.client import gencache, constants, pythoncom

#sw = win32com.client.Dispatch("SldWorks.Application")
#sw = sldworks.ISldWorks(DispatchEx('SldWorks.Application'))
sldmod = gencache.EnsureModule('{83A33D31-27C5-11CE-BFD4-00400513BB57}', 0, 22, 0) #SW2014

#sldmod = gencache.EnsureModule('{83A33D31-27C5-11CE-BFD4-00400513BB57}', 0, 19, 0)

import win32com.client
import time

import win32api
USER = win32api.GetUserName()

import ConfigParser
import os.path
import win32clipboard

#sprawdzenie, czy solid chodzi:
import win32pdh
#try:
#    import solidnums
#except:
#    import imp
#    cnc = imp.load_source("solidnums","../cnc/solidnums.py")

def get_processes():
    win32pdh.EnumObjects(None, None, win32pdh.PERF_DETAIL_WIZARD)
    junk, instances = win32pdh.EnumObjectItems(None,None,'Process', win32pdh.PERF_DETAIL_WIZARD)

    proc_dict = {}
    for instance in instances:
        if proc_dict.has_key(instance):
            proc_dict[instance] = proc_dict[instance] + 1
        else:
            proc_dict[instance]=0

    proc_ids = []
    for instance, max_instances in proc_dict.items():
        for inum in xrange(max_instances+1):
            hq = win32pdh.OpenQuery() # initializes the query handle 
            try:
                path = win32pdh.MakeCounterPath( (None, 'Process', instance, None, inum, 'ID Process') )
                counter_handle=win32pdh.AddCounter(hq, path) #convert counter path to counter handle
                try:
                    win32pdh.CollectQueryData(hq) #collects data for the counter 
                    type, val = win32pdh.GetFormattedCounterValue(counter_handle, win32pdh.PDH_FMT_LONG)
                    proc_ids.append((instance, val))
                except win32pdh.error, e:
                    #print e
                    pass

                win32pdh.RemoveCounter(counter_handle)

            except win32pdh.error, e:
                #print e
                pass
            win32pdh.CloseQuery(hq) 

    return proc_ids

try:
	PROCS = dict(get_processes())
	while not "SLDWORKS" in PROCS:
		print(u"Uruchom solidworksa i naciśnij ENTER lub wpisz \"q\" i naciśnij ENTER aby zakończyc")
		KOM = raw_input()
		if KOM=="q":
			import sys
			sys.exit()
		else:
			PROCS = dict(get_processes())
		
	print "solid dziala"
except:
	print "solid powinien byc uruchomiony."
	
PRINT=0

if PRINT:
	print "pobieram obiekt aplikacji solida"

#swx = sldmod.ISldWorks(win32com.client.DispatchEx('SldWorks.Application'))
swx = win32com.client.Dispatch('SldWorks.Application')
swx = sldmod.ISldWorks(swx)

if PRINT:
	print "pobieram obiekt aplikacji excela"
#excel = win32com.client.Dispatch("Excel.Application")

import os

try:
	import numerator
except:
	print "nie znalazlem modulu numerator, moze to powodowac blady przy niektorych funkcjach!"

import copy
cp=copy.deepcopy

def is_sw_file(nazwa_pliku):
	#zwraca 0, jeśli plik nie jest plikiem solida, 1 jeśli to część, 2 jeśli złorzenie i 3 jeśli rysunek
	ext=nazwa_pliku[nazwa_pliku.rindex("."):]
	if ext.lower()==".sldprt":
		return 1
	elif ext.lower()==".sldasm":
		return 2
	elif ext.lower()==".slddrw":
		return 3
	else:
		return 0
def str3(NAPIS):
	if isinstance(NAPIS,float):
		NAPIS=str(int(NAPIS))
	elif isinstance(NAPIS,int):
		NAPIS=str(NAPIS)
	return NAPIS.encode("cp1250")
def get_files_dirs(katalog):
	#zwraca listę plików i listę katalogów
	if not katalog[len(katalog)-1]=="\\":
		katalog=katalog+"\\"
	PLIKI=[]
	KATALOGI=[]
	LISTA=os.listdir(katalog)
	for i in range(0,len(LISTA)):
		if os.path.isdir(katalog+LISTA[i]):
			KATALOGI.append(LISTA[i])
		else:
			PLIKI.append(LISTA[i])
	return KATALOGI, PLIKI

def find_parts(PLIK=None,CONFIG=""):
	#otwarcie dokumentu:
	if not PLIK==None:
		#ustalenie typu:
		#if is_sw_file(PLIK)==1:
		#	TYP=0x1
		#elif is_sw_file(PLIK)==2:
		#	TYP=0x2
		#DOC=swx.OpenDoc6(PLIK,TYP,0,CONFIG,None, None)
		DOC=PLIK
		#ustawienie konfiguracji:
		cmgr=sldmod.IConfigurationManager(DOC.ConfigurationManager)
		aConf=sldmod.IConfiguration(cmgr.ActiveConfiguration)
		ConfName=aConf.Name
	else:
		DOC = sldmod.IModelDoc2(swx.ActiveDoc)            # get active document
	if PLIK=="":
		PLIK=(DOC.GetPathName()).encode('cp1250')
	#słownik opisujący ten dokument:
	thispart=get_data_from_model(DOC,CONFIG)
	"""
	if CONFIG=="":
		cmgr=sldmod.IConfigurationManager(DOC.ConfigurationManager)
		aConf=sldmod.IConfiguration(cmgr.ActiveConfiguration)
		ConfName=aConf.Name
		thispart['config']=ConfName.encode('cp1250')
	else:
		thispart['config']=CONFIG
	SCIEZKA=(DOC.GetPathName()).encode('cp1250')
	thispart['plik']=SCIEZKA[SCIEZKA.rindex('\\')+1:]
	thispart['katalog']=SCIEZKA[0:SCIEZKA.rindex('\\')+1]
	#teraz dane z modelu:
	thispart['PartNo']=(DOC.GetCustomInfoValue(thispart["config"],"PartNo")).encode('cp1250')
	if thispart['PartNo']=="":
		thispart['PartNo']=(DOC.GetCustomInfoValue("","PartNo")).encode('cp1250')
	thispart['SubPartNo']=(DOC.GetCustomInfoValue(thispart["config"],"SubPartNo")).encode('cp1250')
	if thispart['SubPartNo']=="":
		thispart['SubPartNo']=(DOC.GetCustomInfoValue("","SubPartNo")).encode('cp1250')
	"""
	ADOC = sldmod.IAssemblyDoc(DOC)            # get active document
	ADOC.ResolveAllLightWeightComponents(False)	#przywrócenie do pełnej pamięci komponentów
		
	#teraz przeszukanie wszystkich części i wrzucenie ich do bazy danych
	PARTS=[]
	#pobranie komponentu root:
	if PRINT:
		print "\npobieram listę komponentów złorzenia:"
	cmgr=sldmod.IConfigurationManager(DOC.ConfigurationManager)
	aConf=sldmod.IConfiguration(cmgr.ActiveConfiguration)
	ROOT=sldmod.IComponent2(aConf.GetRootComponent())
	KOMPONENTY=ROOT.GetChildren()
	IKOMPONENTY=[]
	PARTS=[]
	for i in range(0,len(KOMPONENTY)):
		IKOMPONENTY.append(sldmod.IComponent2(KOMPONENTY[i]))
		#print IKOMPONENTY[i].GetSuppression()
		#raw_input()
		if IKOMPONENTY[i].GetSuppression()==0:	#sprawdzenie, czy komponent nie jest wygaszony
			continue
		MOD=sldmod.IModelDoc2(KOMPONENTY[i].GetModelDoc)
		MOD_CONF=(KOMPONENTY[i].ReferencedConfiguration).encode('cp1250')
		newpart=get_data_from_model(MOD,MOD_CONF)
		#dodanie dodatkowych informacji:
		newpart["whereis"]=[thispart['full_path']]
		if thispart.has_key("poz"):
			newpart["poz"]=thispart["poz"]
		"""
		newpart['config']=(KOMPONENTY[i].ReferencedConfiguration).encode('cp1250')
		SCIEZKA=(KOMPONENTY[i].GetPathName).encode('cp1250')
		newpart['plik']=SCIEZKA[SCIEZKA.rindex('\\')+1:]
		newpart['katalog']=SCIEZKA[0:SCIEZKA.rindex('\\')+1]
		#teraz dane z modelu:
		newpart['PartNo']=(MOD.GetCustomInfoValue(newpart["config"],"PartNo")).encode('cp1250')
		if newpart['PartNo']=="":
			newpart['PartNo']=(MOD.GetCustomInfoValue("","PartNo")).encode('cp1250')
		newpart['SubPartNo']=(MOD.GetCustomInfoValue(newpart["config"],"SubPartNo")).encode('cp1250')
		if newpart['SubPartNo']=="":
			newpart['SubPartNo']=(MOD.GetCustomInfoValue("","SubPartNo")).encode('cp1250')
		newpart["whereis"]=[SCIEZKA]
		"""
		
		#print newpart['plik']
		#raw_input()
		PARTS.append(cp(newpart))
		#jeśli to złorzenie to trzebaby znależć wszystkie części w tym złorzeniu:
		if is_sw_file(newpart['plik'])==2:
			PARTS=PARTS+find_parts(MOD,newpart['config'])
	return PARTS

def close_all_sw_docs():
	#zamyka wszystkie dokumenty solidworksa:
	while 1:
		DOC=swx.ActiveDoc
		if DOC==None:
			break
		else:
			DOC=sldmod.IModelDoc2(DOC)
			DANE=get_data_from_model(DOC)
			if PRINT:
				print "zamykam ", DANE['plik']
			swx.CloseDoc(DANE['full_path'])

def get_data_from_model(MOD=None,CONFIG=""):
	#pobiera dane z modelu lub otwartego dokumentu i zwraca je
	#otwarcie dokumentu:
	if not MOD==None:
		DOC=MOD
		#ustawienie konfiguracji:
		if not is_sw_file(MOD.GetPathName())==3:
			cmgr=sldmod.IConfigurationManager(DOC.ConfigurationManager)
			aConf=sldmod.IConfiguration(cmgr.ActiveConfiguration)
			ConfName=aConf.Name
		else:
			ConfName=""
	else:
		DOC = sldmod.IModelDoc2(swx.ActiveDoc)            # get active document
	#słownik opisujący ten dokument:
	thispart={}
	SCIEZKA=DOC.GetPathName()
	SCIEZKA=SCIEZKA.encode('cp1250')
	thispart['plik']=SCIEZKA[SCIEZKA.rindex('\\')+1:]
	thispart['name']=thispart['plik'][0:thispart['plik'].rindex(".")]
	thispart['katalog']=SCIEZKA[0:SCIEZKA.rindex('\\')+1]
	thispart['path']=thispart['katalog']
	thispart['full_path']=SCIEZKA
	#teraz dane z modelu:
	if (CONFIG=="") and (not is_sw_file(thispart['full_path'])==3):
		cmgr=sldmod.IConfigurationManager(DOC.ConfigurationManager)
		aConf=sldmod.IConfiguration(cmgr.ActiveConfiguration)
		ConfName=aConf.Name
		thispart['config']=ConfName
	else:
		thispart['config']=CONFIG
	thispart['PartNo']=(DOC.GetCustomInfoValue(thispart["config"],"PartNo"))
	if thispart['PartNo']=="":
		thispart['PartNo']=(DOC.GetCustomInfoValue("","PartNo"))
	thispart['SubPartNo']=(DOC.GetCustomInfoValue(thispart["config"],"SubPartNo"))
	if thispart['SubPartNo']=="":
		thispart['SubPartNo']=(DOC.GetCustomInfoValue("","SubPartNo"))
	return part(thispart)

def readtime(StringTime):
	#zwraca czas w pstaci sekund (jak time.time()) z ciągu znaków "YYYYMMDDHHMMSS"
	S = StringTime
	T = (int(S[:4]) , int(S[4:6]) , int(S[6:8]) , int(S[8:10]) , int(S[10:12]) , int(S[12:14]) , int(S[14:15]) , int(S[15:18]) , int(S[18:]))
	return time.mktime(T)

def timestamp(t = None):
	Format = "%Y%m%d%H%M%S%w%j0"
	if t == None:
		t = time.time()
	return time.strftime(Format,time.localtime(t))
def show_time_diff(t):
	#pokazuje roznice miedzy obecnym a podanym (w sekundach) czasem 
	Time = time.localtime(time.time()-t)
	Text = ""
	if Time[0]-1970 > 1:
		#rok
		Text = str(Time[0]-1970)+" l. , "
	elif  Time[0]-1970 == 1:
		Text = str(Time[0]-1970)+" r. , "
	if not Time[1]-1 == 0:
		#miesiac
		Text += str(Time[1]-1)+" mies. , "
	if not Time[2]-1 == 0:
		#dni
		Text += str(Time[2]-1)+" d. , "
	if not Time[3]-1 == 0:
		#godz
		Text += str(Time[3]-1)+" godz. i "
	if not Time[4] == 0:
		#min.
		Text += str(Time[4])+" min. i "
	if not Time[5] == 0:
		#sek.
		Text += str(Time[5])+" sek. # "
	return Text[:-2]
	
	
class whatwasdone:
	#klasa zapisuje i odczytuje co zostalo przetworzone jakakolwiek funkcją w danym katalogu, 
	#przy starcie sprawdza, czy w tym katalogu jest juz rekord operacji (trzeba jakis kod operacji podac) i jak nie wszystko zostalo zrobione to mozna wznowic operacje
	#czyli musi byc:
	#start
	#sprawdzenie czy sie wysypalo wczesniej 
	#sprawdzenie kiedy sie wysypalo wczesniej 
	#sprawdzenie wystartowalo wczesniej 
	#odczytanie listy plikow do pominiecia
	#utworzenie i skasowanie pliku
	#
	def __init__(self,Folder,Kod):
		self._kod = Kod
		self._folder = Folder.replace("/","\\")
		if self._folder[-1] == "\\":
			self._folder = self._folder[:-1]
		self._done = self.readfile()
	def filename(self):
		return "__whatwasdone__"+self._kod+".txt"
	def filepath(self):
		return self._folder+"\\"+self.filename()
	def has_file(self):
		return os.path.isfile(self.filepath())
	def readfile(self):
		#plik sklada sie z naglowka i listy plikow kazdy w osobnej linii
		#nagelowek to trzy linie:
		#data rozpoczecia zadania (jeden ciag 19 cyfr w formacie jak w funkcji timestamp(), albo 0 jak nie rozpoczeto)
		#data ostatniej operacji (tak samo)
		#status: 0 - nie rozpoczete, 1 - w trakcie, 2 - zakonczone (wlasciwie to w 2 nie powinna wystepowac bo po zakonczeniu plik bedzie kasowany)
		self._lastdate = 0
		self._startdate = 0
		self._status = 0
		self._donefilelist = []
		self._file = None
		if os.path.isfile(self.filepath()):
			self._file = file(self.filepath(),"r")
			TEXT = self._file.readlines()
			self._file.close()
			#usuniecie ew pustych wierszy:
			i = 0
			while TEXT[i].strip() == "":
				TEXT = TEXT[i+1:]
				i+=1
			#odczyt daty z naglowka:
			self._lastdate = readtime(TEXT[0].strip())
			self._startdate = readtime(TEXT[1].strip())
			self._status = int(TEXT[2].strip())
			for i in range(3,len(TEXT)):
				L = TEXT[i]
				self._donefilelist.append(L.replace("/","\\").strip().lower())
	def __contains__(self, filename):
		filename = filename.replace("/","\\")
		filename = filename.lower()
		return filename in self._donefilelist
	def startdate(self):
		return self._startdate
	def lastdate(self):
		return self._lastdate
	def howold(self):
		return time.time()-self.lastdate()
	def strhowold(self):
		return show_time_diff(self.howold())
	def setstartdate(self, t = None):
		if t == None:
			t = time.time()
		self._file = file(self.filepath(),"w")
		self._file.seek(0)
		self._file.write(timestamp())
		self._file.close()
	def createfile(self):
		if not os.path.isfile(self.filepath()):
			self._file = file(self.filepath(),"w")
			self._file.write(timestamp())
			self._file.write("\r\n")
			self._file.write(timestamp())
			self._file.write("\r\n0\r\n")
			self._file.close()
	def rmfile(self):
		os.remove(self.filepath())
	def append(self, fpath):
		if not os.path.isfile(self.filepath()):
			self.createfile()
		self._file = file(self.filepath(),"r")
		Lines = self._file.readlines()
		self._lastdate = time.time()
		Lines[0] = timestamp(self._lastdate)+"\r\n"
		Lines.append(fpath.lower()+"\r\n")
		self._file = file(self.filepath(),"w")
		self._file.seek(0)
		self._file.writelines(Lines)
		#self._file = file(self.filepath(),"w+")
		#self._file.write(timestamp(self._lastdate))
		self._file.close()
		
		
	def start(self):
		self.createfile()
	def __len__(self):
		return len(self._donefilelist)
		
		

	
class part:
	#klasa przechowuje informacje na temat części maszyny
	#może to byś podzespół lub cząść
	def __init__(self,dane):
		self._dane=dane	#słownik zwracany przez funkcję find_parts
		#pole _class - klasa części (np rolka)
		self._class=None
		#pole _PartNo - Numer części
		self._PartNo=None
		#Pole _SubPartNo - subnumer części (patrz Numerator.py)
		self._SubPartNo=None
		#Pole _drawings - na jakich rysunkach w projekcie znajduje się część
		self._drawings=[]
		self._modeldoc=None
		self._custom_info={}
	def __getitem__(self,item):
		try:
			return self._dane[item]
		except:
			print "brak pozycji ",item," w danych części"
			return None
	def __setitem__(self,item,val):
		self._dane[item]=val
	def has_key(self,key):
		return self._dane.has_key(key)
	def kod_typu(self):
		if is_sw_file(self['plik'])==1:
			return 0x1
		elif is_sw_file(self['plik'])==2:
			return 0x2
		elif is_sw_file(self['plik'])==3:
			return 0x3
	def open_opcje(self):
		if is_sw_file(self['plik'])==1:
			return 0
		elif is_sw_file(self['plik'])==2:
			return 0
		elif is_sw_file(self['plik'])==3:
			return 0x0
	def open(self):
		#otwiera część w solidzie i czyni jej dokument aktywnym
		if self._modeldoc==None:
			MOD=sldmod.IModelDoc2(swx.OpenDoc6(self['full_path'],self.kod_typu(),self.open_opcje(),self['config'],None,None))
			self._modeldoc=MOD
	def activate(self):
		self.open()
		swx.ActivateDoc2(self['plik'],1,None)
		self._modeldoc=sldmod.IModelDoc2(swx.ActiveDoc)
	def getPartDoc(self):
		self._partdoc = sldmod.PartDoc(self._modeldoc)
	def Rebuild(self):
		#self._modeldoc.Rebuild(1+2+8)
		self._modeldoc.ForceRebuild3(False)
		#swRebuildAll 1 or 0x1; Assembly or drawing; rebuilds geometry that has not been regenerated 
		#swForceRebuildAll 2 or 0x2; Assembly or drawing; Forces a rebuild of all geometry 
		#swUpdateMates 4 or 0x4; Assembly only; only rebuilds mates, which is much faster than rebuilding the geometry. Especially useful for IComponent2::Transform2 
		#swCurrentSheetDisp 8 or 0x8; Drawing only; only rebuilds the display of the views on the current drawing sheet 
		#swUpdateDirtyOnly 16 or 0x10; Drawing only; only rebuilds drawing views that are dirty when OR'd with swCurrentSheetDisp option 
	def rebuild(self):
		#added just to have lowercase method for this
		self.Rebuild()
	def close(self):
		swx.CloseDoc(self['katalog']+self['plik'])
		self._modeldoc=None
	def get_ci(self,item):
		#od get_custom_info
		self.open()
		self.activate()
		VAL=self._modeldoc.GetCustomInfoValue(self['config'],item)
		self._custom_info[item]=VAL	#zapisanie, żeby można było używać offline
		return VAL
	def set_ci(self,item,val):
		#od set_custom_info
		self.open()
		self.activate()
		#val=unicode(val.decode())
		self._modeldoc.AddCustomInfo3(self['config'],item,0x1e,val)
		self._modeldoc.SetCustomInfo2(self['config'],item,val)
	def set_global_ci(self,item,val):
		self.open()
		self.activate()
		self._modeldoc.AddCustomInfo3("",item,0x1e,val)
	def save(self,opcje=0x1):
		"""		
		swSaveAsOptions_AvoidRebuildOnSave=0x8        # from enum swSaveAsOptions_e
		swSaveAsOptions_Copy          =0x2        # from enum swSaveAsOptions_e
		swSaveAsOptions_DetachedDrawing=0x80       # from enum swSaveAsOptions_e
		swSaveAsOptions_OverrideSaveEmodel=0x20       # from enum swSaveAsOptions_e
		swSaveAsOptions_SaveEmodelData=0x40       # from enum swSaveAsOptions_e
		swSaveAsOptions_SaveReferenced=0x4        # from enum swSaveAsOptions_e
		swSaveAsOptions_Silent        =0x1        # from enum swSaveAsOptions_e
		swSaveAsOptions_UpdateInactiveViews=0x10       # from enum swSaveAsOptions_e
		"""
		self.open()
		self.activate()
		self._modeldoc.Save3(opcje,None,None)
	def check_ness_infos(self,prompt=1,ness_infos=['description','note1','rev']):
		#sprawdza, czy w pliku rysunku i modeli są zachowane odpowiednie informacje
		#config to obiekt typu ConfigParser
		NESS_INFOS=ness_infos
		##########################################
		#sekcja z konfiguracją - nieaktualna!!!
		#if not config==None:
		#	#import ConfigParser
		#	if not config.has_section(self['name']):
		#		config.add_section(self['name'])
		#		config.set(self['name'],'file',self['full_path'])
		###########################################
			
		STATUS=1	#zwraca 1 jeśli wszystko jest, 0 jak nie
		self.open()
		self.activate()
		self._modeldoc.ShowConfiguration(self['config'])
		if is_sw_file(self._dane['plik'])==3:
			IS_SM_C=0
			HAS_SM_C=0
		else:
			IS_SM_C=self.is_sm_config()
			#HAS_SM_C=self.has_sm_config()
			HAS_SM_C=0
		for I in NESS_INFOS:
			if (len(self.get_ci(I).strip())==0) or (HAS_SM_C) or (I=='rev' and (not self.get_ci(I).strip().isdigit())):
				#sprawdzenie, czy infos są w konfiguracji powiązanej:
				CINFO=u""
				"""
				if IS_SM_C:
					#sprawdzenie konfiguracji zagiętej
					NOT_SM=get_data_from_model(self._modeldoc,self._dane['config'][0:len(self._dane['config'])-18])
					NOT_SM_P=part(NOT_SM)
					NOT_SM_P.open()
					NOT_SM_P.activate()
					CINFO=NOT_SM_P.get_ci(I)
					if len(CINFO)==0:
						STATUS=0
					NOT_SM_P.save()
					NOT_SM_P.close()
					self.open()
					self.activate()
				if HAS_SM_C:
					#sprawdzenie konfiguracji rozlozonej
					self.save()
					SM=get_data_from_model(self._modeldoc,self._dane['config']+u"sm-rozłożony-model")
					SM_P=part(SM)
					SM_P.open()
					SM_P.activate()
					CINFO=SM_P.get_ci(I)
					if len(CINFO)==0:
						STATUS=0
					SM_P.save()
					SM_P.close()
					self.open()
					self.activate()
				if len(CINFO)>0:
					continue
					#if IS_SM_C:
					#	print "\nustawiam wartość ",I," z konfiguracji ",self._dane['config'][0:len(self._dane['config'])-18]
					#if HAS_SM_C:
					#	print "\nustawiam wartość ",I," z konfiguracji ",self._dane['config']+u"sm-rozłożony-model"
					#self.set_ci(I,CINFO)
				elif prompt:
				"""
				if I == 'rev':
					self.set_ci(I,'01')
					self.set_ci('Rev.:','01')
					#config.set(self['name'],'rev','01')
					continue
				else:
					print "\n\nUWAGA!!>>\n",("plik").ljust(10),self['plik'],("\nkonfiguracja ").ljust(10),self['config'],("\nbrak pola ").ljust(10), I
					VALUE = raw_input("    podaj wartość > ")
					self.set_ci(I,VALUE)
					#config.set(self['name'],I,VALUE)
					print "\n"
			self.RACcopy(I)
			#if self.is_sm_config():
		#sprawdzenie, czy nie ma tego w konfiguracji powiązanej:
		self.set_sm_uwagi()
				
		if prompt:
			self.save()
		return STATUS#, config
	def RACcopy(self,I):
		#funkcja do kopiowania properities z starych templatów do RAStemplatów:
		#['description','note1','rev']
		if I=='description':
			self.set_ci('Title:',self.get_ci('description'))
		elif I=='rev':
			self.set_ci('Rev.:',self.get_ci(I))
		elif I=='Quantity':
			self.set_ci('Quantity:',self.get_ci(I))
		elif I=='heat':
			self.set_ci('Heat Treatment:',self.get_ci(I))
		elif I=='chemical':
			self.set_ci('Chemical Treatment:',self.get_ci(I))
		elif I=='Standard':
			self.set_ci('Measurements without tolerances:',self.get_ci(I))
		elif I=='Mass':
			self.set_ci('Weight',self.get_ci(I))
	def check_mand_infos(self,prompt=1,mand_infos=['Quantity','Date','Creator','Material','Standard','Mass','Surface Area']):
		#sprawdza, czy w pliku rysunku i modeli są zachowane odpowiednie informacje - zmienia je za każdym razem
		MAND_INFOS=mand_infos
		##########################################
		#sekcja z konfiguracją - nieaktualna!!!
		#if not config==None:
		#	#import ConfigParser
		#	if not config.has_section(self['name']):
		#		config.add_section(self['name'])
		#		config.set(self['name'],'file',self['full_path'])
		##########################################
		STATUS=1	#zwraca 1 jeśli wszystko jest, 0 jak nie
		self.open()
		self.activate()
		self._modeldoc.ShowConfiguration(self['config'])
		for I in MAND_INFOS:
			#sprawdzenie, czy infos są w konfiguracji powiązanej:
			CINFO=u""
			if I.lower() == 'date':
				TIME = time.strftime("%d.%m.%y")
				self.set_ci(I,TIME)
				#config.set(self['name'],'date',TIME)
			elif I.lower() == 'creator':
				self.set_ci(I,USER)
				#config.set(self['name'],'creator',USER)
			elif I.lower() == 'standard':
				self.set_ci(I,'DS/ISO 2768-1 f')
			elif I.lower() in ['mass','weight']:
				self.set_ci(I,'\"SW-Mass@@'+self['config']+'@'+self['plik']+'\"')
			elif I.lower() in ['surface area','area']:
				self.set_ci(I,'\"SW-SurfaceArea@@'+self['config']+'@'+self['plik']+'\"')
			elif I.lower() == 'material':
				self.set_ci(I,'\"SW-Material@@'+self['config']+'@'+self['plik']+'\"')
			elif prompt:
				print "\n\nUWAGA!!>>\n",("plik").ljust(10),self['plik'],("\nkonfiguracja ").ljust(10),self['config'],("\nbrak pola ").ljust(10), I
				VALUE = raw_input("    podaj wartość > ")
				self.set_ci(I,VALUE)
				#config.set(self['name'],I,VALUE)
				#self.set_ci(I,raw_input("    podaj wartość > "))
				print "\n"
			self.RACcopy(I)
				
		if prompt:
			self.save()
		return STATUS#, config

	def check_const_infos(self,mand_infos={'Creator':'LDU'}):
		#sprawdza, czy w pliku rysunku i modeli są zachowane odpowiednie informacje - zmienia je za każdym razem
		MAND_INFOS=mand_infos
		STATUS=1	#zwraca 1 jeśli wszystko jest, 0 jak nie
		self.open()
		self.activate()
		self._modeldoc.ShowConfiguration(self['config'])
		prompt=1
		for I in MAND_INFOS:
			#sprawdzenie, czy infos są w konfiguracji powiązanej:
			CINFO=u""
			if prompt:
				if I == 'Date':
					self.set_ci(I,time.strftime("%d.%m.%y"))
				else:
					#print "\n\nUWAGA!!>>\n",("plik").ljust(10),self['plik'],("\nkonfiguracja ").ljust(10),self['config'],("\nbrak pola ").ljust(10), I
					self.set_ci(I,MAND_INFOS[I])
					print "\n"
				
		if prompt:
			self.save()
		return STATUS
	def is_sm_config(self):
		if u"SM-ROZŁOŻONY-MODEL" in self._dane['config']:
			return 1
		else:
			return 0
	def get_configs(self):
		#pobiera nazwy wszystkich konfiguracji:
		return self._modeldoc.GetConfigurationNames()
	def has_sm_config(self):
		#sprawdza, czy część ma konfigurację rozłożony model
		if self.is_sm_config():
			return 0
		C=self.get_configs()
		if C==None:
			return 0
		for c in C:
			if c==self._dane['config']+u"SM-ROZŁOŻONY-MODEL":
				return 1 
	def set_sm_uwagi(self):
		#funkcja ustawia uwagi i rys_laser jeśli element ma konfigurację ...sm-rozłożony-model
		if not is_sw_file(self._dane['plik'])==1:
			return
		IS_SM_C=self.is_sm_config()
		#HAS_SM_C=self.has_sm_config()
		PartNo=self.get_ci('PartNo')
		SubPartNo=self.get_ci('SubPartNo')
		if IS_SM_C:
			self.save()
			NOT_SM=get_data_from_model(self._modeldoc,self._dane['config'][0:len(self._dane['config'])-18])
			NOT_SM_P=part(NOT_SM)
			NOT_SM_P.open()
			NOT_SM_P.activate()
			UWAGI=NOT_SM_P.get_ci('uwagi')
			HAS_UWAGI=len(UWAGI)>0
			LN=NOT_SM_P.get_ci('LaserPartNo')
			HAS_LN=len(LN)>0
			#SUR=NOT_SM_P.get_ci('surf')	#wykonczenie pow.
			HAS_LN=len(LN)>0
			if not HAS_UWAGI:
				UWAGI="rys. laser "+PartNo
				if len(SubPartNo)>0:
					UWAGI=UWAGI+"."+SubPartNo
				NOT_SM_P.set_ci("uwagi",UWAGI)
			if not HAS_LN:
				LN=PartNo
				NOT_SM_P.set_ci("LaserNo",PartNo)
				NOT_SM_P.set_ci("SubLaserNo",SubPartNo)
				NOT_SM_P.set_ci("LaserPartNo",PartNo+"."+SubPartNo)
			NOT_SM_P.save()
			NOT_SM_P.close()
		self.open()
		self.activate()
	def set_colors(self):
		#pyta lub ustawia kolor (wykończenie powierzchni)
		if is_sw_file(self._dane['plik'])==3:
			return
		elif is_sw_file(self._dane['plik'])==1:
			return
	def set_project_config(self,dane):
		#met. tworzy lub edytuje plik project.conf w danym katalogu z informacjami pomocnymi przy tworzenia dokumentacji technicznej
		FILENAME='project.conf'
		PATH=self['path']
		#sprawdzenie, czy w katalogu projektu znajduje się plik project.conf
		FULL_PATH = PATH+'\\'+FILENAME
		FILE=open(FULL_PATH)
		CONF=ConfigParser.ConfigParser()
		if os.path.isfile(FULL_PATH):
			#plik już istnieje, trzeba sprawdzić, czy są wszystkie infos:
			CONF.read(FULL_PATH)
		#numer projektu:
		if not CONF.has_section('project'):
			CONF.add_section('project')
		if CONF.has_option('project','number'):
			PROJ_NO=CONF.get('project','number')
			print 'w pliku konfiguracyjnym istnieje numer projektu - '+PROJ_NO
			raw_input('')
	def material(self):
		#odczytuje material z pliku czesci
		self.open(); self.activate()
		nf = self._modeldoc.GetFeatureCount()
		F=self._modeldoc.FirstFeature()
		for i in range(0,nf):
			if F.GetTypeName=='MaterialFolder':
				mat = F.Name
				if not mat:
					mat = PRT.get_ci('Material')
				if not mat or mat == "Material <not specified>":
					mat="N.A."
				return mat
			F=F.GetNextFeature
		#jesli nie znalazl operacji z materialem:
		return "N.A."
	def set_units(self):
		#ustawia jednostki na MMGS (dodano w wersji 1.18, 2012-08-09)
		#swUnitSystem_MMGS             =5          # from enum swUnitSystem_e
		#swUnitSystem                  =263        # from enum swUserPreferenceIntegerValue_e
		#cytat z drukowanie_LD (dot. masy):
		#BoolStatus = pModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, 0, swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms)
		BoolStatus = self._modeldoc.Extension.SetUserPreferenceInteger(263, 0, 5)
	def getFeatureByName(self, featName):
		#pobiera ceche na podst jej nazwy
		self.open(); self.activate()
		nf = self._modeldoc.GetFeatureCount()
		F=self._modeldoc.FirstFeature()
		for i in range(0,nf):
			if F.Name==featName:
				return F
			F=F.GetNextFeature
	def skipInstances(self, featName, instancesToSkip):
		#funkcja do pomijania wystapien w szyku solidworksa
		swFeat = self.getFeatureByName(featName)
		swFeat = sldmod.IFeature(swFeat)
		swLocPatt = swFeat.GetDefinition()
		swLocPatt = sldmod.ILinearPatternFeatureData(swLocPatt)
		M = sldmod.IModelDoc(self._modeldoc)
		swLocPatt.AccessSelections(M, None)
		swLocPatt.SkippedItemArray = instancesToSkip
		#swLocPatt.ISetSkippedItemArray(2,instancesToSkip)
		swFeat.ModifyDefinition(swLocPatt, M, None)
	def editSketch(self, name):
		#sets sketch sketchName to edit mode
		SketchFeat = self.getFeatureByName(name)
		print SketchFeat
		SketchFeat.Select2(False, 0)
		#Sketch = Sketch.GetSpecificFeature2()
		#value = self._modeldoc.Extension.SelectByID2(name, "SKETCH", 0, 0, 0, False, 0, None, 0)
		self._modeldoc.EditSketch()
	def finishEditingSketch(self):
		#finishes editing currently edited sketch
		self._modeldoc.InsertSketch2(True)
	def getBox(self):
		#gets part box (uses API function)
		return self._modeldoc.GetBox
	def link(self):
		#inserts link to open doc to clipboard
		FULL_PATH = self['full_path']
		win32clipboard.OpenClipboard()
		win32clipboard.EmptyClipboard()
		win32clipboard.SetClipboardText('<file://'+FULL_PATH+'>', win32clipboard.CF_UNICODETEXT)
		win32clipboard.CloseClipboard()

def compare_filedate(file1, file2):
	#porownanie dat dwoch plikow, jesli data modyfikacji file1 jest pozniejsza to zwraca 1, jezeli data file2 jest pozniejsza to zwraca 2, jesli sa rowne to zwraca 0
	stats1 = os.stat(file1); lastmod_date1 = stats1[8]
	stats2 = os.stat(file2); lastmod_date2 = stats2[8]
	if lastmod_date1>lastmod_date2:
		return 1
	elif lastmod_date2<lastmod_date1:
		return 2
	else:
		return 0
	
def get_data_from_drawing(MOD=None):
	#pobiera dane z rysunku lub otwartego dokumentu i zwraca je
	#otwarcie dokumentu:
	if not MOD==None:
		DOC=MOD
	else:
		DOC = sldmod.IModelDoc2(swx.ActiveDoc)            # get active document
	#słownik opisujący ten dokument:
	thispart={}
	SCIEZKA=(DOC.GetPathName())
	if not is_sw_file(SCIEZKA)==3:
		return None
	thispart['plik']=SCIEZKA[SCIEZKA.rindex('\\')+1:]
	thispart['name']=thispart['plik'][0:thispart['plik'].rindex(".")]
	thispart['katalog']=SCIEZKA[0:SCIEZKA.rindex('\\')+1]
	thispart['path']=thispart['katalog']
	thispart['full_path']=SCIEZKA
	thispart["config"]=""
	#teraz dane z modelu:
	thispart['PartNo']=(DOC.GetCustomInfoValue("","PartNo"))
	thispart['SubPartNo']=(DOC.GetCustomInfoValue("","SubPartNo"))
	return draw(thispart)


#stałe używane przez klasę <draw>
FORMATY = {'A4':9, 'A3':8,'A2':66, 'A1':8, 'A0':8}
SIZES = {'210x297':['A4','p'], '297x210':['A4','l'], '420x297':['A3','l'], '297x420':['A3','p'], '594x420':['A2','l'], '420x594':['A2','p'], '594x841':['A1','p'], '841x594':['A1','l'], '1189x841':['A0','l'], '841x1189':['A0','p']}

"""
	swDwgPapersUserDefined        =12         # from enum swDwgPaperSizes_e
	swDwgTemplateA0size           =11         # from enum swDwgTemplates_e
	swDwgTemplateA1size           =10         # from enum swDwgTemplates_e
	swDwgTemplateA2size           =9          # from enum swDwgTemplates_e
	swDwgTemplateA3size           =8          # from enum swDwgTemplates_e
	swDwgTemplateA4size           =6          # from enum swDwgTemplates_e
	swDwgTemplateA4sizeVertical   =7          # from enum swDwgTemplates_e
	swDwgTemplateAsize            =0          # from enum swDwgTemplates_e
	swDwgTemplateAsizeVertical    =1          # from enum swDwgTemplates_e
	swDwgTemplateCustom           =12         # from enum swDwgTemplates_e
"""
swPAPERSIZES = [""]

class empty_config:
	def __init__(self):
		A = None
	def has_section(self,section):
		return False
	def has_option(self,section,option):
		return False

class draw(part):
	#klasa do zarządzania rysunkami:
	def __init__(self,dane):
		part.__init__(self,dane)
		self._parts=[]
		self._drawingdoc = None
	def open(self):
		part.open(self)
		#self._drw_modeldoc=sldmod.IDrawingDoc(self._modeldoc)
	def active(self):
		part.active(self)
		#self._drw_modeldoc=sldmod.IDrawingDoc(self._modeldoc)
	def get_parts(self):
		#otwiera swój plik i pobiera dane o tym, co jest na nim narysowane
		self.open()
		self.activate()
		DOC=self._modeldoc
		DDOC=sldmod.IDrawingDoc(DOC)
		#/////////////////////////////////////////////////////////////////////////////////////////////////////////
		#print DDOC.GetSheetNames()
		#V=sldmod.IView(DDOC.GetFirstView())
		#V=sldmod.IView(V.GetNextView())
		#V=sldmod.IView(DDOC.ActiveDrawingView)

		#print "kupa hehe ", V
		#raw_input()
		#COMPS=[]
		#while 1:
			#C=sldmod.IModelDoc2(V.ReferencedDocument)
			#C=V.ReferencedConfiguration
			#C=V.Position
			#CC=V.GetReferencedModelName
			#CC=V.ScaleRatio
			#CC=V.LoadModel()
			#CC=V.GetName2
			#print CC
			#raw_input()
			#print COMPS
			#if CC==None:
			#	break
			#COMPS.append(CC)
			#print V
			#V=V.GetNextView()
		#adoc.SetAddToDB(1)              # add ents directly todatabase
		#DOC.SetDisplayWhenAdded(0)     # don't display ents untilredraw
		#/////////////////////////////////////////////////////////////////////////////////////////////////////////
		if len(self._parts)>0:
			return self._parts
		PARTS=[]; PARTNAMES=[]
		f=DOC.FirstFeature()
		while not f==None:
			#f=sldmod.IFeature(f)
			#PartNo	Text	DIN 912	DIN 912
			#print f.GetTypeName
			sf=f.GetFirstSubFeature
			while not sf==None:
				#sf=sldmod.IFeature(sf)
				#print "    ", sf.GetTypeName
				#print "    ", sf.Name
				#sf.Select2
				#print (sf.GetTypeName).encode('cp1250')
				#raw_input()
				if sf.GetTypeName in ["AbsoluteView","RelativeView", "UnfoldedView"]:
					#print "widokok"
					#widok nadrzędny
					#PartNo	Text	DIN 912	DIN 912
					
					DDOC.ActivateView(sf.Name)
					V=sldmod.IView(DDOC.ActiveDrawingView)
					#print V.Position
					#print V.LoadModel()	#zwraca 0
					COMPONENT=sldmod.IModelDoc2(V.ReferencedDocument)
					CONFIG= V.ReferencedConfiguration
					PRT=part(get_data_from_model(COMPONENT,CONFIG))
					#if PRT._modeldoc.GetPathName() not in PARTNAMES:
					#print PRT['plik']; raw_input()
					if PRT['plik']+PRT['config'] not in PARTNAMES:
						PARTS.append(PRT)
						#PARTNAMES.append(PRT._modeldoc.GetPathName())
						PARTNAMES.append(PRT['plik']+PRT['config'])
						#PartNo	Text	DIN 912	DIN 912
					#ssf=sf.GetFirstSubFeature
					#while not ssf==None:
					#	print "        ", ssf.GetTypeName
					#	ssf=ssf.GetNextSubFeature
					#V=sldmod.IView(sldmod.IFeature(sf).GetDefinition)
					#print V.Position
					#raw_input()
				sf=sf.GetNextSubFeature
				#PartNo	Text	DIN 912	DIN 912
				
			#Ch=f.GetChildren
			#for O in Ch:
				#print(2*"         "+O.GetTypeName)
			f=f.GetNextFeature
		#print "koniec"
		self._parts=PARTS
		return 
	def check_nums(self):
		#sprawdza, czy są ustawione wszystkie numery
		if len(self._parts)==0:
			self.get_parts()
		NUMER=self.get_ci("PartNo")
		USTAW=0	#zm. logiczna czy ustawiać numery od nowa
		if not numerator.is_valid(NUMER):
			USTAW=1
		if USTAW:
			self.set_drw_no_s()
			return 1
		self.get_parts()
		SUBNUMERY_P=[]
		for i in range(0,len(self._parts)):
			NUMER_P=self._parts[i].get_ci("PartNo")
			SUBNUMER_P=self._parts[i].get_ci("SubPartNo")
			if (not numerator.is_valid(NUMER_P)) or (not NUMER_P==NUMER) or (SUBNUMER_P in SUBNUMERY_P):
				USTAW=1
			SUBNUMERY_P.append(SUBNUMER_P)
		if USTAW:
			self.activate()
			self.set_drw_no_s()
			return 1
		#zamknięcie wszystkich modeli:
		for i in range(0,len(self._parts)):
			self._parts[i].close()
		return USTAW	#jak 0 to nie ustawiał numerów
	def set_drw_no_s(self):
		#funkcja ustawia numery rysunków
		#najpierw trzeba sprawdzić, czy jest jakiś numer rysunku
		#po co hehe
		NUM=numerator.numerator()
		if len(self._parts)==0:
			self.get_parts()
		N=NUM.add_no()
		#print self._parts
		for i in range(0,len(self._parts)):
			#sprawdzenie, czy plik ma jakiś numer już:
			if len(self._parts[i].get_ci("PartNo"))>0:
				self._parts[i].set_ci("__OldPartNo",self._parts[i].get_ci("PartNo"))
			if len(self._parts[i].get_ci("SubPartNo"))>0:
				self._parts[i].set_ci("__OldSubPartNo",self._parts[i].get_ci("SubPartNo"))
			#
			self._parts[i].set_ci("PartNo",NUM.no_str())
			NUM.add_subno()
			self._parts[i].set_ci("SubPartNo",NUM.subno_str())
			self._parts[i].save()
			self._parts[i].close()
		self.set_ci("PartNo",NUM.no_str())
		self.save(0x10)
		if PRINT:
			print "koniec ustawiania numerów"

	def check_ness_infos(self,prompt=1,NESS_INFOS=['Description']):
		#sprawdza, czy w pliku rysunku i modeli są zachowane odpowiednie informacje
		#STATUS=[part.check_ness_infos(self,prompt,NESS_INFOS)]
		CHECK_SELF = 0
		STATUS=[]
		self.get_parts()
		for PART in self._parts:
			STATUS_ = PART.check_ness_infos(prompt)
			STATUS.append(STATUS_)
			STATUS_ = PART.check_mand_infos(prompt)
			STATUS.append(STATUS_)
			#STATUS.append(PART.check_const_infos({'Creator':'PAZ'}))
		#teraz rysunek:
		self.open()
		self.activate()
		if CHECK_SELF:
			for I in NESS_INFOS:
				CINFO=self.get_ci(I)
				if (len(CINFO.strip())==0):
					if prompt:
						print "\n\nUWAGA!!>>",("\nplik ").ljust(15),(self['plik']),("\nbrak pola:").ljust(15), I
						CINFO=raw_input("    podaj wartość > ")
						print "\n"
					#print self._parts; raw_input()
					if (len(CINFO)==0) and len(self._parts)==1:
						#kopiowanie CINFO z pierwszego widoku rysunku
						#P1=self._parts[0]
						#P1.open(); P1.activate()
						#CINFO=P1.get_ci(I)
						CINFO='$PRPSHEET:"'+I+'"'
					
				self.set_ci(I,CINFO)
		return max(STATUS)
	def close_parts(self):
		for P in self._parts:
			P.close()
	def laser_filename(self):
		#zwraca nazwe pliku jaką zrzucam na laser
		#nazwa_pliku-numer_rys.dwg
		return self['name']+"-"+self.get_ci("PartNo")+".dwg"
	def saveas(self,katalog=".",nazwa_pliku="",format="dwg"):
		self.open()
		self.activate()
		"""
		swSaveAsOptions_AvoidRebuildOnSave=0x8        # from enum swSaveAsOptions_e
		swSaveAsOptions_Copy          =0x2        # from enum swSaveAsOptions_e
		swSaveAsOptions_DetachedDrawing=0x80       # from enum swSaveAsOptions_e
		swSaveAsOptions_OverrideSaveEmodel=0x20       # from enum swSaveAsOptions_e
		swSaveAsOptions_SaveEmodelData=0x40       # from enum swSaveAsOptions_e
		swSaveAsOptions_SaveReferenced=0x4        # from enum swSaveAsOptions_e
		swSaveAsOptions_Silent        =0x1        # from enum swSaveAsOptions_e
		swSaveAsOptions_UpdateInactiveViews=0x10

		swSaveAsCurrentVersion        =0x0        # from enum swSaveAsVersion_e
		swSaveAsDetachedDrawing       =0x4        # from enum swSaveAsVersion_e
		swSaveAsFormatProE            =0x2        # from enum swSaveAsVersion_e
		swSaveAsSW98plus              =0x1        # from enum swSaveAsVersion_e
		swSaveAsStandardDrawing       =0x3        # from enum swSaveAsVersion_e
		"""
		if format=="dwg":
			if len(nazwa_pliku)==0:
				nazwa_pliku=self.laser_filename()
			if katalog==".":
				katalog=self['katalog']
			#print katalog
			#print nazwa_pliku
			#ST=self._modeldoc.SaveAs4(katalog+nazwa_pliku,0,0x2+0x1+0x8,None,None)
			#print self._modeldoc.GetPathName()
			ST=self._modeldoc.SaveAs2(katalog+nazwa_pliku,0,True,False)
			#print "status: ",ST
		elif format == 'pdf':
			if len(nazwa_pliku) == 0:
				nazwa_pliku = self['name']
	def create_folders(self, folders = ['archive', 'PDF_DXF_DWG','Analyse','Data sheets','Electronics','Other','Parts list','Standards']):
		#tworzy w katalogu projektu foldery podane w liscie, zastepuje funkcje "archive"
		if folders in [1,'a']:
			folders = ['archive']
		elif isinstance(folders,str):
			folders = [folders]
		PATH=self['path']
		for F in folders:
			if not os.path.isdir(PATH+F):
				os.mkdir(PATH+F)
				
			
	def archive(self):
		#sprawdza, czy jest katalog archive i inne standardowe katalogi w tym samym katalogu co rysunek, jeśli nie to go tworzy
		#jesli zmianna all = 1 to tworzy wszystkie katalogi, jesli all = 0 to tylko archive
		PATH=self['path']
		#print PATH; raw_input()
		if not os.path.isdir(PATH+'archive'):
			os.mkdir(PATH+'archive')
		if not os.path.isdir(PATH+'PDF_DXF_DWG'):
			os.mkdir(PATH+'PDF_DXF_DWG')
		if not os.path.isdir(PATH+'Analyse'):
			os.mkdir(PATH+'Analyse')
		if not os.path.isdir(PATH+'Data sheets'):
			os.mkdir(PATH+'Data sheets')
		if not os.path.isdir(PATH+'Electronics'):
			os.mkdir(PATH+'Electronics')
		if not os.path.isdir(PATH+'Other'):
			os.mkdir(PATH+'Other')
		if not os.path.isdir(PATH+'Parts list'):
			os.mkdir(PATH+'Parts list')
		if not os.path.isdir(PATH+'Pneumatic'):
			os.mkdir(PATH+'Pneumatic')
		if not os.path.isdir(PATH+'Standards'):
			os.mkdir(PATH+'Standards')
	def close(self):
		swx.QuitDoc(self['katalog']+self['plik'])
		#swx.CloseDoc(self['katalog']+self['plik'])
		#self._modeldoc.Close()
		self._modeldoc=None
	#def revision(self):
	def material(self):
		self.get_parts()
		if len(self._parts)==1:
			MAT = self._parts[0].material()
			self._parts[0].close()
		else:
			MAT=[]
			for i in range(0,len(self._parts)):
				MAT.append(self._parts[i].material())
				self._parts[i].close()
		self.open(); self.activate()
		return MAT
	def set_current_pdf(self):
		#sprawdza, czy jest aktualny pdf, jak nie to eksportuje do pdf
		plik_pdf = self['path']+"\\PDF_DXF_DWG\\"+self['name']+".pdf"
		if not os.path.isdir(self['path']+"\\PDF_DXF_DWG\\"):
			os.mkdir(self['path']+"\\PDF_DXF_DWG")
		if not os.path.isfile(plik_pdf):
			ST=self._modeldoc.SaveAs2(plik_pdf,0,True,False)
		elif os.stat(self['full_path'])[8]>os.stat(plik_pdf)[8]:
			ST=self._modeldoc.SaveAs2(plik_pdf,0,True,False)
		elif os.path.isfile(self['path']+"/"+self["name"]+".sldprt") and os.stat(self['path']+"/"+self["name"]+".sldprt")[8]>os.stat(plik_pdf):
			ST=self._modeldoc.SaveAs2(plik_pdf,0,True,False)
		elif os.path.isfile(self['path']+"/"+self["name"]+".sldasm") and os.stat(self['path']+"/"+self["name"]+".sldasm")[8]>os.stat(plik_pdf):
			ST=self._modeldoc.SaveAs2(plik_pdf,0,True,False)
		#print plik_pdf
	def set_current_dxf(self):
		#sprawdza, czy jest aktualny dxf, jak nie to eksportuje do dxf
		plik_dxf = self['path']+"\\PDF_DXF_DWG\\"+self['name']+".dxf"
		if not os.path.isdir(self['path']+"\\PDF_DXF_DWG\\"):
			os.mkdir(self['path']+"\\PDF_DXF_DWG")
		if not os.path.isfile(plik_dxf):
			ST=self._modeldoc.SaveAs2(plik_dxf,0,True,False)
		elif os.stat(self['full_path'])[8]>os.stat(plik_dxf)[8]:
			ST=self._modeldoc.SaveAs2(plik_dxf,0,True,False)
		elif os.path.isfile(self['path']+"/"+self["name"]+".sldprt") and os.stat(self['path']+"/"+self["name"]+".sldprt")[8]>os.stat(plik_dxf):
			ST=self._modeldoc.SaveAs2(plik_dxf,0,True,False)
		elif os.path.isfile(self['path']+"/"+self["name"]+".sldasm") and os.stat(self['path']+"/"+self["name"]+".sldasm")[8]>os.stat(plik_dxf):
			ST=self._modeldoc.SaveAs2(plik_dxf,0,True,False)
		#print plik_pdf
	def get_pagesize(self):
		#funkcja pobiera rozmiar arkusza oraz orientację strony (dodano w wersji 1.08, 2010-08-16)
		self._drawingdoc = sldmod.IDrawingDoc(self._modeldoc)
		S = self._drawingdoc.GetCurrentSheet() #obiekt aktualnego arkusza
		PROPS = S.GetProperties
		#print str(int(PROPS[5]*1000.))+'x'+str(int(PROPS[6]*1000.))
		return str(int(PROPS[5]*1000.))+'x'+str(int(PROPS[6]*1000.))
		#PAGESETUP = self._modeldoc.PageSetup
		#print "Rozmiar: %d" % PAGESETUP.PrinterPaperSize
		#print PAGESETUP.Orientation
	def get_pagesize_num(self):
		S = self.get_pagesize()
		return [float(S[:S.index("x")]),float(S[S.index("x")+1:])]
	def get_num_pagesize(self):
		#funkcja pobiera rozmiar arkusza oraz orientację strony (dodano w wersji 1.18, 2012-08-08)
		self._drawingdoc = sldmod.IDrawingDoc(self._modeldoc)
		S = self._drawingdoc.GetCurrentSheet() #obiekt aktualnego arkusza
		PROPS = S.GetProperties
		return [PROPS[5]*1000.,PROPS[6]*1000.]
	def get_sheets(self):
		#pobiera wszystkie arkusze rysunku
		"hehe"
	def round_sheetsize(self):
		#zwraca ciag znakow "AX" najblizszego formatu arkusza
		S = self.get_pagesize()
		X = int(S[:S.index("x")])
		Y = int(S[S.index("x")+1:])
		if X>Y:
			#Landscape
			Xb = 297	#X formatu A4
			Yb = 210	#Y formatu A4
			FORMATS_AREAS = ['297x210', '420x297', '594x420', '841x594', '1189x841']
		else:
			#Portrait
			Xb = 210	#X formatu A4
			Yb = 297	#Y formatu A4
			FORMATS_AREAS = ['210x297', '297x420', '420x594', '594x841', '841x1189']
		#iX = round(float(X)/float(Xb)) #ile razy Xb mieści się w X
		#iY = round(float(Y)/float(Yb)) #ile razy Yb mieści się w Y
		iX = float(X)/float(Xb) #ile razy Xb mieści się w X
		iY = float(Y)/float(Yb) #ile razy Yb mieści się w Y
		#print "IX = "+str(iX)+", iY = "+str(iY)
		AREA = iX*iY
		import math
		n = int(round(math.log(AREA)/math.log(2)))+1	#n-ty wyraz ciągu geometrycznego 1, 1*2, 1*2*2 itd
		#print "Format: "+FORMATS_AREAS[n-1]
		return FORMATS_AREAS[n-1]
	def set_pagesetup(self, config=empty_config()):
		#funkcja ustawia preferencje drukowania na podst danych z obiektu <config> - obiekt typu ConfigParser.ConfigParser
		#określenie rozmiaru i orientacji papieru:
		PAGESIZE = self.get_pagesize()
		if not PAGESIZE in SIZES:
			print("niestandardowy rozmiar papieru.")
			PAGESIZE = self.round_sheetsize()
		SIZE = SIZES[PAGESIZE]
		FIT = True
		#print SIZE
		#printer = '\\\\http://szczecin\Szczecin - Printer 67 : LaserJet 5200'
		printer = '\\PL01W03W04\Szczecin - Printer 67 : LaserJet 5200'
		#printer = '\\\\http://szczecin\Szczecin - Printer 3 : TOSHIBA e-STUDIO281c'
		color = 3
		#swPageSetup_AutomaticDrawingColor=1          # from enum swPageSetupDrawingColor_e
		#swPageSetup_BlackAndWhite     =3          # from enum swPageSetupDrawingColor_e
		#swPageSetup_ColorGrey         =2          # from enum swPageSetupDrawingColor_e
		if config.has_section(SIZE[0]):
			if config.has_option(SIZE[0],'printer'):
				printer = config.get(SIZE[0],'printer')
			if config.has_option(SIZE[0],'color'):
				color = int(config.get(SIZE[0],'color').strip())
		elif config.has_section('general'):
			if config.has_option('general','printer'):
				printer = config.get('general','printer')
			if config.has_option('general','color'):
				color = int(config.get('general','color').strip())
		
		if config.has_section(SIZE[0]):
			if config.has_option(SIZE[0],'size'):
				SIZE[0] = config.get(SIZE[0],'size')
				FIT = True
		#print SIZE[0]
		#print printer
		#print self._modeldoc.Printer
		#raw_input()
		self._modeldoc.Printer = printer
		#print self._modeldoc.Printer
		#raw_input()
		PS = self._modeldoc.PageSetup
		ORIENTS = {'p':1, 'l':2}
		orientation = ORIENTS[SIZE[1]]
		PS.Orientation = orientation
		PS.ScaleToFit = FIT
		PS.PrinterPaperSize = FORMATY[SIZE[0]]
		PS.DrawingColor = color
		#self.save()
	def Print(self, config = empty_config()):
		self.set_pagesetup(config)
		self._modeldoc.PrintDirect()
	def orientation(self):
		#zwraca orientacje strony(dodano w wersji 1.18, 2012-08-08)
		SIZE = self.get_num_pagesize()
		if SIZE[0]>SIZE[1]:
			return "l"
		else:
			return "p"
	def GetSheetCount(self):
		self._drawingdoc = sldmod.IDrawingDoc(self._modeldoc)
		return self._drawingdoc.GetSheetCount()
	def replace_sheet_format(self, config=empty_config()):
		#zmienia format arkusza na ten podany w pliku, rozmiar będzie powielony z aktualnego (dodano w wersji 1.18, 2012-08-09)
		o = self.orientation()
		SIZE = self.round_sheetsize()
		SIZE2 = self.get_pagesize_num()
		S = SIZES[SIZE][0]
		FNAME = config.get("general","templates_folder") + "\\"
		if o == "l":
			FNAME+=config.get(S,"sheet_landscape")
		elif o == "p":
			FNAME+=config.get(S,"sheet_portrait")
		#print FNAME
		self._drawingdoc = sldmod.IDrawingDoc(self._modeldoc)
		S = self._drawingdoc.GetCurrentSheet() #obiekt aktualnego arkusza
		N = S.GetName
		PROPS = S.GetProperties
		width0 = float(SIZE[:SIZE.index("x")])/1000.
		height0 = float(SIZE[SIZE.index("x")+1:])/1000.
		width = SIZE2[0]/1000.
		height = SIZE2[1]/1000.
		if abs(width-width0)/width0>0.1 or abs(height-height0)/height0>0.1:
			width = width0
			height = height0
		#print SIZE
		#print width
		#print height
		boolstatus = self._drawingdoc.SetupSheet5(N, 12, 12, PROPS[2], PROPS[3], True, FNAME, width, height, "Default", True)
		boolstatus = self._modeldoc.ForceRebuild3(True)
		#S.SetTemplateName("F:\Common\MDC\PL\General\Medical Engineering\DEPT\Tools for SolidWorks\PL-MDC-Template-sld-drawings\\temp_PL_a3_20120719.slddrt")
		#boolstatus = Part.SetupSheet5("Sheet1", 12, 12, 1, 4, False, "f:\common\mdc\pl\general\medical engineering\dept\tools for solidworks\pl-mdc-template-sld-drawings\temp_pl_a3_20120719.slddrt", 0.5588, 0.4318, "Default", True)
	def replace_sheet_formats(self, config=empty_config()):
		self._drawingdoc = sldmod.IDrawingDoc(self._modeldoc)
		NAMES = self._drawingdoc.GetSheetNames()
		for N in NAMES:
			#self._drawingdoc.SheetNext()
			self._drawingdoc.ActivateSheet(N)
			self.replace_sheet_format(config)
			#self._modeldoc.EditRebuild
			self.save()
			
			
			
def set_draw_no(MOD=None):
	if Mod==None:
		#pobranie 
		a='kupa'

class partlist:
	#klasa zaw. listę cząści
	#ma działać maxymalnie szybko
	def __init__(self):
		self._lista=[]	#docelowa lista
		self._temp_lista=[]	#tymczasowa lista, w słownikach w niej nie ędzie pola ilość
		
	def add_part(self,part):
		if isinstance(part,list):
			self._temp_lista=self._temp_lista+part	#dodaje listę części
		else:
			self._temp_lista.append(part)	#dodaje do templisty nie dbając o to, ile razy wystąpił już dany komponent
	def compare(self,i,j):
		#porównuje dwie części o indeksach i w liście temp i j w liście uporządkowanej
		if (self._temp_lista[i]['plik']==self._lista[j]['plik']) and \
				(self._temp_lista[i]['katalog']==self._lista[j]['katalog']) and \
				(self._temp_lista[i]['config']==self._lista[j]['config']):
					return 1	#ta sama część
		else:
			return 0
		
	def pack(self):
		#przeszukuje listę tymczasową i przenosi jej elementy do listy uporządkowanej
		IND_DO_USUNIECIA=[]
		for i in range(0,len(self._temp_lista)):
			if len(self._lista)==0:
				self._lista.append(self._temp_lista[i])
				self._lista[len(self._lista)-1]['szt']=1
				
				POZ=self._temp_lista[i]['poz'][0]
				self._lista[len(self._lista)-1]['poz_count']={POZ:1}		#ile sztuk na daną pozycję
				WHERE=self._temp_lista[i]['whereis'][0]
				self._lista[len(self._lista)-1]['whereis_count']={WHERE:1}		#ile sztuk na dany podzespół
				continue
			for j in range(0,len(self._lista)):
				#print self.compare(i,j)
				if self.compare(i,j):
					IND_DO_USUNIECIA.append(i)
					if self._lista[j].has_key("szt"):
						self._lista[j]["szt"]+=1
					else:
						self._lista[j]["szt"]=1
					#dodanie do siebie pól mówiących, gdzie występuje część:
					POZ=self._temp_lista[i]['poz'][0]
					if self._temp_lista[i]["poz"][0] not in self._lista[j]["poz"]:
						self._lista[j]["poz"]=self._lista[j]["poz"]+[POZ]
						
					if not self._lista[j]['poz_count'].has_key(POZ):
						self._lista[j]['poz_count'][POZ]=1
					else:
						self._lista[j]['poz_count'][POZ]+=1

					
					WHERE=self._temp_lista[i]['whereis'][0]
					if WHERE not in self._lista[j]["whereis"]:
						self._lista[j]["whereis"]=self._lista[j]["whereis"]+[WHERE]
						
					if not self._lista[j]['whereis_count'].has_key(WHERE):
						self._lista[j]['whereis_count'][WHERE]=1
					else:
						self._lista[j]['whereis_count'][WHERE]+=1
					break
				elif j==len(self._lista)-1:
					IND_DO_USUNIECIA.append(i)
					self._lista.append(self._temp_lista[i])
					self._lista[len(self._lista)-1]['szt']=1
					
					POZ=self._temp_lista[i]['poz'][0]
					self._lista[len(self._lista)-1]['poz_count']={POZ:1}		#ile sztuk na daną pozycję

					WHERE=self._temp_lista[i]['whereis'][0]
					self._lista[len(self._lista)-1]['whereis_count']={WHERE:1}		#ile sztuk na dany podzespół
		#usunięcie wszystkich elementów z listy tymczasowej:
		self._temp_lista=[]
	def print_parts(self):
		#wyświetla wszystkie (spakowane) części
		print self._lista
		print self._temp_lista
		print "plik                 konfiguracja                             poz            ilosc"
		for i in range(0,len(self._lista)):
			print (self._lista[i]['plik']).ljust(20), (self._lista[i]['config']).ljust(35), str(self._lista[i]['poz']).ljust(15), (self._lista[i]['szt'])
		for i in range(0,len(self._lista)):
			print self._lista[i]['poz_count']
			print self._lista[i]['whereis_count']
		
class project:
	def __init__(self):
		self._numerator=numerator.numerator()
		self._tops=[]	#top złorzenia
		self._current_top=-1
		self._parts=partlist()		#uporządkowana lista części
		self._unpack_part_list={}	#nieuporządkowana lista części
	def get_excel_file(self):
		#funkcja sprawdza aktywny arkusz excela i odczytuje nazwe projektu, odgaduje katalog projektu
		#i może coś jeszcze
		try:
			NAZWA_PLIKU=self._plik_xls
			wkbook = excel.ActiveWorkbook
		except:
			#excel = win32com.client.Dispatch("Excel.Application")	#zainicjowane globalnie
			wkbook = excel.ActiveWorkbook
			nazwa_pliku=(wkbook.FullName)	#pełna nazwa pliku
			#wyciągnięcie scieżki:
			max_bs=nazwa_pliku.rindex("\\")
			self._katalog_projektu=nazwa_pliku[0:max_bs+1]
			self._plik_xls=nazwa_pliku[max_bs+1:]
		return wkbook
	def get_all_files(self):
		#tworzy listę wszystkich plików solida i wpisuje ją do słownika
		LISTA_PLIKOW_SW=[]
		GENERATOR=os.walk(self._katalog_projektu)	#obiekt typu generator
		KONIEC_PLIKOW=0
		while not KONIEC_PLIKOW:
			#KATALOG=GENERATOR.next()
			try:
				KATALOG=GENERATOR.next()
				for i in range(0,len(KATALOG[2])):
						LISTA_PLIKOW_SW.append(KATALOG[0][len(self._katalog_projektu):]+"\\"+KATALOG[2][i])
			except:
				KONIEC_PLIKOW=1
		self._pliki=LISTA_PLIKOW_SW
		return LISTA_PLIKOW_SW
	def get_tops(self):
		#szuka w katalogu złorzeni będących top-urządzeniami i wpisuje je do słownika
		#doc2=swx.OpenDoc6('e:\\lex\\sw\\blacha.SLDPRT',0x1,0,"Domyślna",Error, Warn)
		#///////////////////////////////////////////////////////////////////////////////
		#nie tak, ile to będzie trwało
		#trzeba samemu wpisać topy w excelu w tabelce
		#TOPS=[]
		#try:
		#	WSZYSTKIE_PLIKI=self._pliki
		#except:
		#	self.get_all_files()
		#	WSZYSTKIE_PLIKI=self._pliki
		#for i in range(0,len(WSZYSTKIE_PLIKI))
		#	if (is_sw_file(WSZYSTKIE_PLIKI[i])==2) and (not WSZYSTKIE_PLIKI[i] in self._tops):
		#		#mamy złorzenie które nie jest w topsach
		#		#DOC=swx.OpenDoc6(self._katalog_projektu+WSZYSTKIE_PLIKI[i])
		#///////////////////////////////////////////////////////////////////////////////
		wkbook=self.get_excel_file()
		sheet=wkbook.ActiveSheet
		#odczytanie tabelki z pola E5
		self._tops=[]; i=5
		while 1:
			if sheet.Cells(i,5).Value==None:
				break
			POZ=sheet.Cells(i,5).Value; PLIK=sheet.Cells(i,6).Value; KAT=sheet.Cells(i,7).Value; CONFIG=sheet.Cells(i,8).Value;
			if isinstance(POZ,float):
				POZ=int(POZ)
			if isinstance(KAT,float):
				KAT=int(KAT)
			elif KAT==None:
				KAT=""
			if isinstance(CONFIG,float):
				CONFIG=int(CONFIG)
			elif CONFIG==None:
				CONFIG="Domyślna"
			if isinstance(PLIK,float):
				PLIK=int(PLIK)
			POZ=str(POZ).encode("cp1250")
			CONFIG=str3(CONFIG)
			KAT=str(KAT).encode("cp1250")
			if (not len(KAT)==0) and (not KAT[len(KAT)-1]=="\\"):
				KAT+="\\"	#dodanie separatora
			PLIK=str(PLIK).encode("cp1250")
			self._tops.append({"poz":POZ,'katalog':KAT,'plik':PLIK,"config":CONFIG})
			i=i+1
	def find_parts(self,POZ):
		try:
			TOPY=self._tops
		except:
			self.get_tops()
			TOPY=self._tops
		for i in range(0,len(TOPY)):
			if TOPY[i]['poz']==POZ:
				DANE=TOPY[i]
				break
			elif i==len(TOPY)-1:
				print "brak pozycji ",POZ
				return
		PELNA_NAZWA=self._katalog_projektu+DANE['katalog']+DANE['plik']
		CONFIG=DANE['config']
		#zamknięcie wszystkich otwartych dokumentów solida:
		close_all_sw_docs()
		#otwarcie dokumentu:
		swx.OpenDoc6(PELNA_NAZWA,0x2,0,CONFIG,None,None)
		MODEL=sldmod.IModelDoc2(swx.ActiveDoc)
		#PARTS=find_parts(MODEL)
		PARTS=find_parts()
		close_all_sw_docs()
		#dodanie wspólnych kluczy do listy części:
		for i in range(0,len(PARTS)):
			PARTS[i]['poz']=[POZ]
		#zapisanie bazy części:
		self._unpack_part_list[POZ]=PARTS
		return PARTS
	def find_all_parts(self):
		try:
			TOPY=self._tops
		except:
			self.get_tops()
			TOPY=self._tops
		#teraz troche zabawy z solidem:
		#zamknięcie wszystkich otwartych dokumantów:
		#print "zamykam otwarte dokumenty"
		for i in range(0,len(TOPY)):
			self.find_parts(TOPY[i]['poz'])
	def pack_parts(self):
		for POZ in self._unpack_part_list:
			self._parts.add_part(self._unpack_part_list[POZ])
		self._parts.pack()
		self._parts.print_parts()
	def find_drw(self):
		#znajduje wszystkie pliki rysunków w katalogu projektu
		self._drw_file_list=[]
		LISTA_PLIKOW=self.get_all_files()
		for i in range(0,len(LISTA_PLIKOW)):
			if is_sw_file(LISTA_PLIKOW[i])==3:
				self._drw_file_list.append(LISTA_PLIKOW[i])
		#mamy listę plików rysunków, teraz trzebaby ją wzbogacić o informacje, jakie części znajdują się na rysunku

class PatternSketch:
    #class to perform operations on sketch that drives pattern
    #such sketch contains one solid line (for monodirectional patern) and sketch points, one of the points is placed on one endpoint of the line
    def __init__(self, sketchName, modelName = None):
        #if no modelName provided than currently open drawing, else open specific part
        if modelName == None:
            self.Model = get_data_from_model()
        else:
	        self.Model=get_data_from_model(sldmod.IModelDoc2(swx.OpenDoc6(pathname+"\\"+fname,0x1,0,'',None,None)))
        self.sketchName = sketchName
    def startEdit(self):
        self.Model.editSketch(self.sketchName)
        self.Sketch = self.Model._modeldoc.GetActiveSketch2
    def endEdit(self):
        self.Model.finishEditingSketch()
    def getEntities(self):
        #gets line, first point and all other points of sketch
        self.startEdit()
        SEGMENTS = slef.Sketch.GetSketchSegments
    #def setPositions(self, distances):
        #distances is list of dists between points. First point is always line point. Second point is always line point shared with sketch point. rest is 

if __name__=="__main__":
	P=project()
	P.get_excel_file()
	P.get_all_files()
	print P._pliki
	P.get_tops()
	P.find_parts('miecz')
	P.find_parts('miecz2')
	print find_parts()
