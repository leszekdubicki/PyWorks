# -*- coding: cp1250 -*-
#Version - 1.38 (2012-08-10)
#dodana komenda 'mat', zmieniona 'full'
#dodana komenda cc - wrzucanie linka do rysunku do schowka
#rozszerzone poszukiwanie listy (dodana funkcja get_list_path)
#dodalem opcje "ccp" - dwie linijki z linkiem do rysunku i pdf-a
#dodany zapis do jpg aktualnego widoku po wcisnieciu samej komendy v
#dodana obsloga bledow pryz komendyie full pryz yapiszwaniu do listz do wzcenz
#obsloga nowych list do wyc... (dodana funkcja list_version())
#1.35 - dodana funkcja tl, balinit c i c star w obr. chem
#1.36 - po wydruku (opcja p) rysunek zamyka sie
#1.37 - zamykanie po wydruku zlikwidowane - powodowalo czeste zawieszanie - SW, dodanie opcji cl - zamykanie rys.
#1.38 - dodana opcja sf (podmianka sheet format), problemy z listami do wyceny usuniete.
import prochecker
from prochecker import *
import os
from copy import deepcopy
import time as timemodule
import ConfigParser
import win32clipboard, win32com

#def send_mail_via_com(text, subject, recipient, profilename="Outlook2003",attachments = []):
def send_mail_via_com(text, subject, recipient = '', attachments = []):
    #s = win32com.client.Dispatch("Mapi.Session")
    o = win32com.client.Dispatch("Outlook.Application")
    #s.Logon(profilename)
    
    Msg = o.CreateItem(0)
    Msg.To = recipient
    
    #Msg.CC = "moreaddresses here"
    #Msg.BCC = "address"
    #attachment1 = "h:\\Pomiar czasu_ME_LDU.xlsx"
    #attachment2 = "Path to attachment no. 2"
    if isinstance(attachments, str):
        Msg.Attachments.Add(attachments)
    else:
        for A in attachments:
            Msg.Attachments.Add(A)
    
    Msg.Subject = subject
    Msg.Body = text
    
		
    #Msg.Attachments.Add(attachment2)
 
    #Msg.Send()
    Msg.Display()


def revision_change(DRW, REV_NO = None, REV_CI = "Rev.:"):
		#revision - zapisanie kopii pdf z numerem rev
		#REV_CI - custom info do zmiany w rysunku (jesli ni podano to "Rev.:")
		#REV_NO - 
		#sprawdzenie str. katalogów:
		#DRW.archive()
		DRW.create_folders()
		#pobranie numeru rev z modelu:
		DRW.get_parts()
		PRT = DRW._parts[0]
		if REV_NO==None:
			PRT.open(); PRT.activate()
			REV_NO = str(PRT.get_ci(REV_CI)).strip()
			if len(REV_NO)==0:
				REV_NO = str(PRT.get_ci('rev')).strip()
			#print REV_NO; raw_input()
			if (not isinstance(REV_NO,str)) or len(REV_NO)==0:
				NOTE3 = str(PRT.get_ci('note3')).strip()
				if len(NOTE3)==2 and NOTE3.isdigit():
					REV_NO = NOTE3
					PRT.set_ci('rev',NOTE3)
					PRT.set_ci('REV_CI',NOTE3)
				else:
					PRT.set_ci('rev','01')
					PRT.set_ci(REV_CI,'01')
					PRT.set_ci('Note3','01')
					PRT.save()
					REV_NO='01'
			PRT.close()
		elif not isinstance(REV_NO,str):
			REV_NO = str(REV_NO)
		#ustalenie nazwy pliku do zapisania:
		DRW.open()
		DRW.activate()
		NAME_TO_SAVE=DRW['path']+'archive\\'+DRW['name']+'rev'+REV_NO+'.slddrw'
		NAME_TO_SAVE_PDF=DRW['path']+'archive\\'+DRW['name']+'rev'+REV_NO+'.pdf'
		NAME_TO_SAVE_DWG=DRW['path']+'archive\\'+DRW['name']+'rev'+REV_NO+'.dwg'
		NAME_TO_OPEN=DRW['full_path']
		ERR = None; WARN = None
		DRW._modeldoc.SaveAs4(NAME_TO_SAVE,0x0,0x8,ERR,WARN)
		DRW._modeldoc.SaveAs2(NAME_TO_SAVE_PDF,0,True,False)
		DRW._modeldoc.SaveAs2(NAME_TO_SAVE_DWG,0,True,False)
		DRW = prochecker.get_data_from_drawing()
		DRW.open(); DRW.activate()
		DRW.close()
		os.remove(NAME_TO_SAVE)

		sldmod.IModelDoc2(swx.OpenDoc6(NAME_TO_OPEN,0x3,0,None,None, None))
		DRW2 = prochecker.get_data_from_drawing()
		DRW2.open(); DRW2.activate()
		#inkrementacja REV
		DRW2.get_parts()
		PRT = DRW2._parts[0]
		PRT.open(); PRT.activate()
		#REV_NO = PRT.get_ci(REV_CI)
		if REV_NO in [None, "", " "]:
			REV_NO = PRT.get_ci('Rev')
		
		if REV_NO.isdigit():
			REV_NO_2 = int(REV_NO)+1
			REV_NO_2 = str(REV_NO_2)
		else:
			REV_NO_2 = int(REV_NO[:-1])+1
			REV_NO_2 = str(REV_NO_2)
			
		if len(REV_NO_2)==1:
			REV_NO_2 = "0"+REV_NO_2
		PRT.set_ci(REV_CI,REV_NO_2)
		PRT.set_ci('rev',REV_NO_2)
		PRT.set_ci('note3',REV_NO_2)
		DATE = time.strftime("%Y-%m-%d")
		PRT.set_ci('LastSaveDate',DATE)
		PRT.save()
		#PRT.close()
		DRW2.open(); DRW2.activate()

def get_list_path(DRW, CONF):
	if CONF.has_option('machine','list'):
		LIST_FILENAME = CONF.get('machine','list')
		PATH = DRW['path'][:-1]
		#PATH=PRT['path']
		if os.path.isabs(LIST_FILENAME):
			return LIST_FILENAME
		if os.path.basename(PATH).lower()=="standards":
			PATH = os.path.dirname(PATH)
		if os.path.isdir(PATH+'\\parts list'):
			LIST_PATH = PATH+'\\parts list'+'\\'+LIST_FILENAME
		elif os.path.isdir(PATH+'\\part list'):
			LIST_PATH = PATH+'\\part list'+'\\'+LIST_FILENAME
		else:
			#print("BRAK LISTY!!!!!!!!!!!")
			LIST_PATH = None
	else:
		 LIST_PATH = None
	return LIST_PATH

def list_version(WORKSHEET):
	#zwraca numer wersji listy (0 dla pierwotnego, 1 dla utworzonego przez PAZ na poczatku 2012r)
	if (isinstance(WORKSHEET.Cells(4,"Q").Value,unicode)) and (WORKSHEET.Cells(4,"Q").Value.lower().strip() == "author:"):
		return 1
	elif (isinstance(WORKSHEET.Cells(4,"K").Value,unicode)) and (WORKSHEET.Cells(4,"K").Value.lower().strip() == "author:"):
		return 0
def columns(LIST_VER):
	C = {}
	if LIST_VER == 1:
		C['chem'] = "P"
		C['heat'] = "Q"
	elif LIST_VER == 0:
		C['chem'] = "J"
		C['heat'] = "K"
	return C

CHECK=1
def proceed_command(k, prompt = 1, argument1 = None, argument2 = None):
	if k=='help':
		#pomoc:
		print("\n\ncommands:")
		print("        c - prompt for drawing necessary information (may be joined with \"h\" and \"s\" command)\n")
		print("        h - prompt for heat treatment (may be joined with \"c\" and \"s\" command)\n")
		print("        s - prompt for chemical treatment (may be joined with \"c\" and \"h\" command)\n")
		print("        r - archive the drawing in \"archive\" subfolder and create new revision\n")
		print("        cancel - archive the drawing in \"archive\" subfolder with \"obs\" suffix (means drawing is obsolete) and delete it\n")
		print("        e - export drawing to pdf and dxf format in \"PDF_DXF_DWG\" subfolder\n")
		print("        f - change referenced part filename to drawing filename (run with assembly open)\n")
		print("        d - prompt for \"LastSaveDate\" custom info\n")
		print("        rr - prompt for \"rev\" (revision) custom info\n")
		print("        n1 - prompt for \"note1\" (project info) custom info\n")
		print("        a - prompt for standard part information (if changing catalogue number run with assembly open)\n")
		print("        al - like \"a\", but adding entry in parts list (using parts list file name from \".\\other\\project.txt\", under development)\n")
		print("        l - add entry in parts list (using parts list file name from \".\\other\\project.txt\", under development)\n")
		print("        dir - opens explorer in directory of the drawing")

	elif k=='e':
		#export do pdf-a i dxf-a
		DRW=get_data_from_drawing()
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		DRW.create_folders()
		DRW.open(); DRW.activate()
		NAME_PDF=DRW['path']+'PDF_DXF_DWG\\'+DRW['name']+'.pdf'
		NAME_DXF=DRW['path']+'PDF_DXF_DWG\\'+DRW['name']+'.dxf'
		DRW._modeldoc.SaveAs2(NAME_PDF,0,True,False)
		DRW._modeldoc.SaveAs2(NAME_DXF,0,True,False)
		win32clipboard.OpenClipboard()
		win32clipboard.EmptyClipboard()
		link = '<file://'+NAME_PDF+'>'
		win32clipboard.SetClipboardText(link, win32clipboard.CF_UNICODETEXT)
		win32clipboard.CloseClipboard()
		#DRW.get_parts()
		#PART = DRW._parts[0]
		#PART.open(); PART.activate()
		return link
	elif k=='pdf':
		#export do pdf-a
		if not argument1 == None:
			prochecker.swx.OpenDoc6(argument1,0x3,2,None,None,None) #otwarcie RO! - argument nr 3 wynosi 2
		DRW=get_data_from_drawing()
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		DRW.create_folders()
		DRW.open(); DRW.activate()
		NAME_PDF=DRW['path']+DRW['name']+'.pdf'
		DRW._modeldoc.SaveAs2(NAME_PDF,0,True,False)
		DRW.close()
		
	elif k=="f":
		#zmiana nazwy plików modelu (pierwszego) na numer
		DRW=get_data_from_drawing()
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		#DRW.get_parts()
		#DRW.set_drw_no_s()
		#STATUS=DRW.check_nums()
		NUM = DRW['name']
		DRW.get_parts()
		PRT = DRW._parts[0]
		if not PRT['name'] == DRW['name']:
			PRT.open(); PRT.activate()
			FULL_ERAASE_NAME = PRT['full_path']
			ERASE_NAME = PRT['name']
			ERR = None; WARN = None
			if is_sw_file(FULL_ERAASE_NAME)==1:
				EXT = 'sldprt'
			elif is_sw_file(FULL_ERAASE_NAME)==2:
				EXT = 'sldasm'
			#print PRT['path']; raw_input()
			PRT._modeldoc.SaveAs4(DRW['path']+DRW['name']+'.'+EXT,0x0,0x8,ERR,WARN)
			#print ERR, WARN, FULL_ERAASE_NAME
			#raw_input()
			PRT = get_data_from_model()
			PRT.open(); PRT.activate()
			PRT.set_ci('PartNo',DRW['name'])
			PRT.set_ci('LastDocumentName',ERASE_NAME)
			PRT.save()
			PRT.close()
			os.remove(FULL_ERAASE_NAME)
			DRW.open(); DRW.activate()
	elif k == "tl":
		#dopisanie do pliku textowego list-<data>.txt nazwy czesci
		import time
		DRW=get_data_from_model()
		NazwaPliku = DRW['path']+"list-"+time.strftime("%Y.%m.%d")+".txt"
		PLIK = open(NazwaPliku,"a")
		PLIK.write(DRW['plik']+"\r\n")
		PLIK.close()
	elif k == "d":
		DRW=get_data_from_drawing()
		#ustawia pole LastSaveDate na aktualna date
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		DRW.get_parts()
		PRT = DRW._parts[0]
		PRT.open(); PRT.activate()
		PRT = get_data_from_model()
		PRT.open(); PRT.activate()
		DATE = time.strftime("%Y.%m.%d")
		PRT.set_ci('LastSaveDate',DATE)
		print "ustawiono date na ", time.strftime("%Y.%m.%d")
		PRT.close()
		DRW.open(); DRW.activate()
	elif k=='r':
		DRW=get_data_from_drawing()
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		revision_change(DRW)
	elif k=="rev0":
		DRW=get_data_from_drawing()
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		revision_change(DRW,"00a")
	elif k=='cancel':
		#kasacja - zapisanie kopii pdf z numerem revXX-obs i usunięcie pliku modelu i rysunku z katalogu
		PROMPT=raw_input("\nUWAGA!!!\nskasujesz pliki rysunku i modelu, aby kontynuowac wpisz tak lub yes!!>>> ")
		if PROMPT=="yes" or PROMPT=="tak":
			DRW=get_data_from_drawing()
			if DRW==None:
				print ">>>>>>>>> to nie jest rysunek"
				#continue
				return(3)
			#sprawdzenie str. katalogów:
			DRW.create_folders()
			#pobranie numeru rev z modelu:
			DRW.get_parts()
			if len(DRW._parts)==0:
				print ">>>>>>>>> brak rozpoznawalnego widoku rysunku na arkuszu"
				#continue
				return(3)
			PRT = DRW._parts[0]
			PRT.open(); PRT.activate()
			REV_NO = str(PRT.get_ci('Rev.:')).strip()
			if len(REV_NO)==0:
				REV_NO = str(PRT.get_ci('rev')).strip()
			#print REV_NO; raw_input()
			if (not isinstance(REV_NO,str)) or len(REV_NO)==0:
				NOTE3 = str(PRT.get_ci('note3')).strip()
				if len(NOTE3)==2 and NOTE3.isdigit():
					REV_NO = NOTE3
					PRT.set_ci('Rev.:',NOTE3)
					PRT.set_ci('rev',NOTE3)
				else:
					PRT.set_ci('Rev.:','01')
					PRT.set_ci('rev','01')
					PRT.set_ci('Note3','01')
					PRT.save()
					REV_NO='01'
			NAME_PRT=PRT['name']
			NAME_TO_DELETE_PRT=PRT['full_path']
			timemodule.sleep(1) #dodane żeby nie powodować zawieszania SW
			PRT.close()
			#ustalenie nazwy pliku do zapisania:
			DRW.open()
			DRW.activate()
			NAME_TO_SAVE=DRW['path']+'archive\\'+DRW['name']+'rev'+REV_NO+'-obs.slddrw'
			NAME_TO_SAVE_PDF=DRW['path']+'archive\\'+DRW['name']+'rev'+REV_NO+'-obs.pdf'
			NAME_TO_SAVE_DWG=DRW['path']+'archive\\'+DRW['name']+'rev'+REV_NO+'-obs.dwg'
			NAME_TO_DELETE_DRW=DRW['full_path']
			if not NAME_PRT==DRW['name']:
				print "nazwa modelu inna niz nazwa rysunku, nie kasuje modelu!"
				NAME_TO_DELETE_PRT=''
			ERR = None; WARN = None
			DRW._modeldoc.SaveAs4(NAME_TO_SAVE,0x0,0x8,ERR,WARN)
			DRW._modeldoc.SaveAs2(NAME_TO_SAVE_PDF,0,True,False)
			DRW._modeldoc.SaveAs2(NAME_TO_SAVE_DWG,0,True,False)
			DRW = prochecker.get_data_from_drawing()
			DRW.open(); DRW.activate()
			timemodule.sleep(1) #dodane żeby nie powodować zawieszania SW
			DRW.close()
			os.remove(NAME_TO_SAVE)
			os.remove(NAME_TO_DELETE_DRW)
			if len(NAME_TO_DELETE_PRT.strip())>0:
				os.remove(NAME_TO_DELETE_PRT)
	elif k=="rr":
		#argument1 - numer revizji do ustawienia
		#argument2 - nazwa właściwości do ustawienia, domyślnie Rev.:, rev i Note3
		DRW=get_data_from_drawing()
		#sztucznie ustawia pole rev
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		#pobranie numeru rev z modelu:
		DRW.get_parts()
		PRT = DRW._parts[0]
		PRT.open(); PRT.activate()
		if argument2 == None:
			REV_NO = str(PRT.get_ci('Rev.:')).strip()
			argument2 = "rev"
			#argument2 = 'Rev.:'
		else: 
			REV_NO = str(PRT.get_ci(argument2)).strip()
		if len(REV_NO)==0:
			REV_NO = str(PRT.get_ci('rev')).strip()
		print "Obecny revision number: "+REV_NO
		if argument1 == None:
			PRT.check_mand_infos(1,[argument2])
		else:
			if argument1 == None:
				argument1 = "01"
			if isinstance(argument1, int):
				argument1 = str(argument1)
				argument1 = (2-len(argument1))*"0" + argument1
			PRT.set_ci(argument2, argument1)
			PRT.set_ci("Rev.:", argument1)
			PRT.set_ci("rev", argument1)
			PRT.set_ci("rev.", argument1)
			PRT.set_ci("Note3", argument1)
		PRT.save()
		PRT.close()
		DRW.open(); DRW.activate()
	elif k=="n1":
		DRW=get_data_from_drawing()
		#pyta i ustawia pole note1
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		#pobranie note1 z modelu:
		DRW.get_parts()
		PRT = DRW._parts[0]
		PRT.open(); PRT.activate()
		NOTE1 = str(PRT.get_ci('note1')).strip()
		print "Obecny note1: "+NOTE1
		PRT.check_mand_infos(1,['note1'])
		PRT.save()
		PRT.close()
		DRW.open(); DRW.activate()
	elif k=="n2":
		DRW=get_data_from_drawing()
		#pyta i ustawia pole note2
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		#pobranie note2 z modelu:
		DRW.get_parts()
		PRT = DRW._parts[0]
		PRT.open(); PRT.activate()
		NOTE2 = str(PRT.get_ci('note2')).strip()
		print "Obecny note2: "+NOTE2
		PRT.check_mand_infos(1,['note2'])
		PRT.save()
		PRT.close()
		DRW.open(); DRW.activate()
	elif k=="nn":
		DRW=get_data_from_drawing()
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		#pobranie note 1 i 2 z modelu:
		DRW.get_parts()
		PRT = DRW._parts[0]
		PRT.open(); PRT.activate()
		NOTE1 = str(PRT.get_ci('note1')).strip()
		NOTE2 = str(PRT.get_ci('note2')).strip()
		print "Obecny note1: "+NOTE1
		print "Obecny note2: "+NOTE2
		PRT.check_mand_infos(1,['note1','note2'])
		PRT.save()
		PRT.close()
		DRW.open(); DRW.activate()
	elif k=="dir":
		import win32api
		DRW=get_data_from_model()
		win32api.WinExec('explorer '+DRW['path'].replace('/','\\'))
	elif k=="al" or k=="a":
		#ustawia nazwe pliku cz. standardowej i jej opis, kasuje stary plik
		PRT=get_data_from_model()
		PRT.open(); PRT.activate()
		if os.path.isfile(PRT['path']+PRT['name']+'.slddrw'):
			print "w katalogu znajduje sie rysunek tej czesci, przerywam!!"
		elif is_sw_file(PRT._dane['plik'])==3:
			print "To jest rysunek, przerywam!!"
		else:
			NAME=raw_input("Nazwa (numer kat.)>> ")
			ERASE=1
			if len(NAME)==0 or NAME==PRT['name']:
				ERASE=0
			DESC=raw_input("Opis (description)>> ")
			PROD=raw_input("Produc. (producer)>> ")
			QTY=raw_input("Ilosc (quantity)>> ")
			#MAT=raw_input("Material (Material)>> ")
			MAT="N.A."
			FULL_ERAASE_NAME = PRT['full_path']
			ERASE_NAME = PRT['name']
			OLD_PART_NO=PRT.get_ci('PartNo')
			ERR = None; WARN = None
			if is_sw_file(FULL_ERAASE_NAME)==1:
				EXT = 'sldprt'
			elif is_sw_file(FULL_ERAASE_NAME)==2:
				EXT = 'sldasm'
			if ERASE:
				PRT._modeldoc.SaveAs4(PRT['path']+NAME+'.'+EXT,0x0,0x8,ERR,WARN)
				PRT = get_data_from_model()
				PRT.open(); PRT.activate()
				PRT.set_ci('LastDocumentName',ERASE_NAME)
				PRT.set_global_ci('LastDocumentName',ERASE_NAME)
				os.remove(FULL_ERAASE_NAME)
			if (not NAME==OLD_PART_NO) and (len(NAME)>0):
				print "zmiana PartNo"
				PRT.set_ci('PartNo',NAME)
				PRT.set_global_ci('PartNo',NAME)
			if (not DESC==PRT.get_ci('Description')) and (len(DESC)>0):
				print "zmiana Description"
				PRT.set_ci('Description',DESC)
				PRT.set_global_ci('Description',DESC)
			if (not DESC==PRT.get_ci('Title:')) and (len(DESC)>0):
				print "zmiana Title:"
				PRT.set_ci('Title:',DESC)
				PRT.set_global_ci('Title:',DESC)
			if (not PROD==PRT.get_ci('Producer')) and (len(PROD)>0):
				print "zmiana Producer"
				PRT.set_ci('Producer',PROD)
				PRT.set_ci('Notes',PROD)
				PRT.set_global_ci('Producer',PROD)
				PRT.set_global_ci('Notes',PROD)
			if (PRT.get_ci("Notes")=="" or PRT.get_ci("Notes")==None) and (len(PROD)>0):
				PRT.set_ci('Notes',PROD)
			if (not QTY==PRT.get_ci('quantity')) and (len(QTY)>0):
				print "zmiana Quantity"
				PRT.set_ci('quantity',QTY)
				PRT.set_global_ci('quantity',QTY)
			PRT.set_ci('material',MAT)
			PRT.set_global_ci('material',MAT)
			PRT.save()
			#_______________________________________________________________________________
			#sprawdzenie, czy w poleceniu jest l, jesli tak to dopisanie do listy do wyceny:
			if "l" in k:
				FILENAME='project.txt'
				PATH=PRT['path'][:-1]
				#print PATH
				#print("basename: "+os.path.basename(PATH).lower())
				#raw_input()
				if os.path.basename(PATH).lower()=="standards":
					PATH = os.path.dirname(PATH)
				CONF_PATH = PATH+'\\other'+'\\'+FILENAME
				CONF=ConfigParser.ConfigParser()
				if os.path.isfile(CONF_PATH):
					CONF.read(CONF_PATH)
				if CONF.has_option('machine','list'):
					LIST_FILENAME = CONF.get('machine','list')
					"""
					#PATH=PRT['path']
					if os.path.isdir(PATH+'\\parts list'):
						LIST_PATH = PATH+'\\parts list'+'\\'+LIST_FILENAME
					elif os.path.isdir(PATH+'\\part list'):
						LIST_PATH = PATH+'\\part list'+'\\'+LIST_FILENAME
					else:
						print("BRAK LISTY!!!!!!!!!!!")
						exit
					"""
					LIST_PATH = get_list_path(PRT, CONF)
					if not LIST_PATH:
						print("BRAK LISTY!!!!!!!!!!!")
						exit
						 
					N_A="N.A."
					if os.path.isfile(CONF_PATH):
						print LIST_PATH
						xl = win32com.client.Dispatch("Excel.Application")
						xl.Visible = True
						wb = xl.Workbooks.Open(LIST_PATH)
						ws = wb.Worksheets('Parts list')
						#mamy rysunek lub czesc specjalna
						COL=2
						ROW=8
						#poszukanie poczatku sekcji standardowych komponentow
						FOUND_STANDARD = 0; MAX_ROW = 1000
						while not FOUND_STANDARD:
							if ws.Cells(ROW,2).Value=="Standard components - mechanics":
								FOUND_STANDARD = 1
								#print "znalazlem sekcje z komponentami standardowymi, jest w wierszu "+str(ROW)
								#raw_input()
							elif ROW>1000:
								print "///////////////////////////////////////////////////////////////"
								print "///////////////////////////////////////////////////////////////"
								print "brak miejsca, dodaj wiersze"
								print "///////////////////////////////////////////////////////////////"
								print "///////////////////////////////////////////////////////////////"
							ROW+=1
							
						NO_ROWS=0
						CATNO=PRT['name']
						while not NO_ROWS:
							if str(ws.Cells(ROW,2).Value).strip()=="Standard components - electronics":
								print "///////////////////////////////////////////////////////////////"
								print "///////////////////////////////////////////////////////////////"
								print "brak miejsca, dodaj wiersze"
								print "///////////////////////////////////////////////////////////////"
								print "///////////////////////////////////////////////////////////////"
								break
							elif (unicode(ws.Cells(ROW,2).Value).strip()=="" or unicode(ws.Cells(ROW,2).Value).strip()==u"N.A.") and (unicode(ws.Cells(ROW,3).Value).strip()=="" or unicode(ws.Cells(ROW,3).Value).strip()=="N.A."):
								#mamy wolna linijke:
								#wyciagniecie danych z pliku:
								print "wiersz " + str(ROW) + " jest wolny"
								PRODUCER=PRT.get_ci('producer').strip()
								if len(PRODUCER)==0:
									PRODUCER=N_A
								DESC=PRT.get_ci('description')
								QUANT=PRT.get_ci('quantity')
								if len(QUANT)==0:
									QUANT=""
								#zapisanie
								ws.Cells(ROW,6).Value=CATNO
								ws.Cells(ROW,3).Value=DESC
								ws.Cells(ROW,4).Value=QUANT
								ws.Cells(ROW,7).Value=PRODUCER
								print "Dodalem do listy do wyceny"
								wb.Save()
								break
							elif isinstance(ws.Cells(ROW,6).Value,int) and str(int(ws.Cells(ROW,6).Value))==CATNO:
								#wiersz juz istnieje, trzeba tylko zaktualizowac dane(zakladam, ze w liscie nie ma wolnych pol!!):
								print "component juz wpisany w liste"
								DESC=PRT.get_ci('description')
								QUANT=PRT.get_ci('quantity')
								if len(QUANT)==0:
									QUANT=""
								else:
									ws.Cells(ROW,4).Value=QUANT
								#zapisanie
								ws.Cells(ROW,6).Value=CATNO
								ws.Cells(ROW,3).Value=DESC
								print "Uaktualnilem liste do wyceny"
								wb.Save()
								break
							#print ws.Cells(ROW,6).Value
							#print ROW
							#raw_input()
							ROW+=1
					else:
						print ">>>>>>>>> Brak pliku listy"
						#continue
						return(6)
				else:
					print ">>>>>>>>> Brak informacji o liscie w pliku project.txt"
					#continue
					return(6)
				
	elif k=='i':
		#zapisanie pliku informacyjnego o projekcie z którego będzie czytana informacja w pliku project.conf w katalogu "katalog części\other"
		DRW=get_data_from_model()
		DRW.set_project_config(DATA)
	elif k=="l":
		#wpisanie rysunku do listy do wyceny (info o tym, jaki plik excela otworzyc musi byc w pliku project.txt)
		STANDARD = 0
		DRW=get_data_from_drawing()
		if DRW==None:
			DRW=get_data_from_model()
			STANDARD=1
		#DRW.archive()
		DRW.create_folders()
		FILENAME='project.txt'
		PATH=DRW['path'][:-1]
		CONF_PATH = PATH+'\\other'+'\\'+FILENAME
		CONF=ConfigParser.ConfigParser()
		if os.path.isfile(CONF_PATH):
			CONF.read(CONF_PATH)
		if CONF.has_option('machine','list'):
			LIST_FILENAME = CONF.get('machine','list')
			PATH=DRW['path']
			#LIST_PATH = PATH+'\\parts list'+'\\'+LIST_FILENAME
			LIST_PATH = get_list_path(DRW, CONF)
			N_A="N.A."
			if os.path.isfile(CONF_PATH):
				#print LIST_PATH
				xl = win32com.client.Dispatch("Excel.Application")
				xl.Visible = True
				wb = xl.Workbooks.Open(LIST_PATH)
				ws = wb.Worksheets('Parts list')
				LIST_VER = list_version(ws)
				SPEC_COLS = columns(LIST_VER)
				if STANDARD==0:
					#mamy rysunek lub czesc specjalna
					COL=2
					ROW=8
					NO_ROWS=0
					PARTNO=DRW['name']
					while not NO_ROWS:
						if ws.Cells(ROW,2).Value=="Standard components - mechanics":
							print "///////////////////////////////////////////////////////////////"
							print "///////////////////////////////////////////////////////////////"
							print "brak miejsca, dodaj wiersze"
							print "///////////////////////////////////////////////////////////////"
							print "///////////////////////////////////////////////////////////////"
							break
						elif (str(ws.Cells(ROW,2).Value).strip()=="" or str(ws.Cells(ROW,2).Value).strip()=="N.A.") and (str(ws.Cells(ROW,3).Value).strip()=="" or str(ws.Cells(ROW,3).Value).strip()=="N.A."):
							#mamy wolna linijke:
							#wyciagniecie danych z pliku:
							print "wiersz " + str(ROW) + " jest wolny"
							DRW.get_parts()
							PRT=DRW._parts[0]
							PRT.open(); PRT.activate()
							HEAT=PRT.get_ci('heat').strip()
							if len(HEAT)==0:
								HEAT=N_A
							elif "<MOD-PM>" in HEAT:
								HEAT=HEAT.replace('<MOD-PM>','+/-')
							CHEM=PRT.get_ci('chemical')
							if len(CHEM)==0:
								CHEM=N_A
							DESC=PRT.get_ci('description')
							QUANT=PRT.get_ci('quantity')
							MAT=PRT.material()
							#zapisanie
							ws.Cells(ROW,2).Value=PARTNO
							ws.Cells(ROW,3).Value=DESC
							ws.Cells(ROW,4).Value=QUANT
							ws.Cells(ROW,8).Value=MAT
							ws.Cells(ROW,SPEC_COLS['chem']).Value=CHEM
							ws.Cells(ROW,SPEC_COLS['heat']).Value=HEAT
							print "Dodalem do listy do wyceny"
							PRT.close()
							DRW.open(); DRW.activate()
							wb.Save()
							break
						elif str(int(ws.Cells(ROW,2).Value))==PARTNO:
							#wiersz juz istnieje, trzeba tylko zaktualizowac dane(zakladam, ze w liscie nie ma wolnych pol!!):
							print "component juz wpisany w liste"
							DRW.get_parts()
							PRT=DRW._parts[0]
							PRT.open(); PRT.activate()
							HEAT=PRT.get_ci('heat').strip()
							if len(HEAT)==0:
								HEAT=N_A
							elif "<MOD-PM>" in HEAT:
								HEAT=HEAT.replace('<MOD-PM>','+/-')
							CHEM=PRT.get_ci('chemical')
							if len(CHEM)==0:
								CHEM=N_A
							DESC=PRT.get_ci('description')
							QUANT=PRT.get_ci('quantity')
							MAT=PRT.material()
							#zapisanie
							ws.Cells(ROW,2).Value=PARTNO
							ws.Cells(ROW,3).Value=DESC
							ws.Cells(ROW,4).Value=QUANT
							ws.Cells(ROW,8).Value=MAT
							ws.Cells(ROW,SPEC_COLS['chem']).Value=CHEM
							ws.Cells(ROW,SPEC_COLS['heat']).Value=HEAT
							print "Uaktualnilem liste do wyceny"
							PRT.close()
							DRW.open(); DRW.activate()
							wb.Save()
							break
						ROW+=1
			else:
				print ">>>>>>>>> Brak pliku listy"
				#continue
				return(6)
		else:
			print ">>>>>>>>> Brak informacji o liscie w pliku project.txt"
			#continue
			return(6)
				
	elif k=="p":
		#drukowanie
		CONFIG_FILE = None
		try:
			CONFIG_PATH = sys.argv[0][0:sys.argv[0].rindex("\\")]+"\\"
		except:
			CONFIG_PATH = ".\\"
			
		CONFIG = ConfigParser.ConfigParser()
		CONFIG_FILE = CONFIG_PATH+"config-drwnum.txt"
		if os.path.isfile(CONFIG_FILE):
			CONFIG.read(CONFIG_FILE)
		else:
			print "brak pliku konfiguracyjnego"
			CONFIG.set("general", "printer", "\\\\PL01W03W04\\Szczecin - Printer 67 : LaserJet 5200")
		DRW=get_data_from_drawing()
		DRW.open(); DRW.activate()
		DRW.Print(CONFIG)
		#DRW.close()
	elif k=="mat":
		#automatyczne dodanie danych materialu (wczesniej prosba o jego zdefiniowanie )
		DRW=get_data_from_drawing()
		if is_sw_file(DRW['full_path'])==3:
			DRW.get_parts()
			PRT = DRW._parts[0]
		#else:
		#	PRT=get_data_from_model()
		PRT.open(); PRT.activate()
		if is_sw_file(PRT['full_path'])==1:
			MAT = PRT.material()
			print "MAT: " + MAT
			#raw_input()
			if MAT in ["Material <not specified>", "", "N.A."]:
				MAT = None
			while not MAT:
				print("zdefiniuj material i nacisnij enterinio :D")
				raw_input()
				MAT = PRT.material()
				if MAT in ["Material <not specified>", "", "N.A."]:
					MAT = None
			#ustawienie hartowania:
			DRW.open(); DRW.activate()
			if "stavax" in MAT.lower():
				proceed_command("h", argument1 = "Vacuum Hardening 52<MOD-PM>2 HRC")
				proceed_command("s", argument1 = "N.A.")
			elif "aluminium" in MAT.lower():
				proceed_command("h", argument1 = "N.A.")
				proceed_command("s", argument1 = "Anodizing (natural color)")
			else:
				proceed_command("h", argument1 = "N.A.")
				proceed_command("s", argument1 = "N.A.")
		elif is_sw_file(PRT['full_path'])==2:
			PRT.set_ci('Material','N.A.')
			PRT.set_global_ci('Material','N.A.')
			DRW.open(); DRW.activate()
			proceed_command("h", argument1 = "N.A.")
			proceed_command("s", argument1 = "N.A.")
				
			
	elif k=="full":
		print("zmiana nazwy pliku czesci na numer rys:")
		proceed_command('f')
		print("zmiana rev na 01:")
		proceed_command('rr',argument1 = '01')
		print("nadanie wlasciwosci modelowi:")
		print("\n\t\tPamietaj o materiale!!!\n")
		proceed_command('c')
		proceed_command('mat')
		print("dodanie do listy do wyceny:")
		try:
			proceed_command('l')
		except:
			print("blad podczas otwierania excela")
		DRW = get_data_from_drawing()
		DRW.open(); DRW.activate()
		DRW.Rebuild()
		DRW.save()
	elif k=='mail':
		#generuje maila w outlooku z linkami do katalogow, rysunku zlozeniowego, listy itd
		DRW = get_data_from_drawing()
		DRW.open(); DRW.activate()
		#plik wlasny (zlorzenie):
		DRW.get_parts()
		
		plik = '<file://'+DRW._parts[0]['full_path']+'>'
		#link do pdf-a:
		pdf = proceed_command('e')
		#link do katalogu:
		katalog = '<file://'+DRW['path']+'>'
		#link do katalogu pdf-ow:
		pdf_y =  '<file://'+DRW['path']+"PDF_DXF_DWG"+'>'
		#link do listy:
		FILENAME='project.txt'
		PATH=DRW['path'][:-1]
		CONF_PATH = PATH+'\\other'+'\\'+FILENAME
		CONF=ConfigParser.ConfigParser()
		CONF.read(CONF_PATH)
		LIST_FILENAME = CONF.get('machine','list')
		machinename = CONF.get('machine','name')
		LIST_PATH = PATH+'\\parts list'+'\\'+LIST_FILENAME
		if os.path.isfile(LIST_PATH):
			lista = '<file://'+LIST_PATH+'>'
		else:
			 lista = "brak!"
		#tworzenie maila:
		subject = "rysunki: "+machinename
		mailtext = u"Hi\n\nRysunki do urządzenia są gotowe\n\n\n"
		mailtext+= "katalog:\t"+katalog + '\n\n'
		mailtext+= "model 3D:\t"+plik + '\n\n'
		mailtext+= "pdf:\t"+pdf + '\n\n'
		mailtext+= "wszystkie pdf-y:\t"+pdf_y + '\n\n'
		mailtext+= "lista:\t"+lista + '\n\n'
		mailtext+="Pozdrawiam\nLDU\n"
		send_mail_via_com(mailtext, subject)
		
	elif k=="cc":
		#wrzucenie linka do rysunku do schowka
		PRT = get_data_from_model()
		FULL_PATH = PRT['full_path']
		win32clipboard.OpenClipboard()
		win32clipboard.EmptyClipboard()
		win32clipboard.SetClipboardText('<file://'+FULL_PATH+'>', win32clipboard.CF_UNICODETEXT)
		win32clipboard.CloseClipboard()
	elif k=="ccp":
		#wrzucenie linka do rysunku do schowka
		PRT = get_data_from_model()
		pdflink = "\npdf: " + proceed_command("e")
		FULL_PATH = PRT['full_path']
		win32clipboard.OpenClipboard()
		win32clipboard.EmptyClipboard()
		win32clipboard.SetClipboardText('<file://'+FULL_PATH+'>' + pdflink, win32clipboard.CF_UNICODETEXT)
		win32clipboard.CloseClipboard()
	elif k=="t":
		#ustawia Description tak samo jak "Title:"
		PRT = get_data_from_model()
		PRT.set_ci('Description',PRT.get_ci('Title:'))
	elif k[0]=="v":
		#operacje na widokach:
		VIEWS = ["General", "Front", "Back", "Top", "Left", "Right"]
		if len(k)==2 and k[1]=="a":
			#automatyczne tworzenie widokow ze standardowych
			VIEWS2 = ["*Isometric", "*Front", "*Back", "*Top", "*Left", "*Right"]
			PRT = get_data_from_model()
			PRT.open(); PRT.activate()
			for i in range(0,len(VIEWS)):
				PRT._modeldoc.ShowNamedView2(VIEWS2[i], -1)
				if not PRT._modeldoc.Extension.GetNamedViewRotation(VIEWS[i])==(0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0):
					PRT._modeldoc.DeleteNamedView(VIEWS[i])
				PRT._modeldoc.NameView(VIEWS[i])
				PRT._modeldoc.ViewZoomtofit2()
				PRT._modeldoc.SaveAs3(PRT['path']+"\\"+VIEWS[i]+".jpg", 0, 0)
			#PRT._modeldoc.ShowNamedView2("Front", -1)
			#print PRT._modeldoc.Extension.GetNamedViewRotation ("Front")
		elif len(k)==2 and k[1]=="t":
			#test
			PRT = get_data_from_model()
			PRT.open(); PRT.activate()
			C = PRT._modeldoc.GetConfigurationByName("view")
			print "mm-1" in C.GetDisplayStates
			PRT._modeldoc.ShowNamedView2("test", -1)
			print PRT._modeldoc.GetModelViewNames()
		else:
			#zapisanie aktualnego widoku do jpg z data i godz:
			PRT = get_data_from_model()
			PRT.open(); PRT.activate()
			import time
			DATESTR = time.strftime("_%Y-%m-%d_%H-%M-%S")
			NAME_TO_SAVE_JPG = PRT['path']+"\\screen-"+PRT['name']+DATESTR+'.jpg'
			PRT._modeldoc.SaveAs3(NAME_TO_SAVE_JPG, 0, 0)
	elif k=="change all":
		#wprowadza zmiany w calym katalogu, katalog pobiera z otwartego pliku
		PRT = prochecker.get_data_from_model()
		PATH = PRT['path']
		FILES = os.listdir(PATH)
		for F in FILES:
			if F[-7:].lower()==".slddrw":
				print("Zmiana rysunku nr "+F[:-7])
				prochecker.swx.OpenDoc6(PATH+F,0x3,0,None,None,None)
				#proceed_command("r")
				proceed_command("c",prompt = 0)
				PRT = prochecker.get_data_from_drawing()
				PRT.open(); PRT.activate()
				PRT.close()
	elif k=="change all rev":
		#wprowadza zmiany w calym katalogu, katalog pobiera z otwartego pliku
		PRT = prochecker.get_data_from_model()
		PATH = PRT['path']
		FILES = os.listdir(PATH)
		for F in FILES:
			if F[-7:].lower()==".slddrw":
				print("Zmiana rysunku nr "+F[:-7])
				prochecker.swx.OpenDoc6(PATH+F,0x3,0,None,None,None)
				proceed_command("r")
				proceed_command("c",prompt = 0)
				PRT = prochecker.get_data_from_drawing()
				PRT.open(); PRT.activate()
				PRT.close()
	elif k=="e all":
		#exportuje caly katalog
		PRT = prochecker.get_data_from_model()
		PATH = PRT['path']
		FILES = os.listdir(PATH)
		for F in FILES:
			if F[-7:].lower()==".slddrw" and not F[0] == "~":
				print("Export rysunku nr "+F[:-7])
				prochecker.swx.OpenDoc6(PATH+F,0x3,0,None,None,None)
				#proceed_command("r")
				proceed_command("e")
				PRT = prochecker.get_data_from_drawing()
				PRT.open(); PRT.activate()
				PRT.close()
	elif k=="p all":
		#exportuje caly katalog
		import time
		PRT = prochecker.get_data_from_model()
		PATH = PRT['path']
		FILES = os.listdir(PATH)
		for F in FILES:
			if F[-7:].lower()==".slddrw" and not F[0] == "~":
				print("Wydruk rysunku nr "+F[:-7])
				prochecker.swx.OpenDoc6(PATH+F,0x3,0,None,None,None)
				#proceed_command("r")
				proceed_command("p")
				PRT = prochecker.get_data_from_drawing()
				PRT.open(); PRT.activate()
				PRT.close()
				time.sleep(1) #odczekanie na wyczyszczenie bufora
	elif k == "cl":
		PRT = prochecker.get_data_from_model()
		PRT.close()
	elif k == "sf":
		#sf od SheetFormat
		CONFIG_FILE = None
		try:
			CONFIG_PATH = sys.argv[0][0:sys.argv[0].rindex("\\")]+"\\"
		except:
			CONFIG_PATH = ".\\"
			
		CONFIG = ConfigParser.ConfigParser()
		CONFIG_FILE = CONFIG_PATH+"config-drwnum.txt"
		if os.path.isfile(CONFIG_FILE):
			CONFIG.read(CONFIG_FILE)
			PRT = prochecker.get_data_from_drawing()
			PRT.open()
			PRT.activate()
			PRT.replace_sheet_format(CONFIG)
			PRT.set_units()
		else:
			print "brak pliku konfiguracyjnego, przerywam...."
	elif k == "sfs":
		#sfs od SheetFormats - podmienia wszystkie arkusze
		CONFIG_FILE = None
		try:
			CONFIG_PATH = sys.argv[0][0:sys.argv[0].rindex("\\")]+"\\"
		except:
			CONFIG_PATH = ".\\"
			
		CONFIG = ConfigParser.ConfigParser()
		CONFIG_FILE = CONFIG_PATH+"config-drwnum.txt"
		if os.path.isfile(CONFIG_FILE):
			CONFIG.read(CONFIG_FILE)
			PRT = prochecker.get_data_from_drawing()
			PRT.open()
			PRT.activate()
			PRT.replace_sheet_formats(CONFIG)
			#PRT.set_units()
		else:
			print "brak pliku konfiguracyjnego, przerywam...."
	else:
		DRW=get_data_from_drawing()
		if DRW==None:
			print ">>>>>>>>> to nie jest rysunek"
			#continue
			return(3)
		#DRW.get_parts()
		#DRW.set_drw_no_s()
		#STATUS=DRW.check_nums()
		NUM = DRW['name']
		DRW.get_parts()
		PRT = DRW._parts[0]
		#odczytanie pliku konfiguracyjnego:
		FILENAME='project.txt'
		PATH=PRT['path'][:-1]
		CONF_PATH = PATH+'\\other'+'\\'+FILENAME
		#print len(CONF_PATH)
		#print CONF_PATH
		#raw_input()
		CONF=ConfigParser.ConfigParser()
		if os.path.isfile(CONF_PATH):
			CONF.read(CONF_PATH)

		PRT.open(); PRT.activate()
		if 'c' in k:
			if not PRT.get_ci('PartNo')==DRW['name']:
				PRT.set_ci('PartNo',DRW['name'])
			#sprawdzenie pliku konfiguracyjnego i odczytanie danych:
			if os.path.isfile(CONF_PATH):
				CONF.read(CONF_PATH)
				#print CONF.get('machine','name')
				#raw_input('odczytalem')
				PRT.set_ci('note1',CONF.get('project','number')+' MED '+CONF.get('project','short_name')+' '+CONF.get('machine','name')+' (ID '+CONF.get('machine','OT')+')')
				PRT.set_ci('axapta',CONF.get('machine','axapta'))
				PRT.set_ci('OT',CONF.get('machine','OT'))
				PRT.set_ci('Equipment ID:',CONF.get('machine','OT')) #RAS tamplates
				PRT.set_ci('Project no - axapta:',CONF.get('machine','axapta'))#RAS tamplates
				PRT.set_ci('Project no:',CONF.get('project','number'))
				PRT.set_ci('Drawn:',"LDU")#RAS tamplates - tylko LDU, zmienic na username!
				PRT.set_ci('Equipment name:',CONF.get('machine','name'))
				PRT.set_ci('Production:',CONF.get('project','short_name'))#RAS tamplates, w project.txt trzeba pisac MED lub inne na początku!
				if CONF.has_section('properties'):
					#dodanie niestandardowych wlasnosci z pliku:
					CUSTOM_OPTS = CONF.options('properties')
					for O in CUSTOM_OPTS:
						PRT.set_ci(O,CONF.get('properties',O))
			PRT.save()
			PRT.close()
			DRW.open(); DRW.activate()
			if CHECK:
				print "\nsprawdzam inne wlasciwosci:"
				STATUS = DRW.check_ness_infos(prompt)
			#CONF.write(open(CONF_PATH,'w'))
				#DRW.check_mand_infos()
			
		if 'h' in k:
			KODY = ["Vacuum Hardening 52<MOD-PM>2 HRC", "Vacuum Hardening 62<MOD-PM>2 HRC"]
			KODY_NOSW = ["Vacuum Hardening 52+/-2 HRC", "Vacuum Hardening 62+/-2 HRC"]
			"""
			print "\nwybierz kod lub wpisz rodzaj obr. cieplnej:\n\n\tkody:\n\
				1 - \"Vacuum Hardening 52+/-2 HRC (Hartowac prozniowo)\" (stavax)\
				2 - \"Vacuum Hardening 62+/-2 HRC (Hartowac prozniowo)\"\
				"
				"""
			print "\nwybierz kod lub wpisz rodzajobr. cieplnej:\n\tkody:"
			print "\t1 - \"Vacuum Hardening 52+/-2 HRC (Hartowac prozniowo)\" (stavax)"
			print "\t2 - \"Vacuum Hardening 62+/-2 HRC (Hartowac prozniowo)\""

			if k=="h" and not argument1==None:
				text_wlasn = argument1
			else:
				text_wlasn=raw_input("\n\t>> ")
			text_wlasn_nosw=text_wlasn
			if len(text_wlasn)==1 and text_wlasn.isdigit():
				COD_NO = int(text_wlasn)-1
				text_wlasn=KODY[COD_NO]
				text_wlasn_nosw=KODY_NOSW[COD_NO]
			PRT.set_ci('heat',text_wlasn)
			PRT.set_ci('Heat Treatment:',text_wlasn)
			PRT.set_ci('notes',text_wlasn)
			PRT.save()
			#if CONF.has_section(PRT['name']):
			#	CONF.set(PRT['name'],'heat',text_wlasn_nosw)
			#	CONF.write(open(CONF_PATH,'w'))
		if 's' in k:
			KODY = ["Anodizing (natural color)", "PVD Fut. Nano Top", "PVD Balinit Crovega", "PVD Balinit C", "PVD Balinit C Star"]
			print "\nwybierz kod lub wpisz rodzaj obr. chemicznej:\n\tkody:\n"
			print "\t1 - \"Anodizing natural color (anodowac)"
			print "\t2 - \"PVD - BALINIT  FUTURA NANO TOP (pokrywac PVD)\""
			print "\t3 - \"PVD - BALINIT  Crovega\""
			print "\t4 - \"PVD - BALINIT  C\""
			print "\t5 - \"PVD - BALINIT  C Star\""
				

			if k=="s" and not argument1==None:
				text_wlasn = argument1
			else:
				text_wlasn=raw_input("\n\t>> ")
			text_wlasn_nosw=text_wlasn
			if len(text_wlasn)==1 and text_wlasn.isdigit():
				text_wlasn=KODY[int(text_wlasn)-1]
			PRT.set_ci('chemical',text_wlasn)	
			PRT.set_ci('surface',text_wlasn)
			PRT.set_ci('Chemical Treatment:',text_wlasn)
			PRT.set_ci('Surface Treatment:',text_wlasn)
			PRT.set_ci('notes',text_wlasn)
			PRT.save()
			#if CONF.has_section(PRT['name']):
			#	CONF.set(PRT['name'],'chemical',text_wlasn)
			#	CONF.write(open(CONF_PATH,'w'))
		DRW.close_parts()
			
	print "__________________________________________________________________\n"

if __name__=="__main__":
	while 1:
		k=raw_input("nacisnij enter lub wpisz komende >> ")
		if len(k)==0:
			k=last_k
		last_k=deepcopy(k)
		if k=="q":
			break
		STATUS = proceed_command(k)
