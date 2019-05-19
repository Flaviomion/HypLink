Option Explicit
Const Versione = "1.4.0"
'dal 1.3.0 in avanti ho rivoluzionato i Flink per tutti i link (non riuscivo senza le informazioni da dove veniva il link a ricostruire i link spostati)
'ATTENZIONE questo succede quando uno spostamento va a posizionarsi proprio dove c'èra un altro hyperlink.
'ora i link saranno del tipo locali		HyFlink#XL#Pos=PS=xx#PC=jj#PR=rr#Punt=S=xx#C=jj#R=rr#HyElink
							'Remoti		HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
							'Remoti		HyFlink#FR#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
'Con la 1.4.0 sfrutto in nuovi Flink per gestire anche movimenti contemporanei del TO e FR.

Const Name = "HypLink"

Dim sheet, objExcel, objWorkbook, objWorkSecondbook, objWorksheet, objWorkSecondsheet, objRange, objShell, objStdOut
Dim out, report
Dim y, z, CDir, indice, indice_hyp, indice_Lflink, Uscita, sheet_n, n_fileIn, Errore
Dim ris, idx
Dim folderMaster, fileMaster, file_master, fileINI, hyp
Dim sheet_name
Dim aggiuntaLink, scrittoLinkNuovi, hyp_link_new
Dim ptx
Dim foglio_s, col_s, riga_s
Dim NomeProgramma
Dim Fdebug, att, bReadOnly
Dim settaFont, dimFont
att = false
Fdebug = true
NomeProgramma = Name&"_"&Versione '1.4.0 Versione che accetta movimenti su tutti i link in contemporanea
fileINI = "HypLink.ini"
Set objStdOut = WScript.StdOut
'folderMaster = WScript.Arguments.Item(0)
'fileMaster = WScript.Arguments.Item(1)
Dim listahyp(17,500)	 ' 0	sheet (dove è registrato il link)
                         ' 1	riga (dove è registrato il link)
                         ' 2	colonna numero (dove è registrato il link)
                         ' 3	colonna lettere (dove è registrato il link)
                         ' 4    Address = path relativo e nome del file linkato (relativo alla directory del file excel)
                         ' 5 	Path completo del file linkato
                         ' 6 	SubAddress nome_dello_sheet!colonna_lettereRiga
                         ' 7 	Nome sheet linked
                         ' 8	Colonna in lettere linked
                         ' 9 	Colonna numero linked
                         ' 10	Riga linked
                         ' 11	file linked parte iniziale
                         ' 12   link con listaFlink
                         ' 13   link con listaLinkedFlink
                         ' 14   link su se stesso
                         ' 15   7777 = riscontrato da incrocio, 8888 = link creato
                         ' 16   indica link a se stesso o se -1 link non completamente realizzato (spostato)



Dim listaFlink(17,500)	 ' 0	sheetLocale
                         ' 1	rigaLocale
                         ' 2	colLocNum
                         ' 3	colonnaLocale
                         ' 4    cartellaLinked
                         ' 5 	Path completo del file linkato cartellaLinked &"\"& fileCompletoLinked
                         ' 6 	simulazione SubAddress sh_name &"!"& colonnaLinked & rigaLinked
                         ' 7 	sh_name 'devo usare il nome
						 ' 8	colonnaLinked
                         ' 9 	colLinkedNum
                         ' 10	rigaLinked
                         ' 11   file linked parte iniziale
                         ' 12   link con listahyp
                         ' 13   link con listaLinkedFlink
						 ' 14	Posizione sheet
						 ' 15	Posizione colonna in lettere
						 ' 16	Posizione riga
						 ' 17	Parte iniziale del Nome del file di appartenenza


Dim listaLinkedFlink(17,500)  	 ' 0	sheetLocale
                                 ' 1	rigaLocale
                                 ' 2	colLocNum
                                 ' 3	colonnaLocale
                                 ' 4    cartellaLinked
                                 ' 5 	Path completo del file linkato cartellaLinked &"\"& fileCompletoLinked
                                 ' 6 	simulazione SubAddress sh_name &"!"& colonnaLinked & rigaLinked
                                 ' 7 	sh_name 'devo usare il nome
                                 ' 8	colonnaLinked
                                 ' 9 	colLinkedNum
                                 ' 10	rigaLinked
                                 ' 11   file linked parte iniziale
                                 ' 12   link con listahyp
                                 ' 13   link con listaLinkedFlink
								 ' 14	Posizione sheet
								 ' 15	Posizione colonna in lettere
								 ' 16	Posizione riga
								 ' 17	Parte iniziale del Nome del file di appartenenza

Dim listaFile(500) 	'0	File path completo

Dim listaSegnalazioni(100) 	' File segnalato


Const OK_BUTTON = 0
Const CRITICAL_ICON = 16
Const INFO_ICON_YN = 36
Const INFO_ICON = 64
Const AUTO_DISMISS = 0
Const AttesaMessaggioVV = 1
Const AttesaMessaggioV = 1
Const AttesaMessaggio = 1
Const AttesaMessaggioL = 1
Const AttesaMessaggioLL = 1

On Error Resume Next
objStdOut.Write "<font color ='blue'>Partenza "&NomeProgramma&"</font>"&vbCrLf
on error goto 0

inilistaSegnalazioni()
Set objShell = CreateObject("Wscript.Shell")
Set objExcel = CreateObject("Excel.Application")
aggiuntaLink = false
leggiINI folderMaster, fileMaster, report, Fdebug, settaFont, dimFont

dim fso: set fso = CreateObject("Scripting.FileSystemObject")
CDir = fso.GetAbsolutePathName(".")

objExcel.DisplayAlerts = 0
if (Fdebug) then
    objExcel.Visible = True
end if
file_master = cercaFile(fileMaster, folderMaster, fso)

ris = objShell.popup("Elaboro il file:" & folderMaster&"\"&file_master , 5, "Info", INFO_ICON_YN + 4096)
if (ris = 7) then
	On Error Resume Next
	objStdOut.Write "<font color ='orange'>Programma fermato dall'utente</font>"&vbCrLf
    on error goto 0
	Wscript.Quit 0
end if
On Error Resume Next
objStdOut.Write "Carico il file: "&folderMaster&"\"&file_master&vbCrLf
on error goto 0
Uscita = "<html><head><meta charset='utf-8' /><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /></head>"

Uscita = Uscita & "<body><table><tr><td colspan=10><h2><font color='blue'>Ripristina Hyperlink versione " & Versione & "</font></h2></td></tr>"
Uscita = Uscita & "<tr><td colspan=10><font color ='DarkGreen'>riposiziona collegamenti Ipertestuali</font></td></tr>"
Uscita = Uscita & "<tr><td align='center' colspan=10><font color ='DarkBlue'>" & folderMaster&"\"&file_master &"</font></td></tr>"
Uscita = Uscita & "<tr><td colspan=10><font color ='blue'>Rapporto del "&Date&" "&Hour(Now())&":"&Minute(Now())&"</font></td></tr>"
Uscita = Uscita & "<tr></tr>"
On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(folderMaster&"\"&file_master, False, False)
If (Err.Number <> 0) Then
	objStdOut.Write "<font color ='red'>Errore apertura file Master "&folderMaster&"\"&file_master&"</font><br/> Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&vbCrLf
    'objShell.popup "Errore apertura file Master "&folderMaster&"\"&file_master&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
    Err.Clear
End If
on error goto 0

bReadOnly = objWorkbook.ReadOnly
If bReadOnly = True Then
	On Error Resume Next
	objStdOut.Write "<font color ='red'>Errore apertura file "&folderMaster&"\"&file_master&" File OCCUPATO</font><br/>"&vbCrlf
	on error goto 0
    'objShell.popup "Errore apertura file "&folderMaster&"\"&file_master&" File OCCUPATO" , 5, "Errore", CRITICAL_ICON + 4096
	Call objWorkbook.Close
	objExcel.Quit
	On Error Resume Next
	objStdOut.Write  "<br/><font color ='blue'>Programma Terminato a causa di file Occupato</font>"&vbCrLf&vbCrLf
	on error goto 0
	Wscript.quit 1055
End If


sheet_n = objWorkbook.Sheets.Count

indice = 0 'indice viene incrementato da FindTO
indice_hyp = 0 'indice_hyp viene incrementato da FindHyper
indice_Lflink = 0 'indice_hyp viene incrementato da FindHyper

'HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
'Primo giro, raccolgo hyplink e Flink
On Error Resume Next
objStdOut.Write "<font color='blue'>Primo giro, raccolgo hyplink e Flink TO e TL e FL</font>"&vbCrLf
on error goto 0
raccogliHypFlink 'function che raccoglie link hyp e Flink TO 
Uscita = Uscita & "<tr><td align='center' colspan='10'><font color ='DarkBlue'>Link attivi</font></td></tr>"
for y = 0 to indice_hyp-1
    Uscita = Uscita & "<tr><td><font color ='green'>Link "&y&") </font></td><td><font color ='green'>"&folderMaster&"\"&fileMaster&"</font></td><td><font color ='green'> di:"&listahyp(0,y)&"!"&listahyp(3,y)&listahyp(1,y) &"</font></td><td><font color ='green'> a:"&listahyp(5,y)&" "&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font></td></tr>"
	on error resume next
	objStdOut.Write "<font color ='green'>Llink "&folderMaster&"\"&fileMaster&" di:"&listahyp(0,y)&"!"&listahyp(3,y)&listahyp(1,y) &" a:"&listahyp(5,y)&" "&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font>"&vbCrLf
	on error goto 0
Next

incrociaFlinkHyp ' primo incocio h e f

'SOLO PER TEST
'Call objWorkbook.Save
'Call objWorkbook.Close
'objExcel.Quit
'on error resume next
'objStdOut.Write "<font color ='blue'>Processo Concluso per TEST analisi primo giro/incrocio</font>"&vbCrLf&vbCrLf
'on error goto 0
'Wscript.Quit 0
'SOLO PER TEST


n_fileIn = popolaListaFile(listaFile, indice_hyp) 'Crea la lista dei file linked da listahyp con nome gia corretto

on error resume next
objStdOut.Write "<font color='blue'>Creo i Flink Locali mancanti</font>"&vbCrLf
on error goto 0
creaFlinkLocali 'Crea i HyFlink#XL su link locali

'Crea i HyFlink#TO e HyFlink#FR su quelli che puntano all'esterno
on error resume next
objStdOut.Write "<font color='blue'>Creo i Flink Remoti mancanti</font>"&vbCrLf
on error goto 0
for y = 0 to indice_hyp-1
    if (listahyp(14,y) = -1) then
	    if (listahyp(12,y) = -1) then 'non c'è ancora l'HyFlink#TO# e quindi neanche HyFlink#FR sul file linkato#
            if Not (scrittoLinkNuovi) then
                Uscita = Uscita & "<tr><td align='center' colspan='10'><font color ='Navy'>Link nuovi</font></td></tr>"
                scrittoLinkNuovi = true
		    end if
			On Error Resume Next
            Set objWorksheet = objWorkbook.Worksheets(listahyp(0,y))
		    If (Err.Number <> 0) Then
				objStdOut.Write "<font color ='red'>Errore creazione oggetto sheet Master</font><br/> Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&vbCrLf
				'objShell.popup "Errore creazione oggetto sheet Master Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                Err.Clear
		    end if
			on error goto 0
            On Error Resume Next
			objStdOut.Write "Creazione Flink "&listahyp(0,y)&"!"&listahyp(3,y)&listahyp(1,y)& vbCrLf
			on error goto 0
		    ris = scriviLinkTo( objWorksheet, listahyp, y, "TO")
		    if Not (ris) then
				On Error Resume Next
			    objStdOut.Write "<font color ='red'>Errore nella creazione di HyFlink su "&folderMaster&"\"&fileMaster&"</font>"&vbCrLf
				on error goto 0
			    'objShell.popup "Errore nella creazione di HyFlink#TO# su "&folderMaster&"\"&fileMaster , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                Err.Clear
		    end if
            'scrivo il link HyFlink#FR# sul file linkato
		    ris = scrivoLinkSuLinked(y,listahyp, true, "FR")
			On Error Resume Next
			objStdOut.Write "<font color ='green'>Creazione linkedFlink "&listahyp(5,y)&" "&listahyp(7,y)&"!"&listahyp(10,y)&listahyp(8,y)&"</font>"&vbCrLf
			on error goto 0
		    if Not (ris) then
			On Error Resume Next
				objStdOut.Write "<font color ='red'>Errore nella creazione di HyFlink#FR# su "&listahyp(5,y)&"</font>"&vbCrLf
			    'objShell.popup "Errore nella creazione di HyFlink#FR# su "&listahyp(5,y) , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                Err.Clear
		    end if
			on error goto 0
            Uscita = Uscita & "<tr><td><font color ='maroon'>Aggiunto Link</font></td><td><font color ='maroon'>" & folderMaster&"\"&fileMaster&"</font></td><td><font color ='maroon'> di:"&listahyp(0,y)&"!"&listahyp(3,y)&listahyp(1,y) &"</font></td><td><font color ='maroon'> a:"&listahyp(5,y)&" "&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font></td></tr>"
            aggiuntaLink = true
        end if
    end if
Next

if (aggiuntaLink) then
	On Error Resume Next
	objStdOut.Write "<font color='blue'>Secondo giro, raccolgo hyplink e Flink TO e TL e FL e FR</font>"&vbCrLf
	on error goto 0
    'vado a rileggere Flink e linkedFlink
	'On Error Resume Next
    clearLista listaFlink,indice, 14 'setto a zero l'indice dopo perchè l'indicazione serva a sapere cosa cancellare
    indice = 0
    clearLista listaLinkedFlink,indice_Lflink, 14 'setto a zero l'indice dopo perchè l'indicazione serva a sapere cosa cancellare
    indice_Lflink = 0
    for sheet = 1 to sheet_n step 1
        On Error Resume Next
	    Set objWorksheet = objWorkbook.Worksheets(sheet)
        If (Err.Number <> 0) Then
			objStdOut.Write "<font color ='red'>Errore creazione oggetto sheet file Master</font><br/> Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&vbCrLf
            'objShell.popup "Errore creazione oggetto sheet file Master Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
            Err.Clear
	    End If
		on error goto 0
		On Error Resume Next
	    sheet_name = objWorksheet.Name
	    If (Err.Number <> 0) Then
			objStdOut.Write "<font color ='red'>Errore estrae Nome del sheet file Master</font><br/> Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&vbCrLf
            'objShell.popup "Errore estrae Nome del sheet file Master Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
            Err.Clear
	    End If
		on error goto 0
	    ' popola la lista dei link Flink
	    FindToFrom objWorksheet, listaFlink, indice, sheet_name, 1 , folderMaster&file_master ' 1 = TO 
		FindToFrom objWorksheet, listaFlink, indice, sheet_name, 3 , folderMaster&file_master ' 3= TL e FL
    Next
	
	'La lista file la considero acquisita dai link hyp
    raccogliLinkedFlink
    incrociaFlinkHyp
else
	On Error Resume Next
	objStdOut.Write "<font color='blue'>Acquisisco LinkedFlink dopo Aggiunte</font>"&vbCrLf
	on error goto 0
    raccogliLinkedFlink 'non ci sono state aggiunte quindi popolo la lista linkedFlink
    incrociaFlinkHyp
end if


'SOLO PER TEST
'Call objWorkbook.Save
'Call objWorkbook.Close
'objExcel.Quit
'on error resume next
'objStdOut.Write "<font color ='blue'>Processo Concluso per TEST dopo creazione Flink mancanti e riacquisizione</font>"&vbCrLf&vbCrLf
'on error goto 0
'Wscript.Quit 0
'SOLO PER TEST

CambiaFile objWorkbook, sheet_n

scriviSu "ListaLink.txt", "+++++++++++++++Liste Pre ControlloLink++++++++++++" &vbCrLf
appendiA "ListaLink.txt", outListe

controlloLink

'SOLO PER TEST
appendiA "ListaLink.txt", "+++++++++++++++Liste Definitive++++++++++++" &vbCrLf
appendiA "ListaLink.txt", outListe
if (settaFont) then
	if ((CInt(dimFont) > 0 ) and (CInt(dimFont) < 60)) then
		settaFontPerCommenti
	end if
end if

Call objWorkbook.Save
Call objWorkbook.Close
objExcel.Quit

if Not (att) then
	''objShell.popup "Programma Terminato con Successo", AttesaMessaggio, "Fine", INFO_ICON + 4096
	On Error Resume Next
	objStdOut.Write  "<br/><font color ='blue'>Programma Terminato con Successo</font>"&vbCrLf&vbCrLf
	on error goto 0
else
	On Error Resume Next
	objStdOut.Write  "<br/><font color ='Orange'>Programma Terminato con alcune attenzioni</font>"&vbCrLf&vbCrLf
	on error goto 0
end if
Uscita = Uscita & "<table><body><html>"
scriviSu report, Uscita
objShell.run report 'Lancia l'eseguibile definito per il tipo di file da leggere.

Wscript.Quit 0


' ------------------------Funzioni ----------------------------------------------------------------------------------------------
function settaFontPerCommenti()
Dim ws, comm
  For Each ws In objWorkbook.Worksheets
	For Each comm In ws.Comments
        With comm.Shape.TextFrame.Characters.Font
            .Name = "Arial"
            .Size = dimFont
        End With
        comm.Shape.TextFrame.AutoSize = True
    Next
  Next
end function

function cercaContropartePOS(punt)
	'cerco la controparte che abbia come locazione POS il mio puntamento
	Dim z
	for z= 0 to indice-1
		if ((listaFlink(14,z) = listaFlink(7,punt)) and (listaFlink(15,z) = listaFlink(8,punt)) _ 
		    and (listaFlink(16,z) = listaFlink(10,punt))) then
		    cercaContropartePOS = z
            exit function
		end if
	Next
	cercaContropartePOS = -1
end function

function cercalinkedFlinkPOS(punt)
	'cerco la controparte che abbia come locazione POS il mio puntamento
	Dim z
	for z= 0 to indice-1
		if ((listaLinkedFlink(14,z) = listaFlink(7,punt)) and (listaLinkedFlink(15,z) = listaFlink(8,punt)) _ 
		    and (listaLinkedFlink(16,z) = listaFlink(10,punt))) then
		    cercalinkedFlinkPOS = z
            exit function
		end if
	Next
	cercalinkedFlinkPOS = -1
end function

function controlloLink()
Dim ris, hypCorr, hypControparte, Controparte, foglio_s, col_s, riga_s
	On Error Resume Next
    objStdOut.Write "<font color='blue'>Start controllo link</font>"&vbCrLf
	on error goto 0
    'Inizio cercando i link che si sono spostati (0,3,1 <> 14,15,16 su listaflink)
    for y = 0 to indice-1   'spazzola tutta la listaFlink in cerca di link spostati
        if ((listaFlink(0,y) <> listaFlink(14,y)) or (listaFlink(3,y) <> listaFlink(15,y)) or (listaFlink(1,y) <> listaFlink(16,y))) then
            'si è spostato in questo caso devo modificare:
            '                               se link Locale : l'Flink e l'Hyplink di chi mi puntava
            '                               se remoto      : llinkedFlink e l'Hyperlink sul file remoto 
            if (strComp(listaFlink(5,y),"LinkSuSeStesso") = 0) then
                'link Locale modificare l'Flink e l'Hyplink di chi mi puntava
                Controparte = cercaContropartePOS(y) 'Devo cercare il Flink che ha come POS il mio puntamento
                listaFlink(7,Controparte) = listaFlink(0,y)
                listaFlink(8,Controparte) = listaFlink(3,y)
                listaFlink(9,Controparte) = listaFlink(2,y)
                listaFlink(10,Controparte) = listaFlink(1,y)
                ' I POS vanno messi a posto 
				listaFlink(14,y) = listaFlink(0,y)
				listaFlink(15,y) = listaFlink(3,y)
				listaFlink(16,y) = listaFlink(1,y)
                ' modifico l'Flink sul file
				ris = scriviLinkTo(objWorksheet, listaFlink, Controparte, "XL")'serve a scrivere il giusto link sulla controparte
                ris = scriviLinkTo(objWorksheet, listaFlink, y, "XL") 'serve a scrivere la giusta POS sul link spostato
                ' modifico l'hyplink sul file
				modificaHypTo Controparte,listaFlink   
            else
                'remoto      : llinkedFlink e l'Hyperlink sul file remoto
                hypCorr = listaFlink(12,y)
                Controparte = cercalinkedFlinkPOS(y)
                if (Controparte <> -1) then
                    'adesso devo modificare sia il linkedFlink che l'hyperlink perchè puntino alla mia nuova posizione
                    foglio_s = listaLinkedFlink(7,Controparte)
			        col_s = listaLinkedFlink(8,Controparte)
			        riga_s = listaLinkedFlink(10,Controparte)
                    listaLinkedFlink(7,Controparte) = listaFlink(0,y)
                    listaLinkedFlink(8,Controparte) = listaFlink(3,y)
                    listaLinkedFlink(9,Controparte) = listaFlink(2,y)
                    listaLinkedFlink(10,Controparte) = listaFlink(1,y)
                    ' metto a posto il POS sulla listaflink
                    listaFlink(14,y) = listaFlink(0,y)
                    listaFlink(15,y) = listaFlink(3,y)
                    listaFlink(16,y) = listaFlink(1,y)
                    ' modifico l'Flink sul file
                    ris = scriviLinkTo(objWorksheet, listaFlink, y, "TO") 'serve a scrivere la giusta POS sul link spostato
                    ris = scrivoLinkSuLinkedPOS(y,listaFlink,listaFlink(5,y)) 'false per non eseguire anche l'hyperlink in modalità standard
			        if Not (ris) then
					    On Error Resume Next
				        objStdOut.Write "<font color ='red'>Errore nella creazione del HyFlink#FR# su:"&listaFlink(5,y)&"</font>"&vbCrLf
					    on error goto 0
			            'objShell.popup "Errore nella creazione del HyFlink#FR# su:"&listaFlink(5,y) , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                       Err.Clear
			        end if
			        'Corregge l'Hyperlink dal linkedFlink verso di me partendo dalla posizione RPos
			        modificaHypFromPOS y, listaFlink, listaLinkedFlink(0,Controparte), listaLinkedFlink(3,Controparte), listaLinkedFlink(1,Controparte), foglio_s,col_s, riga_s 
					'modificaHypFrom y, listaFlink, foglio_s, col_s, riga_s, false
                else
                    On Error Resume Next
				    objStdOut.Write "<font color ='red'>Errore Non trovo la controparte su listaLinkedFlink a:"&listaFlink(5,y)&" s:"&listaFlink(0,y)&" c:"&listaFlink(3,y)&" r:"&listaFlink(1,y)&"</font>"&vbCrLf
					on error goto 0
                end if
            end if
        end if
    Next
    for y=0 to indice_Lflink-1
        if ((listaLinkedFlink(0,y) <> listaLinkedFlink(14,y)) or (listaLinkedFlink(3,y) <> listaLinkedFlink(15,y)) or (listaLinkedFlink(1,y) <> listaLinkedFlink(16,y))) then
            'cerca in listaFlink un link che punta al mio POS, gli passo: parte iniziale file, sheetName,riga,colonna del mio POS
            Controparte = cercaFlinkDaHyp(listaLinkedFlink(17,y),listaLinkedFlink(14,y),listaLinkedFlink(16,y),listaLinkedFlink(15,y))
			'                             file parziale di questo linkedFlink,sheet di POS          , riga del POS         ,colonna del POS
            if (Controparte <> -1) then
                ' metto a posto il POS sulla listaLinkedFlink
                listaLinkedFlink(14,y) = listaLinkedFlink(0,y)
                listaLinkedFlink(15,y) = listaLinkedFlink(3,y)
                listaLinkedFlink(16,y) = listaLinkedFlink(1,y)
				'metto a posto il puntamento della listaFlink sulla mia nuova posizione
                listaFlink(7,Controparte) = listaLinkedFlink(0,y)
                listaFlink(8,Controparte) = listaLinkedFlink(3,y)
                listaFlink(9,Controparte) = listaLinkedFlink(2,y)
                listaFlink(10,Controparte) = listaLinkedFlink(1,y)

                ris = scrivoLinkSuLinked(Controparte,listaFlink,false,"FR")
                ris = scriviLinkTo(objWorksheet, listaFlink, Controparte, "TO")
                if Not (ris) then
					On Error Resume Next
				    objStdOut.Write "<font color ='red'>Errore nella creazione del HyFlink#TO# su:"&folderMaster&"\"&fileMaster&" s:"&listaFlink(0,Controparte)&" c:"&listaFlink(3,Controparte)&" r:"&listaFlink(1,Controparte)&"</font>"&vbCrLf
					on error goto 0
		        end if
                'Corregge l'Hyperlink
		        modificaHypTo Controparte, listaFlink
            else
                On Error Resume Next
				objStdOut.Write "<font color ='red'>Errore Non trovo la controparte su listaFlink a:"&listaLinkedFlink(11,y)&" s:"&listaLinkedFlink(0,y)&" c:"&listaLinkedFlink(3,y)&" r:"&listaLinkedFlink(1,y)&"</font>"&vbCrLf
			    on error goto 0
            end if
        end if
    Next

end function



function raccogliHypFlink()
	for sheet = 1 to sheet_n step 1
	On Error Resume Next
	Set objWorksheet = objWorkbook.Worksheets(sheet)
	sheet_name = objWorksheet.Name
	If (Err.Number <> 0) Then
		
		objStdOut.Write "<font color='red'>Errore creazione oggetto sheet file Master</font><br/> Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&vbCrLf
		
        'objShell.popup "Errore creazione oggetto sheet file Master Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
        Err.Clear
	End If
	on error goto 0
	On Error Resume Next
	objStdOut.Write "Leggo gli HypLink foglio "&sheet_name&vbCrLf
	on error goto 0
	FindHyper sheet,indice_hyp, sheet_name 'E il primo a trovare dei link
	On Error Resume Next
	objStdOut.Write "Leggo i Flink foglio "&sheet_name&vbCrLf
	on error goto 0
	FindToFrom objWorksheet, listaFlink, indice, sheet_name, 1 , folderMaster&file_master ' 1 = TO
	FindToFrom objWorksheet, listaFlink, indice, sheet_name, 3 , folderMaster&file_master ' 3= TL e FL
Next

end function


function creaHypLinkLocali()
    for y=0 to indice_hyp-1
        if ((listahyp(14,y) = -1) and (StrComp(listahyp(5,y),"LinkSuSeStesso") = 0)) then   'e un link locale e non si trova il corrispettivo
			'On Error Resume Next
            'objStdOut.Write "<font color='red'>Andrei a creare l'hyp "&listahyp(5,y)&"->"&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font><br/>"&vbCrLf
			'on error goto 0
            'Wscript.Echo "<font color='red'>Andrei a creare l'hyp "&listahyp(5,y)&"->"&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font><br/>"&vbCrLf
            creaHypXl y, listahyp
			hyp_link_new = true
        end if
    Next
end function


function creaFlinkLocali()
    for y=0 to indice_hyp-1 step 1
        if Not (listahyp(14,y) = -1) then 'Significa che è linkato con un altro hyp
            if (listahyp(15,y) = -1) then 'Signofica che non è stato ancora creato l'Flink
				if Not (scrittoLinkNuovi) then
					Uscita = Uscita & "<tr><td align='center' colspan='10'><font color ='Navy'>Link nuovi</font></td></tr>"
					scrittoLinkNuovi = true
				end if
				On Error Resume Next
                Set objWorksheet = objWorkbook.Worksheets(listahyp(0,y))
		        If (Err.Number <> 0) Then
				    objStdOut.Write "<font color ='red'>Errore creazione oggetto sheet Master</font><br/> Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&vbCrLf
					
			        'objShell.popup "Errore creazione oggetto sheet Master Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                    Err.Clear
		        end if
				on error goto 0
		        ris = scriviLinkTo(objWorksheet,listahyp,y,"XL")
				On Error Resume Next
			    objStdOut.Write "Creazione Flink Local "&listahyp(0,y)&"!"&listahyp(3,y)&listahyp(1,y)& vbCrLf
				on error goto 0
		        if Not (ris) then
					On Error Resume Next
			        objStdOut.Write "<font color ='red'>Errore nella creazione di HyFlink su "&folderMaster&"\"&fileMaster&"</font>"&vbCrLf
					on error goto 0
			        'objShell.popup "Errore nella creazione di HyFlink#TO# su "&folderMaster&"\"&fileMaster, AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                    Err.Clear
		        end if
                Uscita = Uscita & "<tr><td><font color='maroon'>Aggiunto Link Local </font></td><td><font color='maroon'>" & folderMaster&"\"&fileMaster&"</font></td><td><font color ='maroon'> di:"&listahyp(0,y)&"!"&listahyp(3,y)&listahyp(1,y) &"</font></td><td><font color ='maroon'> a:"&listahyp(5,y)&" "&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font></td></tr>"
                aggiuntaLink = true
                listahyp(15,y) = 8888
            end if
        end if
    Next

end function


function raccogliLinkedFlink()
Dim sheet_nf
	for y = 0 to n_fileIn-1 
        'On Error Resume Next
        if (Instr(1,listaFile(y),"LinkSuSeStesso") >0 ) then 'Il link è locale sullo stesso file
			'Non faccio niente i link Tl e FL sono già raccolti
        else
			On Error Resume Next
	        Set objWorkSecondbook = objExcel.Workbooks.Open(listaFile(y), False, False)
	        If (Err.Number <> 0) Then
				
			    objStdOut.Write "<font color ='red'>Errore nell'apertura file Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
				
                'objShell.popup "Errore nell'apertura file Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                Err.Clear
	        End If
			on error goto 0
			bReadOnly = objWorkSecondbook.ReadOnly
			If bReadOnly = True Then
				On Error Resume Next
				objStdOut.Write "<font color ='red'>Errore apertura file "&listaFile(y)&", File OCCUPATO</font><br/>"&vbCrlf
				on error goto 0
				'objShell.popup "Errore apertura file "&listaFile(y)&", File OCCUPATO" , 5, "Errore", CRITICAL_ICON + 4096
				Call objWorkSecondbook.Close
				Call objWorkbook.Close
				objExcel.Quit
				On Error Resume Next
				objStdOut.Write  "<br/><font color ='blue'>Programma Terminato a causa di file Occupato</font>"&vbCrLf&vbCrLf
				on error goto 0
				Wscript.quit 1055
			End If
	        sheet_nf = objWorkSecondbook.Sheets.Count
	        for sheet = 1 to sheet_nf step 1
				On Error Resume Next
		        Set objWorkSecondsheet = objWorkSecondbook.Worksheets(sheet)
		        If (Err.Number <> 0) Then
				    objStdOut.Write "<font color ='red'>Errore creazione oggetto sheet Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
					
                    'objShell.popup "Errore creazione oggetto sheet Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                    Err.Clear
		        End If
				on error goto 0
				On Error Resume Next
                sheet_name = objWorkSecondsheet.Name
                If (Err.Number <> 0) Then
				    objStdOut.Write "<font color ='red'>Errore Name da oggetto sheet Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
					
                    'objShell.popup "Errore Name da oggetto sheet Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                    Err.Clear
		        End If
				on error goto 0
		        'Set objRange = objWorkSecondsheet.UsedRange 'DAFARE verificare se serve
		        FindToFrom objWorkSecondsheet, listaLinkedFlink, indice_Lflink, sheet_name, 2 , listaFile(y) ' 2 = FROM
	        Next 'Loop su tutti gli sheet di un file Input
			On Error Resume Next
	        objWorkSecondbook.Close False,listaFile(y)
			If (Err.Number <> 0) Then
				    objStdOut.Write "<font color ='red'>Errore RaccogliLinkedFlink Close File Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
                    'objShell.popup "Errore RaccogliLinkedFlink Close File Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                    Err.Clear
		     End If
			 on error goto 0
        end if
    Next
end function

function CambiaFile(ByRef objWb, ByVal sheet_n) 'Modifica il nome su tutti gli hyperlink per il objWorkbook passato
Dim objWsh, sheet_name, hyp, lo_riga, lo_colonna, index_lf, Flink, posF, pos1, file, file_new, folder, info, sub_add,s_name,punt_flink, Part_ini

	for sheet = 1 to sheet_n step 1
        On Error Resume Next
		Set objWsh = objWb.Worksheets(sheet)
		If (Err.Number <> 0) Then
			objStdOut.Write "<font color ='red'>Errore select sheet  Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		    'objShell.popup "Errore select sheet  Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
            Err.Clear
        End If
		on error goto 0
		On Error Resume Next
		sheet_name = objWsh.Name
		If (Err.Number <> 0) Then
			objStdOut.Write "<font color ='red'>Errore Nome sheet  Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		    'objShell.popup "Errore Nome sheet  Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
            Err.Clear
        End If
		on error goto 0
		''objShell.popup "Sheet:" & sheet_name, AttesaMessaggioLL, "info", INFO_ICON + 4096  
		For Each hyp In objWsh.Hyperlinks
            'bisogna estrarre il SubAddress e cercarlo nella listaFlink da questa estrarre il nuovo nome file e proseguire
            if Not (Len(hyp.Address) = 0) then 'controlla che non sia un link su se stesso in questo caso non e necessario fare niente
			    SeparaRigheColonne hyp.Parent.Address(0, 0), lo_riga, lo_colonna
				On Error Resume Next
			    file = hyp.Address
                If (Err.Number <> 0) Then
			        objStdOut.Write "<font color ='red'>Errore estrazione Address di Hyperlink Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		            'objShell.popup "Errore estrazione Address di Hyperlink Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                    Err.Clear
                End If
				on error goto 0
				On Error Resume Next
                sub_add = hyp.SubAddress
                If (Err.Number <> 0) Then
			        objStdOut.Write "<font color ='red'>Errore estrazione SubAddress di Hyperlink Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		            'objShell.popup "Errore estrazione SubAddress di Hyperlink Descrizione: " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                    Err.Clear
                End If
				on error goto 0
                separSheetCR replace(sub_add,"'",""),s_name,lo_riga,lo_colonna
                Part_ini = estraiParteIniziale(hyp.Address)
			    'devo cercare anche per file parziale, altrimenti trovo un altro entry chee può semigliarci
                punt_flink = cercaFlinkDaHyp(Part_ini,s_name,lo_riga,lo_colonna) ' cerco l'entry di listaFlink corrispondente a questo hyplink
                if Not (punt_flink = -1) then
                    file_new = listaFlink(5,punt_flink)
			        if Not (Instr(1,UCase(hyp.Address),UCase(file_new)) > 0) then
			            info = "Address            :" & hyp.Address & vbCrLf & _
				               "SubAddress      :" & hyp.SubAddress & vbCrLf & _
				               "ScreenTip       :" & hyp.ScreenTip & vbCrLf & _
				               "TextToDisplay     :" & hyp.TextToDisplay& vbCrLf & _
				               "Flink             :"&Flink& vbCrLf & _
				               "Nuovo file         :"&file_new
			            hyp.Address = file_new
                        if (Fdebug) then
                            'objShell.popup info, AttesaMessaggioVV, "info", INFO_ICON + 4096
                        end if
                    end if
                end if
            end if 'controlla che non sia un link su se stesso
		Next
	Next
end function

function scrivoLinkSuLinkedPOS(ByVal y, ByRef lista_ref, ByVal file) 'scrive il Flink sul file remoto usando la giusta posizione e non quella precedente
Dim commento_esiste, comm, commento, bReadOnly
			On Error Resume Next
		    Set objWorkSecondbook = objExcel.Workbooks.Open(file, False, False)
		    If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: apertura Linked "&lista_ref(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
			    'objShell.popup "Errore scrivoLinkSuLinked: apertura Linked "&lista_ref(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    End If
			on error goto 0
			bReadOnly = objWorkSecondbook.ReadOnly
			If bReadOnly = True Then
				On Error Resume Next
				objStdOut.Write "<font color ='red'>Errore apertura file "&lista_ref(5,y)&", File OCCUPATO</font><br/>"&vbCrlf
				on error goto 0
				'objShell.popup "Errore apertura file "&lista_ref(5,y)&", File OCCUPATO" , 5, "Errore", CRITICAL_ICON + 4096
				Call objWorkSecondbook.Close
				Call objWorkbook.Close
				objExcel.Quit
				On Error Resume Next
				objStdOut.Write  "<br/><font color ='blue'>Programma Terminato a causa di file Occupato</font>"&vbCrLf&vbCrLf
				on error goto 0
				Wscript.quit 1055
			End If
			On Error Resume Next
            Set objWorkSecondsheet = objWorkSecondbook.Worksheets(lista_ref(0,y)) 'mi setto sul giusto sheet
            If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: set sheet " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore scrivoLinkSuLinked: set sheet " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    End If
			on error goto 0
			On Error Resume Next
	        commento_esiste = true
			comm = objWorkSecondsheet.Cells(lista_ref(1,y), lista_ref(2,y)).Comment.Text
	        If (Err.Number <> 0) Then
			    'Il commento non esiste ancora
                comm =""
                commento_esiste = false
                Err.Clear
	        End If
			on error goto 0
            if (commento_esiste) then
			    if (Instr(1,comm,"HyFlink#") > 0) then 'c'è già un link devo toglierlo
				    'rimuovo il link
				    commento = rimuoviLink(comm)
			    else
				    commento = comm
			    end if
            end if
	        commento = commento & vbCrLf&"HyFlink#FR#Pos=PS="&lista_ref(1,y)&"#PC="&lista_ref(3,y)&"#PR="&lista_ref(1,y)&"#cartella="&folderMaster&"#file="&fileMaster&"#S="&lista_ref(7,y)&"#C="&lista_ref(8,y)&"#R="&lista_ref(10,y)&"#HyElink"
	        commento = commento & vbCrLf&"FcomTime#"&Date&" "&Hour(Now())&":"&Minute(Now())&":"&Second(Now())&"#FcomTime"
            if (commento_esiste) then
                On Error Resume Next
			    objWorkSecondsheet.Cells(lista_ref(10,y), lista_ref(9,y)).ClearComments
			    If (Err.Number <> 0) Then
			        objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	        'objShell.popup "Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	            End If
			    on error goto 0
            end if
			On Error Resume Next
	        objWorkSecondsheet.Cells(lista_ref(10,y), lista_ref(9,y)).AddComment commento
	        If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	        End If
			on error goto 0
			On Error Resume Next
		    objWorkSecondbook.Save
		    If (Err.Number <> 0) Then
				
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: Save linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
				
			    'objShell.popup "Errore scrivoLinkSuLinked: Save linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    end if
			on error goto 0
			On Error Resume Next
		    objWorkSecondbook.Close
		    If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: Close Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
			    'objShell.popup "Errore scrivoLinkSuLinked: Close Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    end if
			on error goto 0
            scrivoLinkSuLinkedPOS = true
end function
		
function scrivoLinkSuLinked(ByVal y, ByRef lista_ref, ByVal hyp_yn, ByVal chiave) 'hyp se deve fare anche l'hyperlink
Dim idx,objSheet, commento, objLink,comm, commento_esiste
'Inizio la scrittura del HyFlink#FR# sul file linkato
		'On Error Resume Next
        if (Instr(1,lista_ref(5,y),"LinkSuSeStesso") >0 ) then 'Il link è locale sullo stesso file
            'devo lavorare sul file master aperto
			chiave = "XL"
			On Error Resume Next
            Set objWorksheet = objWorkbook.Worksheets(lista_ref(7,y)) 'mi setto sul giusto sheet
            If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore set sheet scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore set sheet scrivoLinkSuLinked Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    End If
			on error goto 0
	        'objWorksheet.Cells(lista_ref(10,y), lista_ref(9,y)).ClearComments
			On Error Resume Next
			comm = objWorksheet.Cells(lista_ref(10,y), lista_ref(9,y)).Comment.Text
	        If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore clear comment scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore clear comment scrivoLinkSuLinked Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
	        End If
			on error goto 0
			if (Instr(1,comm,"HyFlink#") > 0) then 'c'è già un link devo toglierlo
				'rimuovo il link
				commento = rimuoviLink(comm)
			else
				commento = comm
			end if
			
			'HyFlink#XL#Pos=S=xx#C=jj#R=rr#Punt=S=xx#C=jj#R=rr#HyElink
			commento = commento & vbCrLf&"HyFlink#"&chiave&"#Pos=PS="&lista_ref(7,y)&"#PC="&lista_ref(8,idx)&"#PR="&lista_ref(10,y)&"#Punt=S="&lista_ref(0,y)&"#C="&lista_ref(3,y)&"#R="&lista_ref(1,y)&"#HyElink"
	        commento = commento & vbCrLf&"FcomTime#"&Date&" "&Hour(Now())&":"&Minute(Now())&":"&Second(Now())&"#FcomTime"
			On Error Resume Next
			objWorksheet.Cells(lista_ref(10,y), lista_ref(9,y)).ClearComments
			If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore clear commento 3 scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore Clear commento 3 scrivoLinkSuLinked Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                Err.Clear
	        End If
			on error goto 0
			On Error Resume Next
            objWorksheet.Cells(lista_ref(10,y), lista_ref(9,y)).AddComment commento
	        If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore add comment scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore add comment scrivoLinkSuLinked Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                Err.Clear
	        End If
			on error goto 0
            if (hyp_yn) then
				On Error Resume Next
	            Set objLink = objWorksheet.Hyperlinks.Add(objWorkbook.Worksheets(objWorksheet.Name).Range("'"&lista_ref(7,y)&"'!"&lista_ref(8,y)&lista_ref(10,y)), _
                    "", _
                    "'"&lista_ref(0,y)&"'!"&lista_ref(3,y)&lista_ref(1,y), _
                    "hypCreato")
	            If (Err.Number <> 0) Then
				    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: di aggiunta"&lista_ref(7,y)&"!"&lista_ref(8,y)&lista_ref(10,y)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
					
	        	    'objShell.popup "Errore scrivoLinkSuLinked: di aggiunta"&lista_ref(7,y)&"!"&lista_ref(8,y)&lista_ref(10,y)& "Descrizione: "&Err.Description
				    Err.Clear
	            End If
				on error goto 0
            end if
            'Fine scrittura HyFlink#FR#
        else
			On Error Resume Next
		    Set objWorkSecondbook = objExcel.Workbooks.Open(lista_ref(5,y), False, False)
		    If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: apertura Linked "&lista_ref(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
				
			    'objShell.popup "Errore scrivoLinkSuLinked: apertura Linked "&lista_ref(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    End If
			on error goto 0
			bReadOnly = objWorkSecondbook.ReadOnly
			If bReadOnly = True Then
				On Error Resume Next
				objStdOut.Write "<font color ='red'>Errore apertura file "&lista_ref(5,y)&", File OCCUPATO</font><br/>"&vbCrlf
				on error goto 0
				'objShell.popup "Errore apertura file "&lista_ref(5,y)&", File OCCUPATO" , 5, "Errore", CRITICAL_ICON + 4096
				Call objWorkSecondbook.Close
				Call objWorkbook.Close
				objExcel.Quit
				On Error Resume Next
				objStdOut.Write  "<br/><font color ='blue'>Programma Terminato a causa di file Occupato</font>"&vbCrLf&vbCrLf
				on error goto 0
				Wscript.quit 1055
			End If
			On Error Resume Next
            Set objWorkSecondsheet = objWorkSecondbook.Worksheets(lista_ref(7,y)) 'mi setto sul giusto sheet
            If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: set sheet " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore scrivoLinkSuLinked: set sheet " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    End If
			on error goto 0
			On Error Resume Next
	        commento_esiste = true
			comm = objWorkSecondsheet.Cells(lista_ref(10,y), lista_ref(9,y)).Comment.Text
	        If (Err.Number <> 0) Then
			    'Il commento non esiste ancora
                comm =""
                commento_esiste = false
                Err.Clear
	        End If
			on error goto 0
            if (commento_esiste) then
			    if (Instr(1,comm,"HyFlink#") > 0) then 'c'è già un link devo toglierlo
				    'rimuovo il link
				    commento = rimuoviLink(comm)
			    else
				    commento = comm
			    end if
            end if
	        commento = commento & vbCrLf&"HyFlink#FR#Pos=PS="&lista_ref(7,y)&"#PC="&lista_ref(8,y)&"#PR="&lista_ref(10,y)&"#cartella="&folderMaster&"#file="&fileMaster&"#S="&lista_ref(0,y)&"#C="&lista_ref(3,y)&"#R="&lista_ref(1,y)&"#HyElink"
	        commento = commento & vbCrLf&"FcomTime#"&Date&" "&Hour(Now())&":"&Minute(Now())&":"&Second(Now())&"#FcomTime"
            if (commento_esiste) then
                On Error Resume Next
			    objWorkSecondsheet.Cells(lista_ref(10,y), lista_ref(9,y)).ClearComments
			    If (Err.Number <> 0) Then
			        objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	        'objShell.popup "Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	            End If
			    on error goto 0
            end if
			On Error Resume Next
	        objWorkSecondsheet.Cells(lista_ref(10,y), lista_ref(9,y)).AddComment commento
	        If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    	    'objShell.popup "Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	        End If
			on error goto 0
            if (hyp_yn) then
                'Crea L'Hyperlink
				On Error Resume Next
	            Set objLink = objWorkSecondsheet.Hyperlinks.Add(objWorkSecondbook.Worksheets(objWorkSecondsheet.name).Range("'"&lista_ref(7,y)&"'!"&lista_ref(8,y)&lista_ref(10,y)), _
                    folderMaster&"\"&file_master, _
                    "'"&lista_ref(0,y)&"'!"&lista_ref(3,y)&lista_ref(1,y), _
                    "hyplink="&folderMaster&"\"&file_master&"-"&lista_ref(0,y)&"'!"&lista_ref(3,y)&lista_ref(1,y))
	            If (Err.Number <> 0) Then
				    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: di aggiunta"&lista_ref(7,y)&"!"&lista_ref(8,y)&lista_ref(10,y)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
	        	    'objShell.popup "Errore scrivoLinkSuLinked: di aggiunta"&lista_ref(7,y)&"!"&lista_ref(8,y)&lista_ref(10,y)& "Descrizione: "&Err.Description
				    Err.Clear
	            End If
				on error goto 0
                'Fine della creazione dell'Hyperlink
            end if
			On Error Resume Next
		    objWorkSecondbook.Save
		    If (Err.Number <> 0) Then
				
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: Save linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
				
			    'objShell.popup "Errore scrivoLinkSuLinked: Save linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    end if
			on error goto 0
			On Error Resume Next
		    objWorkSecondbook.Close
		    If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore scrivoLinkSuLinked: Close Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
			    'objShell.popup "Errore scrivoLinkSuLinked: Close Linked Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    end if
			on error goto 0
        end if
		'Fine scrittura HyFlink#FR#
        scrivoLinkSuLinked = true ' in realtà se non è andata bene sono già uscito dal programma
end function

function inilistaSegnalazioni()
Dim h
    for h = 0 to 99
        listaSegnalazioni(h) = "#"
    Next
end function

function nomeFileSegnalato(ByVal file)
Dim h
    nomeFileSegnalato = false
    for h = 0 to 99
        if (Instr(1,listaSegnalazioni(h),file) > 0) then
            nomeFileSegnalato = true
            exit function
        else 
            if (Instr(1,listaSegnalazioni(h),"#") > 0) then
                listaSegnalazioni(h) = file
                exit function
            end if
        end if
    Next
end function

Function FindHyper(she_n,idx,s_loc_name)
Dim lo_riga, lo_colonna, li_riga, li_colonna, li_col_num, s_name, file_p_linked, folder, file_name
	'On Error Resume Next
	For Each hyp In objWorksheet.Hyperlinks
        SeparaRigheColonne hyp.Parent.Address(0, 0), lo_riga, lo_colonna	
		separSheetCR replace(hyp.SubAddress,"'",""),s_name,li_riga,li_colonna
		s_name = Replace(s_name,"'","")
		li_col_num = calcolaColonna(li_colonna)
        file_p_linked = estraiParteIniziale(hyp.Address) 'Se vuoto = LinkSuSeStesso
        folder = estraiFolderDaAddress(hyp.Address,file_p_linked) 'Se vuoto = LinkSuSeStesso
		if Not (Instr(1,folder,"LinkSuSeStesso") > 0) then
			file_name = cercaFile(file_p_linked,folder,fso)
			If (StrComp(file_name,"NULLA") = 0) Then
				On Error Resume Next
				objStdOut.Write "<font color ='red'>Errore cercaFile folder da hyp " & hyp.address &" NULLO", AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
				on error goto 0
				'objShell.popup "Errore cercaFile folder da hyp " & hyp.address &" NULLO" , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			end if
			if Not (Instr(1,hyp.Address,file_name) > 0) then
				if Not (nomeFileSegnalato(file_name)) then
					Uscita = Uscita & "<tr><td><font color ='orange'>Nome File modificato</font></td><td><font color ='orange'>da:"&hyp.Address&"</font></td><td></td><td><font color ='orange'> a:"&file_name&"</td></tr>"
				end if
			end if
		else
			file_name = "LinkSuSeStesso"
		end if
		listahyp(0,idx) = s_loc_name					            ' 0	sheet (dove è registrato il link)
        listahyp(1,idx) = lo_riga						            ' 1	riga (dove è registrato il link)
        listahyp(2,idx) = calcolaColonna(lo_colonna)	            ' 2	colonna numero (dove è registrato il link)
        listahyp(3,idx) = lo_colonna					            ' 3	colonna lettere (dove è registrato il link)
        listahyp(4,idx) = hyp.Address					            ' 4 Address = path relativo e nome del file linkato (relativo alla directory del file excel)
        if (Instr(1,file_name,"LinkSuSeStesso") > 0) then
			listahyp(5,idx) = "LinkSuSeStesso"
		else
			listahyp(5,idx) = folder&"\"&file_name
		end if														' 5 	Path completo del file linkato
        listahyp(6,idx) = hyp.SubAddress				            ' 6 	SubAddress nome_dello_sheet!colonna_lettereRiga
		listahyp(7,idx) = s_name						            ' 7 	Nome sheet linked
		listahyp(8,idx) = li_colonna					            ' 8		Colonna in lettere linked
		listahyp(9,idx) = li_col_num					            ' 9 	Colonna numero linked
		listahyp(10,idx) = li_riga						            ' 10	Riga linked
        listahyp(11,idx) = file_p_linked                            ' 11    file linked parte iniziale
        listahyp(12,idx) = -1                                       ' 12    link con listaFlink
        listahyp(13,idx) = -1                                       ' 13    link con listaLinkedFlink
        listahyp(14,idx) = -1                                       ' 14    link locale su listahyp
		listahyp(15,idx) = -1                                       ' 15    indica la situazione del Flink
        listahyp(16,idx) = -1                                       ' 16    indica se anche il puntamento di ritorno è giusto (spostamento)
		idx = idx+1
	Next
End Function

function creaHypXl(ByVal punt, ByRef lista)
Dim ws, objSheet, objLink
	On Error Resume Next
    Set objSheet = objWorkbook.Worksheets(lista(7,punt)) 'lo sheet dove risiede la controparte
	If (Err.Number <> 0) Then
	    objStdOut.Write "<font color ='red'>Errore creaHypXl: Sheet "&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
	    'objShell.popup "Errore creaHypXl: Sheet "&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description
	    Err.Clear
	End If
	on error goto 0
	On Error Resume Next
	Set objRange = objExcel.Range("'"&lista(7,punt)&"'!"&lista(8,punt)&lista(10,punt))
	If (Err.Number <> 0) Then
	    objStdOut.Write "<font color ='red'>Errore creaHypXl: Range "&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
	    'objShell.popup "Errore creaHypXl: Range "&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description
	    Err.Clear
	End If
	on error goto 0
	On Error Resume Next
	Set objLink = objSheet.Hyperlinks.Add(objRange, "", "'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt), "hypCreato")
	If (Err.Number <> 0) Then
	    objStdOut.Write "<font color ='red'>Errore creaHypXl: di aggiunta"&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
	    'objShell.popup "Errore creaHypXl: di aggiunta"&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description
	    Err.Clear
	End If
	objStdOut.Write "<font color ='maroon'>creaHypXl: CREATO Link locale "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"</font>"&vbCrLf
	on error goto 0
end function

function modificaHypTo(ByVal punt, ByRef lista)
Dim ws, objSheet, subb, subbTest
	'On Error Resume Next
    Set objSheet = objWorkbook.Worksheets(lista(0,punt))
    subb = objSheet.range("'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)).Hyperlinks(1).SubAddress
	if (InStr(lista(6,punt),"'") <> 0) then
		subbTest = subb
	else
        subbTest = Replace(subb,"'","")
	end if
    if (subbTest = lista(6,punt)) then
		On Error Resume Next
		objSheet.range("'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)).Hyperlinks(1).SubAddress = "'"&lista(7,punt)&"'!"&lista(8,punt)&lista(10,punt)
        If (Err.Number <> 0) Then
	        objStdOut.Write "<font color ='red'>Errore modificaHypTo: modifica SubAddress "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
	        'objShell.popup "Errore modificaHypTo: modifica SubAddress "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description
	        Err.Clear
	    End If
		on error goto 0
		On Error Resume Next
		objStdOut.Write "<font color ='orange'>modificaHypTo: Hyperlink modificato "&objSheet.range("'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)).Hyperlinks(1).SubAddress&"</font>"&vbCrLf
        Uscita = Uscita & "<tr><td colspan=10><font color ='orange'>Hyperlink modificato da:"&subb&" a:"&objSheet.range("'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)).Hyperlinks(1).SubAddress&"</font></td></tr>"
		on error goto 0
    else
		On Error Resume Next
		objStdOut.Write "<font color ='red'>Errore modificaHypTo: L'Hyperlink non corrisponde :"&objSheet.range("'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)).Hyperlinks(1).SubAddress&" diverso da:"&lista(6,punt)&"</font>"&vbCrLf
		on error goto 0
        'objShell.popup "Errore modificaHypTo:L'Hyperlink non corrisponde "&objSheet.range("'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)).Hyperlinks(1).SubAddress
		Err.Clear
    end if
end function

function modificaHypFromPOS(ByVal punt, ByRef lista, ByVal Rsh, ByVal Rcol, ByVal Rriga, ByVal foglio_s,ByVal col_s, ByVal riga_s )
'funzione che va a modificare l'hyperlink remoto puntando su listaFlink selezionata ma scrivendo sulla posizione reale della linkedFlink
'quindi gli passo la listaFlink selezionata ma gli passo anche la RPos della linkedFlink e il foglio,colonna,righa di dove ero prima
Dim objSheet, subb, objRange, objLink
	On Error Resume Next
	Set objWorkSecondbook = objExcel.Workbooks.Open(lista(5,punt), False, False)
	If (Err.Number <> 0) Then
	    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: apertura Linked "&lista(5,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    'objShell.popup "Errore modificaHypFrom: apertura Linked "&lista(5,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	    Err.Clear
	End If
	on error goto 0
	bReadOnly = objWorkSecondbook.ReadOnly
	If bReadOnly = True Then
		On Error Resume Next
		objStdOut.Write "<font color ='red'>Errore apertura file "&lista(5,punt)&", File OCCUPATO</font><br/>"&vbCrlf
		on error goto 0
		'objShell.popup "Errore apertura file "&lista(5,punt)&", File OCCUPATO" , 5, "Errore", CRITICAL_ICON + 4096
		Call objWorkSecondbook.Close
		Call objWorkbook.Close
		objExcel.Quit
		On Error Resume Next
		objStdOut.Write  "<br/><font color ='orange'>Programma Terminato a causa di file Occupato</font>"&vbCrLf&vbCrLf
		on error goto 0
		Wscript.quit 1055
	End If
	On Error Resume Next
    Set objSheet = objWorkSecondbook.Worksheets(Rsh)
	If (Err.Number <> 0) Then
	    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: setta sheet Linked "&lista(7,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    'objShell.popup "Errore modificaHypFrom: setta sheet Linked "&lista(7,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	    Err.Clear
	End If
	on error goto 0
	On Error Resume Next
    subb = objSheet.range("'"&Rsh&"'!"&Rcol&Rriga).Hyperlinks(1).SubAddress
	If (Err.Number <> 0) Then
	    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: Copia SubAddress Linked "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"su:"&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	    'objShell.popup "Errore modificaHypFrom: Copia SubAddress Linked "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"su:"&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	    Err.Clear
	End If
	on error goto 0
	if ((subb = "'"&foglio_s&"'!"&col_s&riga_s) or (subb = foglio_s&"!"&col_s&riga_s)) then
		On Error Resume Next
	    objSheet.range("'"&Rsh&"'!"&Rcol&Rriga).Hyperlinks(1).SubAddress = "'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)
	    If (Err.Number <> 0) Then
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: Modifica SubAddress Linked "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"su:"&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		    Err.Clear
	    End If
		on error goto 0
    else
		On Error Resume Next
	    objStdOut.Write "<font color ='red'>modificaHypFromPOS: L'Hyperlink in modifica non corrisponde a quello atteso "&subb&" invece di "&foglio_s&"'!"&col_s&riga_s&"</font>"&vbCrLf
		on error goto 0
	end if
	On Error Resume Next
	objWorkSecondbook.Save
    If (Err.Number <> 0) Then
		objStdOut.Write "<font color ='red'>Errore modificaHypFromPOS: Salva Linked "&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		Err.Clear
	 End If
     on error goto 0
	 On Error Resume Next
	 objWorkSecondbook.Close
     If (Err.Number <> 0) Then
		objStdOut.Write "<font color ='red'>Errore modificaHypFromPOS: Chiudi Linked "&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		Err.Clear
	 End If
	 on error goto 0
end function

function modificaHypFrom(ByVal punt, ByRef lista, ByVal foglio_s, ByVal col_s, ByVal riga_s, aggiungi)
Dim objSheet, subb, objRange, objLink
	'On Error Resume Next
    if (Instr(1,lista(5,punt),"LinkSuSeStesso") > 0) then 'e un link su se stesso
		On Error Resume Next
        Set objSheet = objWorkbook.Worksheets(lista(7,punt))
	    If (Err.Number <> 0) Then	
		    objStdOut.Write "<font color ='red'>modificaHypFrom: Errore setta sheet Linked SuSe "&lista(7,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		    'objShell.popup "Errore setta sheet Linked SuSe "&lista(7,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		    Err.Clear
	    End If
		on error goto 0
    else  'e un link normale
		On Error Resume Next
	    Set objWorkSecondbook = objExcel.Workbooks.Open(lista(5,punt), False, False)
	    If (Err.Number <> 0) Then
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: apertura Linked "&lista(5,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		    'objShell.popup "Errore modificaHypFrom: apertura Linked "&lista(5,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		    Err.Clear
	    End If
		on error goto 0
		bReadOnly = objWorkSecondbook.ReadOnly
		If bReadOnly = True Then
			On Error Resume Next
			objStdOut.Write "<font color ='red'>Errore apertura file "&lista(5,punt)&", File OCCUPATO</font><br/>"&vbCrlf
			on error goto 0
			'objShell.popup "Errore apertura file "&lista(5,punt)&", File OCCUPATO" , 5, "Errore", CRITICAL_ICON + 4096
			Call objWorkSecondbook.Close
			Call objWorkbook.Close
			objExcel.Quit
			On Error Resume Next
			objStdOut.Write  "<br/><font color ='orange'>Programma Terminato a causa di file Occupato</font>"&vbCrLf&vbCrLf
			on error goto 0
			Wscript.quit 1055
		End If
		On Error Resume Next
        Set objSheet = objWorkSecondbook.Worksheets(lista(7,punt))
	    If (Err.Number <> 0) Then
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: setta sheet Linked "&lista(7,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		    'objShell.popup "Errore modificaHypFrom: setta sheet Linked "&lista(7,punt)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		    Err.Clear
	    End If
		on error goto 0
    end if  'e un link su se stesso o normale
    if (aggiungi) then 'Aggiunta
	    'WScript.Echo "Vado in aggiunta"&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)
	    'n error resume next
		On Error Resume Next
	    Set objRange = objExcel.Range("'"&lista(7,punt)&"'!"&lista(8,punt)&lista(10,punt))
	    If (Err.Number <> 0) Then
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: Range "&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
		    'objShell.popup "Errore modificaHypFrom: Range "&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description
		    Err.Clear
	    End If
		on error goto 0
		On Error Resume Next
	    Set objLink = objSheet.Hyperlinks.Add(objRange, folderMaster&"\"&file_master, "'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt), "hypCreato")
	    If (Err.Number <> 0) Then
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: di aggiunta"&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description&"</font>"&vbCrLf
		    'objShell.popup "Errore modificaHypFrom: di aggiunta"&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)& "Descrizione: "&Err.Description
		    Err.Clear
	    End If
		on error goto 0
	    'WScript.Echo "Finita aggiunta"&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)
    else 'Modifica
		On Error Resume Next
        subb = objSheet.range("'"&lista(7,punt)&"'!"&lista(8,punt)&lista(10,punt)).Hyperlinks(1).SubAddress
	    If (Err.Number <> 0) Then
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: Copia SubAddress Linked "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"su:"&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		    'objShell.popup "Errore modificaHypFrom: Copia SubAddress Linked "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"su:"&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		    Err.Clear
	    End If
		on error goto 0
	    if ((subb = "'"&foglio_s&"'!"&col_s&riga_s) or (subb = foglio_s&"!"&col_s&riga_s)) then
			On Error Resume Next
	        objSheet.range("'"&lista(7,punt)&"'!"&lista(8,punt)&lista(10,punt)).Hyperlinks(1).SubAddress = "'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)
		    If (Err.Number <> 0) Then
			    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: Modifica SubAddress Linked "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"su:"&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
			    'objShell.popup "Errore modificaHypFrom: Modifica SubAddress Linked "&lista(7,punt)&"!"&lista(8,punt)&lista(10,punt)&"su:"&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
			    Err.Clear
		    End If
			on error goto 0
        else
			On Error Resume Next
		    objStdOut.Write "<font color ='red'>modificaHypFrom: L'Hyperlink in modifica non corrisponde a quello atteso "&subb&"</font>"&vbCrLf
			on error goto 0
            'objShell.popup "L'Hyperlink in modifica non corrisponde a quello atteso "&subb
	    end if
    end if
    if Not (Instr(1,lista(5,punt),"LinkSuSeStesso") > 0) then 'e un link su se stesso
		On Error Resume Next
	    objWorkSecondbook.Save
        If (Err.Number <> 0) Then
			
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: Salva Linked "&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
			
		    'objShell.popup "Errore modificaHypFrom: Salva Linked "&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		    Err.Clear
	    End If
		on error goto 0
		On Error Resume Next
	    objWorkSecondbook.Close
        If (Err.Number <> 0) Then
			
		    objStdOut.Write "<font color ='red'>Errore modificaHypFrom: Chiudi Linked "&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
			
		    'objShell.popup "Errore modificaHypFrom: Chiudi Linked "&lista(5,y)&" Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		    Err.Clear
	    End If
		on error goto 0
    end if  'e un link su se stesso
end function

function FindToFrom(ByRef sheet, ByRef lis, ByRef index, ByVal sh_name, ByVal typo, ByVal file)
	'sheet = oggetto, lis Lista su cui registrare, index indice nella lista, sheet_name nome dell0 sheet, typo = 1 = TO 2 = FROM
	Dim cmt, colLocNum,  rigaLocale, colonnaLocale, rigaLinked, colLinkedNum, colonnaLinked, sheetLinked
	Dim PosI, PosF, PosC, PosFi, PosSh, PosCo, PosRi, PosPSh, PosPCo, PosPRi, intermedio, temp, chiave
	Dim fileLinked, fileCompletoLinked, cartellaLinked
    Dim file_completo, pri, pco, psh
	'HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink tipo 1
	'HyFlink#FR#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink tipo 2
	'HyFlink#XL#Pos=PS=xx#PC=jj#PR=rr#Punt=S=xx#C=jj#R=rr#HyElink tipo 3
	'12345678901234567890
	'On Error Resume Next
	if (typo = 1) then
		chiave = "O" 'TO
	else
		if (typo = 3) then
			chiave = "L" ' TL and FL
		else
			chiave = "R" ' FR
		end if
	end if
	For Each cmt In sheet.Comments
		'WScript.Echo "Loop:" & index
		PosI = InStr(1,cmt.text,"HyFlink#",1)
		PosF = InStr(1,cmt.text,"#HyElink",1)
		if (PosI > 0) then 'è un link
			SeparaRigheColonne cmt.Parent.Address(0, 0), rigaLocale, colonnaLocale
			colLocNum = calcolaColonna(colonnaLocale)
			if (Mid(cmt.text,PosI+9,1) = chiave) then  'è un link chiave
				PosC = InStr(1,cmt.text,"#cartella=",1)
				PosFi = InStr(1,cmt.text,"#file=",1)
				if (chiave = "L") then
					PosSh = InStr(1,cmt.text,"=S=",1)
					cartellaLinked = "LinkSuSeStesso"
				else
					PosSh = InStr(1,cmt.text,"#S=",1)
				end if
				PosCo = InStr(1,cmt.text,"#C=",1)
				PosRi = InStr(1,cmt.text,"#R=",1)
				PosPSh = InStr(1,cmt.text,"=PS=",1)
				PosPCo = InStr(1,cmt.text,"#PC=",1)
				PosPRi = InStr(1,cmt.text,"#PR=",1)
				if (PosPRi > 0) then
					intermedio = Mid(cmt.text,PosPRi+4,Len(cmt.text)-(PosPRi+4))
					pri = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link lo sheet di posizione
				else
					On Error Resume Next
					objStdOut.Write "<font color ='red'>Attenzione HyFlink senza sheet di posizione "&cmt.text&"</font>"&vbCrLf
					on error goto 0
					'objShell.popup "HyFlink senza sheet di posizione "&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
					att = true
				end if
				if (PosPSh > 0) then
					intermedio = Mid(cmt.text,PosPSh+4,Len(cmt.text)-(PosPSh+4))
					psh = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link lo sheet di posizione
				else
					On Error Resume Next
					objStdOut.Write "<font color ='red'>Attenzione HyFlink senza sheet di posizione "&cmt.text&"</font>"&vbCrLf
					on error goto 0
					'objShell.popup "HyFlink senza sheet di posizione "&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
					att = true
				end if
				if (PosPCo > 0) then
					intermedio = Mid(cmt.text,PosPCo+4,Len(cmt.text)-(PosPCo+4))
					pco = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link lo sheet di posizione
				else
					On Error Resume Next
					objStdOut.Write "<font color ='red'>Attenzione HyFlink senza Colonna di posizione "&cmt.text&"</font>"&vbCrLf
					on error goto 0
					'objShell.popup "HyFlink senza colonna di posizione "&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
					att = true
				end if
				if Not (chiave = "L") then
					if (PosC > 0) then
						intermedio = Mid(cmt.text,PosC+10,Len(cmt.text)-(PosC+10))
						cartellaLinked = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link il nome della cartellaLinked
					else
						On Error Resume Next
						objStdOut.Write "<font color ='red'>Attenzione HyFlink senza cartella"&cmt.text&"</font>"&vbCrLf
						on error goto 0
						'objShell.popup "HyFlink senza cartella"&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
						att = true
					end if
					if (PosFi > 0) then
						intermedio = Mid(cmt.text,PosFi+6,Len(cmt.text)-(PosFi+6))
						fileLinked = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link il nome del fileLinked
					else
						On Error Resume Next
						objStdOut.Write "<font color ='red'>Attenzione HyFlink senza file"&cmt.text&"</font>"&vbCrLf
						on error goto 0
						'objShell.popup "HyFlink senza file"&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
						att = true
					end if
				end if
				if (PosSh > 0) then
					intermedio = Mid(cmt.text,PosSh+3,Len(cmt.text)-(PosSh+3))
					sheetLinked = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link il nome del sheetLinked
				else
					On Error Resume Next
					objStdOut.Write "<font color ='red'>Attenzione HyFlink senza Sheet"&cmt.text&"</font>"&vbCrLf
					on error goto 0
					'objShell.popup "HyFlink senza Sheet"&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
					att = true
				end if
				if (PosCo > 0) then
					intermedio = Mid(cmt.text,PosCo+3,Len(cmt.text)-(PosCo+3))
					colonnaLinked = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link la colonnaLinked
					colLinkedNum = calcolaColonna(colonnaLinked)
				else
					On Error Resume Next
					objStdOut.Write "<font color ='red'>Attenzione HyFlink senza Colonna"&cmt.text&"</font>"&vbCrLf
					on error goto 0
					'objShell.popup "HyFlink senza Colonna"&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
					att = true
				end if
				if (PosRi > 0) then
					intermedio = Mid(cmt.text,PosRi+3,Len(cmt.text)-(PosRi+3))
					rigaLinked = Mid(intermedio,1,Instr(1,intermedio,"#",1)-1) 'estrae dal link la rigaLinked
				else
					On Error Resume Next
					objStdOut.Write "<font color ='red'>Attenzione HyFlink senza Riga "&cmt.text&"</font>"&vbCrLf
					on error goto 0
					'objShell.popup "Attenzione HyFlink senza Riga "&cmt.text, AttesaMessaggioL, "Attenzione", INFO_ICON + 4096
					att = true
				end if
				if (strComp(cartellaLinked,"LinkSuSeStesso") = 0) then
					cartellaLinked = "LinkSuSeStesso"
					file_completo = "LinkSuSeStesso"
				else
					if (fso.FolderExists(cartellaLinked)) then
						On Error Resume Next
						fileCompletoLinked = cercaFile(fileLinked, cartellaLinked , fso)
						If (Err.Number <> 0) Then
								objStdOut.Write "<font color ='red'>Errore cercaFile Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
								'objShell.popup "Errore Errore cercaFile Descrizione " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
								Err.Clear
						end If
						on error goto 0
						file_completo = cartellaLinked &"\"& fileCompletoLinked
					else
						att = true
						On Error Resume Next
						objStdOut.Write "<font color ='red'>Attenzione DIR-NON-VALIDA: " & cartellaLinked&" Link:"&cmt.text&" da:"&cmt.Parent.Address(0, 0)&"</font>"&vbCrLf
						on error goto 0
						'objShell.popup "Attenzione " & "DIR-NON-VALIDA: " & cartellaLinked&" Link:"&cmt.text&" da:"&cmt.Parent.Address(0, 0), AttesaMessaggio, "errore", CRITICAL_ICON + 4096
						Uscita = Uscita & "<tr><td colspan=10><font color=red> Attenzione la cartella "&cartellaLinked&" Link:"&cmt.text&" da:"&cmt.Parent.Address(0, 0)&" non esiste</font></td></tr>"
					end If
				end if
				lis(0,index) = sh_name
				lis(1,index) = rigaLocale
				lis(2,index) = colLocNum
				lis(3,index) = colonnaLocale
				lis(4,index) = cartellaLinked
				lis(5,index) = file_completo
				lis(6,index) = sheetLinked &"!"& colonnaLinked & rigaLinked
				lis(7,index) = sheetLinked 'devo usare il nome
				lis(8,index) = colonnaLinked
				lis(9,index) = colLinkedNum
				lis(10,index) = rigaLinked
				lis(11,index) = fileLinked
				lis(12,index) = -1   'link con listahyp
				lis(13,index) = -1   'link con listaFlink
				lis(14,index) = psh  'Posizione sheet      La posizione serve ad evidenziar lo spostamento
				lis(15,index) = pco  'Posizione Colonna (lettere)
				lis(16,index) = pri  'Posizione riga
                lis(17,index) = estraiParteIniziale(file)  'parte iniziale file di appartenenza
				index = index+1
			end if 'è un link HyFlink#TO#
		end if 'è un link
	Next
end function

function incrociaFlinkHyp()
'Crea i link fra la listahyp e le liste listaFlink e listaLinkedFlink
'Cerca anche incroci interni a listahyp nel caso di link su se stesso
'On Error Resume Next
Dim y,i,x,  puntatore, controparte
    for y= 0 to indice_hyp-1
		for x = 0 to indice_hyp-1
			if ((StrComp(listahyp(5,y),"LinkSuSeStesso") = 0) and (StrComp(listahyp(5,x),"LinkSuSeStesso") = 0)) then
				if ((listahyp(14,x) = -1) and (listahyp(14,y) = -1)) then
					if ((listahyp(0,y) = listahyp(7,x)) and (listahyp(1,y) = listahyp(10,x)) and (listahyp(2,y) = listahyp(9,x)) and (listahyp(3,y) = listahyp(8,x))_
						and (listahyp(7,y) = listahyp(0,x)) and (listahyp(10,y) = listahyp(1,x)) and (listahyp(9,y) = listahyp(2,x)) and (listahyp(8,y) = listahyp(3,x))) then
						'sono effettivamente uno la controparte dell'altro
						listahyp(14,x) = y
						listahyp(14,y) = x
						listahyp(16,x) = y
						listahyp(16,y) = x
                    end if
                end if
            end if
        Next
    Next
    'A questo punto ho associato tutti quelli che sono senza problemi

    scriviSu "ListaLink.txt", "+++++++++++++++Liste PTest dopo mezzo incrocio++++++++++++" &vbCrLf
    appendiA "ListaLink.txt", outListe
    'SOLO PER TEST
    'Call objWorkbook.Save
    'Call objWorkbook.Close
    'objExcel.Quit
    'on error resume next
    'objStdOut.Write "<font color ='blue'>Processo Concluso per TEST</font>"&vbCrLf&vbCrLf
    'on error goto 0
    'Wscript.Quit 0
    'SOLO PER TEST

    for y= 0 to indice_hyp-1 'Trovo i Flink che corrispondono con gli hyp (Anche quelli locali)
        for i = 0 to indice-1
                if ((listahyp(0,y) = listaFlink(0,i)) and (listahyp(1,y) = listaFlink(1,i)) and (listahyp(2,y) = listaFlink(2,i))) then
					listaFlink(12,i) = y
                    listahyp(12,y) = i
					if (listahyp(14,y) <> -1) then
                        if Not (listahyp(15,y) <> -1) then ' se ha già 9999 o 8888 glieli lascio
						    listahyp(15,y) = 7777 ' setto che l'Flink è già creato
                        end if
					end if
			    end if
            'end if
        Next
    Next
	for y= 0 to indice_hyp-1 'Trovo i LinkedFlink che corrispondono con gli hyp
		for x = 0 to indice_Lflink-1
            if (listahyp(14,y) = -1) then
			    if ((listahyp(0,y) = listaLinkedFlink(7,x)) and (listahyp(1,y) = listaLinkedFlink(10,x)) and (listahyp(2,y) = listaLinkedFlink(9,x))) then
					    listaLinkedFlink(12,x) = y
					    listahyp(13,y) = x
			    end if
            end if
		Next
	Next
	for y= 0 to indice-1 'Trovo i LinkedFlink che corrispondono con i Flink
		for x = 0 to indice_Lflink-1
			if ((listaFlink(0,y) = listaLinkedFlink(7,x)) and (listaFlink(1,y) = listaLinkedFlink(10,x)) and (listaFlink(2,y) = listaLinkedFlink(9,x))) then
                    listaLinkedFlink(13,x) = y
                    listaFlink(13,y) = x
			end if
		Next
	Next
end function

function cercaControparte(ByVal punt)
	'cerco la controparte che abbia come locazione il mio puntamento
	Dim z
	for z= 0 to indice_hyp-1
        if ((listahyp(14,z) = -1) and (listahyp(16,z) = -1)) then
		    if ((listahyp(0,z) = listahyp(7,punt)) and (listahyp(1,z) = listahyp(10,punt)) _ 
			    and (listahyp(2,z) = listahyp(9,punt)) and (listahyp(3,z) = listahyp(8,punt))) then
			    cercaControparte = z
                exit function
		    end if
        end if
	Next
	cercaControparte = -1
end function

function cercalinkedFlink(y)
Dim k
'On Error Resume Next
    cercalinkedFlink = -1
	for k = 0 to indice_Lflink-1
		if ((listaLinkedFlink(0,k) = listaFlink(7,y)) and _
			(listaLinkedFlink(3,k) = listaFlink(8,y)) and _
			(listaLinkedFlink(2,k) = listaFlink(9,y) ) and _
			(listaLinkedFlink(1,k) = listaFlink(10,y))) then
			cercalinkedFlink = k
			exit for
		end if
	Next
end function


function cercaFlink(ByVal sh,ByVal r,ByVal c)
Dim k
'On Error Resume Next
    cercaFlink = -1
	for k = 0 to indice-1
		if ((listaFlink(0,k) = sh) and _
			(listaFlink(1,k) = r) and _
			(listaFlink(3,k) = c)) then
			cercaFlink = k
			exit for
		end if
	Next
end function

function cercaFlinkDaHyp(ByVal p_i,ByVal sh,ByVal r,ByVal c)
Dim k
'On Error Resume Next
    cercaFlinkDaHyp = -1
	for k = 0 to indice-1
		if ((listaFlink(7,k) = sh) and _
			(listaFlink(10,k) = r) and _
            (Ucase(listaFlink(11,k)) = Ucase(p_i)) and _
			(listaFlink(8,k) = c)) then
			cercaFlinkDaHyp = k
			exit for
		end if
	Next
end function

function clearLista(ByRef lista, ByVal ind, ByVal k)
Dim z,x
'On Error Resume Next
    for z = 0 to ind
        for x = 0 to k
            lista(x,z) = Empty
        Next
    Next
end function

function estraiParteIniziale(address)
Dim pos, temp, out
'On Error Resume Next
if (Len(address) = 0) then
	estraiParteIniziale = "LinkSuSeStesso"
else
    Pos = InStrRev(address,"\")
    if (Pos = 0) then
        Pos = InStrRev(address,"/")
    end if
    if (Pos <> 0) then
        temp = Mid(address,Pos+1,Len(address)-Pos)
    else
        on error resume next
        objStdOut.Write "<font color ='red'>Attenzione Indirizzo del hyperlink non secondo standard : "&address&"</font>"&vbCrlf
        objStdOut.Write "<font color ='orange'>uscita dal programma per errore</font>"&vbCrlf
        Wscript.Quit 777
    end if
    Pos = Instr(1,temp,"'")
    if (Pos > 0) then 'se non c'è l'apice mantengo tutto il nome del file
        estraiParteIniziale = Mid(temp,1,Pos-1)
    else
        estraiParteIniziale = Mid(temp,1,Instr(1,temp,".")-1)
    end if
end if
end function

function outListe()
'On Error Resume Next
	out = "##################Lista HyperLink########################" & vbCrLf
	for y = 0 to indice_hyp-1 step 1
        out = out & "hhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhh" & vbCrLf
        out = out & "--  Indice  -------------------:" & y  & vbCrLf
		out = out & "0)  Sheet locale               :" & listahyp(0,y) & vbCrLf
		out = out & "1)  Riga locale                :" & listahyp(1,y) & vbCrLf
		out = out & "2)  Colonna num                :" & listahyp(2,y) & vbCrLf
		out = out & "3)  Colonna lett               :" & listahyp(3,y) & vbCrLf
		out = out & "4)  File link (rel)            :" & listahyp(4,y) & vbCrLf
		out = out & "5)  File link comp             :" & listahyp(5,y) & vbCrLf
		out = out & "6)  SubAddress                 :" & listahyp(6,y) & vbCrLf
		out = out & "7)  Sheet name link            :" & listahyp(7,y) & vbCrLf
		out = out & "8)  Col link lett              :" & listahyp(8,y) & vbCrLf
		out = out & "9)  Col link num               :" & listahyp(9,y) & vbCrLf
		out = out & "10) Riga linked                :" & listahyp(10,y) & vbCrLf
        out = out & "11) File link parziale         :" & listahyp(11,y) & vbCrLf
        out = out & "12) Link su listaFlink         :" & listahyp(12,y)  & vbCrLf
        out = out & "13) Link su listaLinkedFlink   :" & listahyp(13,y)  & vbCrLf
        out = out & "14) Link su se stesso          :" & listahyp(14,y)  & vbCrLf
        out = out & "15) situaz: 7777 o 8888 o 9999 :" & listahyp(15,y)  & vbCrLf
        out = out & "16) situaz link                :" & listahyp(16,y)  & vbCrLf
	Next
    out = out & vbCrLf
	out = out &"##################Lista HyFlink########################" & vbCrLf
	for y = 0 to indice-1 step 1
        out = out & "fffffffffffffffffffffffffffffffffffffffffffffffffffffff" & vbCrLf
        out = out & "--  Indice  -------------------:" & y  & vbCrLf
		out = out & "0)  Sheet locale               :" & listaFlink(0,y)  & vbCrLf
		out = out & "1)  Riga locale                :" & listaFlink(1,y)  & vbCrLf
		out = out & "2)  Colonna num                :" & listaFlink(2,y)  & vbCrLf
		out = out & "3)  Colonna lett               :" & listaFlink(3,y)  & vbCrLf
		out = out & "4)  Cartella Link              :" & listaFlink(4,y)  & vbCrLf
		out = out & "5)  File link comp             :" & listaFlink(5,y)  & vbCrLf
		out = out & "6)  sim subAddress             :" & listaFlink(6,y)  & vbCrLf
		out = out & "7)  Sheet name link            :" & listaFlink(7,y)  & vbCrLf
		out = out & "8)  Col link lett              :" & listaFlink(8,y)  & vbCrLf
		out = out & "9)  Col link num               :" & listaFlink(9,y)  & vbCrLf
		out = out & "10) Riga linked                :" & listaFlink(10,y)  & vbCrLf
        out = out & "11) File link parziale         :" & listaFlink(11,y)  & vbCrLf
        out = out & "12) Link su listahyp           :" & listaFlink(12,y)  & vbCrLf
        out = out & "13) Link su listaLinkedFlink   :" & listaFlink(13,y)  & vbCrLf
        out = out & "14) Posizione sheet            :" & listaFlink(14,y)  & vbCrLf
		out = out & "14) Posizione colonna          :" & listaFlink(15,y)  & vbCrLf
		out = out & "14) Posizione riga             :" & listaFlink(16,y)  & vbCrLf
	Next
    out = out & vbCrLf
	out = out & "##################Lista LinkedFlink########################" & vbCrLf
	for y = 0 to indice_Lflink-1 step 1
        out = out & "lflflflflflflflflflflflflflflflflflflflflflflflf" & vbCrLf
        out = out & "--  Indice  -------------------:" & y  & vbCrLf
		out = out & "0)  Sheet locale               :" & listaLinkedFlink(0,y)  & vbCrLf
		out = out & "1)  Riga locale                :" & listaLinkedFlink(1,y)  & vbCrLf
		out = out & "2)  Colonna num                :" & listaLinkedFlink(2,y)  & vbCrLf
		out = out & "3)  Colonna lett               :" & listaLinkedFlink(3,y)  & vbCrLf
		out = out & "4)  Cartella Link              :" & listaLinkedFlink(4,y)  & vbCrLf
		out = out & "5)  File link comp             :" & listaLinkedFlink(5,y)  & vbCrLf
		out = out & "6)  sim subAddress             :" & listaLinkedFlink(6,y)  & vbCrLf
		out = out & "7)  Sheet name link            :" & listaLinkedFlink(7,y)  & vbCrLf
		out = out & "8)  Col link lett              :" & listaLinkedFlink(8,y)  & vbCrLf
		out = out & "9)  Col link num               :" & listaLinkedFlink(9,y)  & vbCrLf
		out = out & "10) Riga linked                :" & listaLinkedFlink(10,y)  & vbCrLf
        out = out & "11) File link parziale         :" & listaLinkedFlink(11,y)  & vbCrLf
        out = out & "12) Link su listahyp			:" & listaLinkedFlink(12,y)  & vbCrLf
        out = out & "13) Link su listaFlink			:" & listaLinkedFlink(13,y)  & vbCrLf
        out = out & "14) Posizione sheet            :" & listaLinkedFlink(14,y)  & vbCrLf
		out = out & "14) Posizione colonna          :" & listaLinkedFlink(15,y)  & vbCrLf
		out = out & "14) Posizione riga             :" & listaLinkedFlink(16,y)  & vbCrLf
	Next
    out = out & vbCrLf
	out = out & "################Lista Files Linked########################" & vbCrLf
	for y = 0 to n_fileIn-1 step 1
        out = out & "--------------------------------------------------------" & vbCrLf
        out = out & "---Indice  --------------------:" & y  & vbCrLf
		out = out & "   File                        :"& listaFile(y) &vbCrLf
	Next
    outListe = out
end function

function separSheetCR(scr,ByRef s_name,ByRef riga,ByRef colonna)
Dim piece
'On Error Resume Next
	piece = split(scr,"!")
	s_name = piece(0)
	SeparaRigheColonne piece(1), riga, colonna
end function

function creaPathCompleto(address,master)
Dim i, id_n, piece, out, p, temp
'On Error Resume Next
	if (Instr(1,address,"..") > 0) then
		piece = Split(master,"\")
		id_n = UBound(piece)
		out = piece(0)
		for i = 1 to (id_n - 1)
			out = out & "\" & piece(i)
		next
        p = Instr(1,address,"\")
        temp = Mid(address,p,Len(address)-(p-1))
        out = out & temp
	else
        if ((Instr(1,address,":\") > 0) or (Instr(1,address,"\\") > 0)) then
            out = address
        else
		    out = master & "\" & address
        end if
	end if
	creaPathCompleto = out
end function

Function leggiINI(ByRef folderOutput, ByRef fileOutput, ByRef repOutput, ByRef debug, ByRef settaFont, ByRef dimFont)
Dim objFileToRead, linea, x, debug_st, temp
'On Error Resume Next
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(fileINI,1)
	for x = 1 to 6
		linea = objFileToRead.ReadLine()
		if (Instr(1,linea,"cartella=",1) > 0)then
			folderOutput = Mid(linea,10,Len(linea)-9)
		else
			if (Instr(1,linea,"file=",1) > 0)then
				fileOutput = Mid(linea,6,Len(linea)-5)
			else
				if (Instr(1,linea,"rapporto=",1) > 0)then
					repOutput = Mid(linea,10,Len(linea)-9)
				else
					if (Instr(1,linea,"debug=",1) > 0)then
						debug_st = Mid(linea,7,Len(linea)-6)
						if ((Instr(1,Lcase(debug_st),"si",1) > 0) or  (Instr(1,Lcase(debug_st),"yes",1))) then
							debug = true
						else
							debug = false
						end if
					else
						if (Instr(1,linea,"settaFont=",1) > 0)then
							temp = Mid(linea,11,Len(linea)-10)
							if (strComp(temp,"si") = 0) then
								settaFont = true
							end if
						else
							if (Instr(1,linea,"dimFont=",1) > 0)then
								dimFont = Mid(linea,9,Len(linea)-8)
							end if 'dimFont
						end if 'settaFont
					end if 'debug
						
				end if 'rapporto
			end if 'file
		end if 'cartelle
	Next
	objFileToRead.Close
	Set objFileToRead = Nothing
End Function

Function calcolaColonna(ByVal nome_colonna)
'On Error Resume Next
	Dim tot, ci, c, i
	tot = 0
	ci = 0
	if (Len(nome_colonna) > 1) then
		if (Len(nome_colonna) > 2) then
			if (Len(nome_colonna) > 3) then
				On Error Resume Next
				objStdOut.Write "<font color ='olive'>Considero le colonne solo fino alla terza lettera</font>"&vbCrLf
				on error goto 0
				'objShell.popup "Considero le colonne solo fino alla terza lettera" , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
                Att = true
				calcolaColonna = "+ZZZ"
				Exit Function
			end if
			'WScript.Echo "siamo a 3"
			c = Mid(nome_colonna,1,1)
			ci = CInt(Asc(UCase(c))-64)
			ci = ci * 676
			tot = tot + ci
			'WScript.Echo "1tot:"&tot
			'---------------------------
			c = Mid(nome_colonna,2,1)
			ci = CInt(Asc(UCase(c))-64)
			ci = ci * 26
			tot = tot + ci
			'WScript.Echo "2tot:"&tot
			'---------------------------
			c = Mid(nome_colonna,3,1)
			ci = CInt(Asc(UCase(c))-64)
			tot = tot + ci
			'WScript.Echo "3tot:"&tot
		else
			c = Mid(nome_colonna,1,1)
			ci = CInt(Asc(UCase(c))-64)
			ci = ci * 26
			tot = tot + ci
			'---------------------------
			c = Mid(nome_colonna,2,1)
			ci = CInt(Asc(UCase(c))-64)
			tot = tot + ci
		end If
	Else
		if (nome_colonna = "") then
			On Error Resume Next
			objStdOut.Write "<font color ='red'>Errore Apertura Nome colonna vuoto"&"</font>"&vbCrLf
			on error goto 0
			'objShell.popup 13,NomeProgramma, "Errore Apertura Nome colonna vuoto"
			Err.Clear
		end if
		tot = CInt(Asc(UCase(nome_colonna))-64)
		end if
		calcolaColonna = tot
End Function

Function popolaListaFile(ByRef lf, ByVal ind) 'Crea la lista dei file di input per evitare di leggerli due volte
Dim x, y, ce
'On Error Resume Next
y = 0
	for x = 0 to ind-1 step 1
		ce = thereis(listahyp(5,x),lf,y)
		if Not (ce) then
			lf(y) = listahyp(5,x)
			y = y+1
		end if
	Next
	popolaListaFile = y ' esporto il livello al quale è arrivata la listaFile
end Function

Function thereis(ByVal ff, ByRef lis, ByVal upto)
dim i, ce_dir
'On Error Resume Next
	for i = 0 to upto step 1
		if (ff = lis(i)) then 'se la dir c'è
			thereis = true
			exit function
		end if
	Next
	thereis = false
end function

Function SeparaRigheColonne(ByRef indirizzo, ByRef riga, ByRef colonna)
Dim c, i
'On Error Resume Next
	For i=1 To Len(indirizzo)
		c = Mid(indirizzo,i,1)
		if (IsNumeric(c)) then
			Exit For
		End If
	Next 
	'WScript.Echo "Numerico da " & i
	colonna = Mid(indirizzo, 1, i-1)
	riga = Mid(indirizzo, i, Len(indirizzo))
End Function

Function cercaFile(ByVal patt, ByVal folder, ByVal fso) 'pattern da cercare e directory
Dim filenamecompleto, list
Dim f, parte
Dim objFolder
'On Error Resume Next
	filenamecompleto = "NULLA"
    Set list = CreateObject("ADOR.Recordset")
    list.Fields.Append "name", 200, 255
    list.Fields.Append "date", 7
    list.Open

    'list.MoveFirst
    'Do Until list.EOF
    '  WScript.Echo list("date").Value & vbTab & list("name").Value
    '  list.MoveNext
    'Loop

	'WScript.Echo "Folder in cerca:" & folder
	Set objFolder  = fso.GetFolder(folder)
	patt = LCase(patt)

    For Each f In objFolder.Files
      list.AddNew
      list("name").Value = f.Name
      list("date").Value = f.DateLastModified
      list.Update
    Next
    list.Sort = "date DESC"
    'list.Sort = "date ASC"
    list.MoveFirst
    Do Until list.EOF
		parte = Left(LCase(list("name").Value),Len(patt)) 'preleva i primi caratteri
		if Not (InStr(parte,patt) = 0) Then
			filenamecompleto = LCase(list("name").Value)
            exit do
		End If
        list.MoveNext
    Loop
	cercaFile = filenamecompleto
End Function

function estraiFolderDaAddress(ByVal add,ByVal iniz)
Dim Pos,out,piece,id_n,i, temp, temp2, sotto_dir
'On Error Resume Next
	if (Len(add) = 0) then
		estraiFolderDaAddress = "LinkSuSeStesso"
		exit function
	end if
    if ((Instr(1,add,":\") > 0) or (Instr(1,add,"\\") > 0)) then
        'dir completa
        Pos = Instr(1,add,iniz)
        estraiFolderDaAddress = Mid(add,1,Pos-2) 'toglie anche la \
    else
        if (Instr(1,add,"..") > 0) then
            Pos = Instr(1,add,"\")
            if (Pos = 0) then
               Pos = Instr(1,add,"/")
            end if
            temp = Mid(add,Pos+1,Len(add)-Pos)
            sotto_dir = 0
            while (Instr(1,temp,"..") > 0)
                Pos = Instr(1,temp,"\")
                if (Pos = 0) then
                    Pos = Instr(1,temp,"/")
                end if
                temp = Mid(temp,Pos+1,Len(temp)-Pos)
                sotto_dir = sotto_dir+1
            Wend
            Pos = Instr(1,temp,iniz)
            temp = Mid(temp,1,Pos - 2) 'toglie anche la \
            temp2 = folderMaster
            for i=0 to sotto_dir
                Pos = InStrRev(temp2,"\")
                if (Pos = 0) then
                    Pos = InStrRev(temp2,"/")
                end if
                temp2 = Mid(temp2,1,Pos-1)
            Next
            estraiFolderDaAddress = temp2&"\"&temp
        else
            if Not (Instr(1,add,"\") > 0) then
            'contiene solo il nome file la dir è quella del Master
                estraiFolderDaAddress = folderMaster
            end if
        end if
    end if
end function

Function scriviSu(ByVal nome, ByVal dato)
Dim objFileToWrite
'On Error Resume Next
	'WScript.Echo "File di scrittura:"&nome
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(nome,2,true)
	objFileToWrite.WriteLine(dato)
	objFileToWrite.Close
	Set objFileToWrite = Nothing

End Function

Function appendiA(ByVal nome, ByVal dato)
Dim objFileToWrite
'On Error Resume Next
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(nome,8,true)
	objFileToWrite.WriteLine(dato)
	objFileToWrite.Close
	Set objFileToWrite = Nothing

End Function

Function scriviLinkTo(objSheet, lis, idx, chiave)
Dim cartella, file, commento, comm, objCommento
'On Error Resume Next
'DAFARE Devo poi vedere di accodarmi ad un commento
    if (StrComp(lis(5,idx),"LinkSuSeStesso") = 0) then
        cartella = "LinkSuSeStesso"
        chiave = "XL" 'Il controllo viene fatto soltanto sulla seconda lettera che sia TL o FL fa poca differenza
    else
	    cartella = Mid(lis(5,idx),1,InStrRev(lis(5,idx),"\")-1)
    end if
	On Error Resume Next
    Set objSheet = objWorkbook.Worksheets(lis(0,idx)) 'mi setto sul giusto sheet
    If (Err.Number <> 0) Then
		objStdOut.Write "<font color ='red'>Errore select sheet scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		'objShell.popup "Errore select sheet scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		Err.Clear
	End If
	on error goto 0
    '                  (riga,col)
	'objSheet.Cells(lis(1,idx), lis(2,idx)).ClearComments
	On Error Resume Next
    comm = objSheet.Cells(lis(1,idx), lis(2,idx)).Comment.Text
	If (Err.Number <> 0) Then
		objStdOut.Write "<font color ='red'>Errore clear comment scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		'objShell.popup "Errore clear comment scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		Err.Clear
	End If
	on error goto 0
    if (Instr(1,comm,"HyFlink#") > 0) then 'c'è già un link devo toglierlo
        'rimuovo il link
        commento = rimuoviLink(comm)
	else
		commento = comm
    end if
	if (StrComp(lis(5,idx),"LinkSuSeStesso") = 0) then
		'HyFlink#XL#Pos=S=xx#C=jj#R=rr#Punt=S=xx#C=jj#R=rr#HyElink
		commento = commento & vbCrLf&"HyFlink#"&chiave&"#Pos=PS="&lis(0,idx)&"#PC="&lis(3,idx)&"#PR="&lis(1,idx)&"#Punt=S="&lis(7,idx)&"#C="&lis(8,idx)&"#R="&lis(10,idx)&"#HyElink"
	else
		commento = commento & vbCrLf&"HyFlink#"&chiave&"#Pos=PS="&lis(0,idx)&"#PC="&lis(3,idx)&"#PR="&lis(1,idx)&"#cartella="&cartella&"#file="&lis(11,idx)&"#S="&lis(7,idx)&"#C="&lis(8,idx)&"#R="&lis(10,idx)&"#HyElink"
	end if
	commento = commento & vbCrLf&"FcomTime#"&Date&" "&Hour(Now())&":"&Minute(Now())&":"&Second(Now())&"#FcomTime"
	On Error Resume Next
    objSheet.Cells(lis(1,idx), lis(2,idx)).ClearComments
    If (Err.Number <> 0) Then
		objStdOut.Write "<font color ='red'>Errore clear comment scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		'objShell.popup "Errore clear comment scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		Err.Clear
	End If
	on error goto 0
	On Error Resume Next
	objSheet.Cells(lis(1,idx), lis(2,idx)).AddComment commento
	If (Err.Number <> 0) Then
		objStdOut.Write "<font color ='red'>Errore add comment scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
		'objShell.popup "Errore add comment scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
		Err.Clear
	End If
	on error goto 0
	'Set objCommento = objSheet.Cells(lis(1,idx), lis(2,idx)).Comment
	'If (Err.Number <> 0) Then
	'	On Error Resume Next
	'	objStdOut.Write "<font color ='red'>Errore get obj Commento scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096&"</font>"&vbCrLf
	'	on error goto 0
	'	'objShell.popup "Errore get obj Commento scriviLinkTo:" & Err.Number & " Description " & Err.Description , AttesaMessaggioLLL, "Errore", CRITICAL_ICON + 4096
	'	Err.Clear
	'End If
	'objCommento.Shape.TextFrame.Characters(1, l1).Font.Size = 16
	'objCommento.Shape.TextFrame.Characters(l1+1, l2+l3).Font.Bold = false
	'objCommento.Shape.TextFrame.Characters(l1+l2+1, l3).Font.Size = 14
	'objCommento.Shape.TextFrame.Characters(l1+l2+1, l3).Font.Color = RGB(0,255,0)
	scrivilinkTo = true
end function

function rimuoviLink(ByRef cc)
Dim PosI,PosF,tmp, tmp1, temporaneo
    PosI = InStr(1,cc,"HyFlink#",1)
    PosF = InStr(1,cc,"#FcomTime",1) +8
    if ((PosF = Len(cc)) and PosI = 1) then
        rimuoviLink = ""
        exit function
    else
        tmp = Mid(cc,1,PosI-1)
        tmp1 = Mid(cc,PosF+1,Len(cc)-PosF)
    end if
    temporaneo = rimuoviLineeVuoteIniziali(tmp & tmp1)
    rimuoviLink = temporaneo
end function

function rimuoviLineeVuoteIniziali(ByVal str)
    Do
	    if (InStr(1,str,VbCrlf) = 1) then
            str = Mid(str,3,Len(str)-2)
        end if
    Loop While (InStr(1,str,VbCrlf) = 1)
    rimuoviLineeVuoteIniziali = str
end function
