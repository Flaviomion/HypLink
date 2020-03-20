Option Explicit On

Imports System
Imports System.IO
Module Hyplink_m1

    Sub Main()
        Const Versione = "1.8.0"
        On Error GoTo 0
        Dim dati_versione
        Console.ForegroundColor = ConsoleColor.Yellow
        dati_versione = "dal 1.3.0 in avanti ho rivoluzionato i Flink per tutti i link (non riuscivo senza le informazioni da dove veniva il link a ricostruire i link spostati)" & vbCrLf
        dati_versione = dati_versione & "ATTENZIONE questo succede quando uno spostamento va a posizionarsi proprio dove c'èra un altro hyperlink." & vbCrLf
        dati_versione = dati_versione & "ora i link saranno del tipo locali		HyFlink#XL#Pos=PS=xx#PC=jj#PR=rr#Punt=S=xx#C=jj#R=rr#HyElink" & vbCrLf
        dati_versione = dati_versione & "Remoti		HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink" & vbCrLf
        dati_versione = dati_versione & "Remoti		HyFlink#FR#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink" & vbCrLf
        dati_versione = dati_versione & "Con la 1.4.0 sfrutto in nuovi Flink per gestire anche movimenti contemporanei del TO e FR." & vbCrLf
        dati_versione = dati_versione & "1.4.1 gestisce il path dei nomi dei file Linked come relativi alla foldeMaster" & vbCrLf
        dati_versione = dati_versione & "1.5.0 soluzione errori 1.4.1" & vbCrLf
        dati_versione = dati_versione & "1.5.1 salta i link hyplink sconosciuti" & vbCrLf
        dati_versione = dati_versione & "1.5.2 gestisce slash opposta al posto della normale in scriviLinkTo" & vbCrLf
        dati_versione = dati_versione & "1.5.3 evita che il programma venga lanciato direttamente con doppio click sul .vbs" & vbCrLf
        dati_versione = dati_versione & "1.5.4 Corregge gli hyperlink anche sui file linkati e inserisce dati versione" & vbCrLf
        dati_versione = dati_versione & "1.5.5 Inserita richiesta Password (jVB_pass.jar) e gestione password" & vbCrLf
        dati_versione = dati_versione & "1.5.6 Inserito output specifico su errore chiamata cercaFlinkDaHyp" & vbCrLf
        dati_versione = dati_versione & "1.5.7 migliorato output errore di cercaFlinkDaHyp e tolto errore 424 Necessario Oggetto su commento inesistente" & vbCrLf
        dati_versione = dati_versione & "1.5.8 tolto errore chiamata a cercaFlinkDaHyp (linea circa 615)" & vbCrLf
        dati_versione = dati_versione & "1.5.9 modifica a modificaHypTo per potermodificare Hyperlink quando c'è uno spostamento sul file linked (riga 1344 circa)" & vbCrLf
        dati_versione = dati_versione & "1.6.0 Intermedio con errori" & vbCrLf
        dati_versione = dati_versione & "1.7.0 Revisione logica e ricerca Link Persi, modificato modificaLinkTo dove ho scritto ATTENZIONE" & vbCrLf
        dati_versione = dati_versione & "1.8.0 Estrae anche gli LinkedHyp, li incrocia e aggiunge al controllo il controllo che tutti gli hyp abbiano una controparte LinkedHyp" & vbCrLf
        dati_versione = dati_versione & ""
        Dim myPass = ""

        'Solo per debug
        'myPass = "segreto"

        Const Name = "HypLink"

        Dim sheet, objExcel, objWorkbook
        Dim objWorkSecondbook = Nothing
        Dim objWorksheet = Nothing
        Dim objWorkSecondsheet = Nothing
        Dim objRange, objShell, objStdOut
        Dim out = ""
        Dim report
        Dim y, z, CDir, indice, indice_hyp, indice_lhyp, indice_Lflink, Uscita, sheet_n, Errore
        Dim n_fileIn = 0
        Dim ris, idx
        Dim folderMaster, file_master, fileMaster, fileINI, hyp
        Dim sheet_name
        Dim aggiuntaLink
        Dim scrittoLinkNuovi = False
        Dim hyp_link_new
        Dim ptx
        Dim foglio_s, col_s, riga_s
        Dim NomeProgramma
        Dim Fdebug, bReadOnly
        Dim settaFont, dimFont
        Dim Started
        Dim Att = False
        Dim mess
        Fdebug = True
        NomeProgramma = Name & "_" & Versione
        fileINI = "HypLink.ini"

        Dim ofile As TextWriter = File.CreateText("Output.txt")
        Dim oOut As TextWriter = Console.Out
        Dim width = Console.WindowWidth
        Dim heigth = Console.WindowHeight
        Console.SetWindowSize(width * 2, heigth * 2)
        WriteMia(ConsoleColor.White, "Partenza " & NomeProgramma, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Yellow
        'Console.WriteLine("Partenza " & NomeProgramma) 'pass:"&myPass&""

        Dim objArgs() As String = Environment.GetCommandLineArgs()
        Console.BackgroundColor = ConsoleColor.Black
        If (objArgs.Count > 1) Then
            If (StrComp(objArgs(1), "versione") = 0) Then
                versioneDisp(ofile, oOut, Versione, dati_versione)
            Else
                myPass = objArgs(1)
            End If
        Else
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.Write("Dammi la Password:")
            Console.ForegroundColor = ConsoleColor.Black
            myPass = Console.ReadLine()
            If (myPass = "versione") Then
                versioneDisp(ofile, oOut, Versione, dati_versione)
            End If
        End If

        'folderMaster = WScript.Arguments.Item(0)
        'file_master = WScript.Arguments.Item(1)
        Dim listahyp(17, 500)      ' 0	sheet (dove è registrato il link)
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
        ' 15   Puntatore al corrispondente hyp o LinkedHyp
        ' 16   indica link a se stesso o se -1 link non completamente realizzato (spostato)

        Dim listaLinkedhyp(17, 500) ' 0	sheet (dove è registrato il link)
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
        ' 15   Puntatore al corrispondente hyp o LinkedHyp
        ' 16   indica link a se stesso o se -1 link non completamente realizzato (spostato)


        Dim listaFlink(17, 500)    ' 0	sheetLocale
        ' 1	rigaLocale
        ' 2	colLocNum
        ' 3	colonnaLocale
        ' 4    cartellaLinked (Rel)
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


        Dim listaLinkedFlink(17, 500)    ' 0	sheetLocale
        ' 1	rigaLocale
        ' 2	colLocNum
        ' 3	colonnaLocale
        ' 4    cartellaLinked (Rel)
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

        Dim listaFile(500)  '0	File path completo

        Dim listaSegnalazioni(100)  ' File segnalato


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


        Dim fso : fso = CreateObject("Scripting.FileSystemObject")
        CDir = fso.GetAbsolutePathName(".")

        inilistaSegnalazioni(listaSegnalazioni)
        objShell = CreateObject("Wscript.Shell")
        objExcel = CreateObject("Excel.Application")
        aggiuntaLink = False
        folderMaster = ""
        fileMaster = ""
        report = ""
        settaFont = ""
        dimFont = 0
        leggiINI(folderMaster, fileMaster, report, Fdebug, settaFont, dimFont, fileINI)
        If (Fdebug) Then
            mess = "Debug Attivo"
            WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
        End If
        objExcel.DisplayAlerts = 0
        If (Fdebug) Then
            objExcel.Visible = True
        End If
        file_master = cercaFile(fileMaster, folderMaster, fso)

        'ris = objShell.popup("Elaboro il file:" & folderMaster & "\" & file_master, 1, "Info", INFO_ICON_YN + 4096)
        'If (ris = 7) Then
        'Console.ForegroundColor = ConsoleColor.Magenta
        'Console.WriteLine("Programma fermato dall'utente")
        'Console.ForegroundColor = ConsoleColor.White
        'End
        'End If
        WriteMia(ConsoleColor.White, "Carico il file: " & folderMaster & "\" & file_master, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.White
        'Console.WriteLine("Carico il file: " & folderMaster & "\" & file_master)


        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Prima di Apertura File Attenzione" & Att)

        Uscita = "<html><head><meta charset='utf-8' /><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /></head>"

        Uscita = Uscita & "<body><table><tr><td colspan=10><h2><font color='blue'>Ripristina Hyperlink versione " & Versione & "</font></h2></td></tr>"
        Uscita = Uscita & "<tr><td colspan=10><font color ='DarkGreen'>riposiziona collegamenti Ipertestuali</font></td></tr>"
        Uscita = Uscita & "<tr><td align='center' colspan=10><font color ='DarkBlue'>" & folderMaster & "\" & file_master & "</font></td></tr>"
        Uscita = Uscita & "<tr><td colspan=10><font color ='blue'>Rapporto del " & Now & " " & Hour(Now()) & ":" & Minute(Now()) & "</font></td></tr>"
        Uscita = Uscita & "<tr></tr>"
        On Error Resume Next
        objWorkbook = objExcel.Workbooks.Open(folderMaster & "\" & file_master, False, False)
        If (Err.Number <> 0) Then
            WriteMia(ConsoleColor.Red, "Errore apertura file Master " & folderMaster & "\" & file_master & "</font><br/> Descrizione " & Err.Description, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore apertura file Master " & folderMaster & "\" & file_master & "</font><br/> Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0

        bReadOnly = objWorkbook.ReadOnly
        If bReadOnly = True Then
            WriteMia(ConsoleColor.Red, "Errore apertura file " & folderMaster & "\" & file_master & " File OCCUPATO</font><br/>" & Err.Description, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore apertura file " & folderMaster & "\" & file_master & " File OCCUPATO</font><br/>")
            Call objWorkbook.Close
            objExcel.Quit
            mess = "Programma Terminato a causa di file Occupato"
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Programma Terminato a causa di file Occupato")
            mess = "Enter per Terminare"
            WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Yellow
            'Console.WriteLine("Enter per Terminare")
            Console.ForegroundColor = ConsoleColor.White
            Console.ReadLine()
            ofile.Close()
            End
        End If

        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Prima di SetPassword Attenzione" & Att)

        settaPassword(objWorkbook, myPass, Att, ofile, oOut, Fdebug)

        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Dopo SetPassword Attenzione" & Att)

        sheet_n = objWorkbook.Sheets.Count

        indice = 0 'indice viene incrementato da FindTO
        indice_hyp = 0 'indice_hyp viene incrementato da FindHyper
        indice_lhyp = 0 'indice_hyp viene incrementato da FindHyper
        indice_Lflink = 0 'indice_hyp viene incrementato da 

        'HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink
        'Primo giro, raccolgo hyplink e Flink

        mess = "Primo giro, raccolgo hyplink e Flink TO e TL e FL"
        WriteMia(ConsoleColor.White, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.White
        'Console.WriteLine("Primo giro, raccolgo hyplink e Flink TO e TL e FL")
        'End '-------------------per test


        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)


        popolaListe(indice, n_fileIn, listaFile, listaFlink, indice_hyp, listahyp, indice_lhyp, listaLinkedhyp, indice_Lflink, listaLinkedFlink, sheet_n,
                    objWorksheet, objWorkbook,
                    objWorkSecondbook, objWorkSecondsheet, Att, folderMaster, file_master, objExcel, myPass,
                    listaSegnalazioni, Uscita, fso, fileMaster, ofile, oOut, Fdebug) 'Questa prima volta i LinkedHyp e i LinkedFlink non li raccoglie perchè non è ancora stata definita la lista dei file linked

        Uscita = Uscita & "<tr><td align='center' colspan='10'><font color ='DarkBlue'>Link attivi</font></td></tr>"
        For y = 0 To indice_hyp - 1
            Uscita = Uscita & "<tr><td><font color ='green'>Link " & y & ") </font></td><td><font color ='green'>" & folderMaster & "\" & file_master & "</font></td><td><font color ='green'> di:" & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y) & "</font></td><td><font color ='green'> a:" & listahyp(5, y) & " " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y) & "</font></td></tr>"
            mess = "Llink " & folderMaster & "\" & file_master & " di:" & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y) & " a:" & listahyp(5, y) & " " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y)
            WriteMia(ConsoleColor.Green, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Green
            'Console.WriteLine("Llink " & folderMaster & "\" & file_master & " di:" & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y) & " a:" & listahyp(5, y) & " " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y))
        Next

        'scriviSu "ListaPI.txt", "+++++++++++++++Liste dopo primo incrocio++++++++++++" &vbCrLf
        'appendiA "ListaPI.txt", outListe
        'SOLO PER TEST
        'Call objWorkbook.Save
        'Call objWorkbook.Close
        'objExcel.Quit
        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Processo Concluso per TEST analisi primo giro/incrocio"&vbCrLf&vbCrLf
        'Wscript.Quit 0
        'SOLO PER TEST

        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)

        n_fileIn = popolaListaFile(listaFile, indice_hyp, listahyp) 'Crea la lista dei file linked da listahyp con nome gia corretto
        mess = "Creo i Flink Locali mancanti"
        WriteMia(ConsoleColor.White, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.White
        'Console.WriteLine("Creo i Flink Locali mancanti")

        creaFlinkLocali(indice_hyp, listahyp, scrittoLinkNuovi, Uscita, Att, objWorksheet, objWorkbook, folderMaster,
                        file_master, aggiuntaLink, ofile, oOut, Fdebug) 'Crea i HyFlink#XL su link locali

        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)

        'Crea i HyFlink#TO e HyFlink#FR su quelli che puntano all'esterno
        mess = "Creo i Flink Remoti mancanti"
        WriteMia(ConsoleColor.White, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.White
        'Console.WriteLine("Creo i Flink Remoti mancanti")
        creaFlinkMancanti(indice_hyp, listahyp, scrittoLinkNuovi, Uscita, objWorkbook, objWorksheet, Att, folderMaster, file_master, aggiuntaLink,
                          objWorkSecondbook, objWorkSecondsheet, objExcel, myPass, fileMaster, ofile, oOut, Fdebug)
        If (Fdebug) Then
            mess = "Debug: Popolo le liste"
            WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.White
            'Console.WriteLine("Debug: Popolo le liste")
        End If
        popolaListe(indice, n_fileIn, listaFile, listaFlink, indice_hyp, listahyp, indice_lhyp, listaLinkedhyp, indice_Lflink, listaLinkedFlink, sheet_n, objWorksheet, objWorkbook,
                    objWorkSecondbook, objWorkSecondsheet, Att, folderMaster, file_master, objExcel, myPass,
                    listaSegnalazioni, Uscita, fso, fileMaster, ofile, oOut, Fdebug) 'qyuesta vota popola tutto
        If (Fdebug) Then
            mess = "Debug: Incrocio link dopo eventuale Aggiunta"
            WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.White
            'Console.WriteLine("Debug: Incrocio link dopo eventuale Aggiunta")
        End If


        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)

        CambiaFile(objWorkbook, sheet_n, Att, Uscita, listaFlink, indice, listaLinkedFlink, indice_Lflink, Fdebug, file_master, fileMaster, ofile, oOut) 'Rimette a posto gli hyperlink del file principale verso i file Linked con il giusto nome File
        cambiaFileNameInLinked(n_fileIn, listaFile, objWorkbook, objWorkSecondbook, objExcel, Att, myPass, Uscita, listaFlink, indice,
                               listaLinkedFlink, indice_Lflink, Fdebug, file_master, fileMaster, ofile, oOut) 'Rimette a posto gli hyperlink nei files linked verso il file master
        If (Fdebug) Then
            mess = "Debug: Popolo le liste"
            WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.White
            'Console.WriteLine("Debug: Popolo le liste")
        End If
        popolaListe(indice, n_fileIn, listaFile, listaFlink, indice_hyp, listahyp, indice_lhyp, listaLinkedhyp, indice_Lflink, listaLinkedFlink, sheet_n, objWorksheet, objWorkbook,
                    objWorkSecondbook, objWorkSecondsheet, Att, folderMaster, file_master, objExcel, myPass,
                    listaSegnalazioni, Uscita, fso, fileMaster, ofile, oOut, Fdebug)

        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)

        If (Fdebug) Then
            mess = "Debug: Incrocio link dopo eventuale Cambio file"
            WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.White
            'Console.WriteLine("Debug: Incrocio link dopo eventuale Cambio file")
        End If
        mess = "Controllo e riposiziono i Link"
        WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Controllo e riposiziono i Link")
        controlloLink(indice, listaFlink, objWorksheet, listaLinkedFlink, indice_Lflink, Att, folderMaster, file_master,
                      objWorkSecondbook, objWorkSecondsheet, objWorkbook, objExcel, myPass, Uscita, fileMaster, ofile, oOut, Fdebug)
        If (Fdebug) Then
            mess = "Debug: Popolo le liste dopo Controllo Link"
            WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.White
            'Console.WriteLine("Debug: Popolo le liste dopo Controllo Link")
        End If
        popolaListe(indice, n_fileIn, listaFile, listaFlink, indice_hyp, listahyp,
                          indice_lhyp, listaLinkedhyp, indice_Lflink, listaLinkedFlink, sheet_n,
                          objWorksheet, objWorkbook, objWorkSecondbook, objWorkSecondsheet,
                          Att, folderMaster, file_master, objExcel, myPass, listaSegnalazioni, Uscita, fso, fileMaster, ofile, oOut, Fdebug)
        mess = "Controllo che tutti gli HyperLink(Microsoft) abbiano il Corrispondente"
        WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Controllo che tutti gli HyperLink(Microsoft) abbiano il Corrispondente")

        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)

        If Not (controlloHyp(indice_hyp, listahyp, indice_lhyp, listaLinkedhyp, ofile, oOut)) Then
            Call objWorkbook.Save
            Call objWorkbook.Close
            objExcel.Quit
            mess = "Processo Concluso per Controllo HyperLink con Link guasti vedi sopra"
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Processo Concluso per Controllo HyperLink con Link guasti vedi sopra")
            mess = "Enter per Terminare"
            WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Yellow
            'Console.WriteLine("Enter per Terminare")
            Console.ForegroundColor = ConsoleColor.White
            Console.ReadLine()
            ofile.Close()
            End
        End If

        scriviSu("ListaDC.txt", "+++++++++++++++Liste Dopo controlloLink++++++++++++" & vbCrLf)
        appendiA("ListaDC.txt", outListe(out, indice_hyp, indice_lhyp, indice, listahyp,
                      listaLinkedhyp, listaFlink, listaLinkedFlink, indice_Lflink, n_fileIn, listaFile))

        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)

        Dim Persi
        Persi = controlloLinkPersi(indice_hyp, listahyp, indice, listaFlink, indice_Lflink, listaLinkedFlink, ofile, oOut)

        If (Persi) Then
            Call objWorkbook.Save
            Call objWorkbook.Close
            objExcel.Quit
            mess = "Processo Concluso per Controllo Link Persi"
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Processo Concluso per Controllo Link Persi")
            mess = "Enter per Terminare"
            WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Yellow
            'Console.WriteLine("Enter per Terminare")
            Console.ForegroundColor = ConsoleColor.White
            Console.ReadLine()
            ofile.Close()
            End
        End If


        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)


        If (settaFont) Then
            If ((CInt(dimFont) > 0) And (CInt(dimFont) < 60)) Then
                settaFontPerCommenti(objWorkbook, dimFont, Att, ofile, oOut, Fdebug)
            End If
        End If

        Call objWorkbook.Save
        Call objWorkbook.Close
        'SOLO PER TEST
        'objExcel.Quit
        'Console.WriteLine("Processo Concluso per TEST analisi primo giro/incrocio"&vbCrLf&vbCrLf
        'Wscript.Quit 0
        'SOLO PER TEST

        ' Setto Font Commenti su file linked
        For y = 0 To n_fileIn - 1
            If Not (InStr(1, listaFile(y), "LinkSuSeStesso") > 0) Then
                On Error Resume Next
                objWorkbook = objExcel.Workbooks.Open(listaFile(y), False, False)
                If (Err.Number <> 0) Then
                    mess = "Errore nell'Apertura di " & objWorkbook.Name & " " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore nell'Apertura di " & objWorkbook.Name & " " & Err.Description)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
                settaPassword(objWorkbook, myPass, Att, ofile, oOut, Fdebug)
                settaFontPerCommenti(objWorkbook, dimFont, Att, ofile, oOut, Fdebug)
                Call objWorkbook.Save
                Call objWorkbook.Close
            End If
        Next


        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine("Attenzione" & Att)

        objExcel.Quit

        If Not (Att) Then
            mess = "Programma Terminato con Successo"
            WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.White
            'Console.WriteLine("Programma Terminato con Successo")
        Else
            mess = "Programma Terminato con alcune attenzioni"
            WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Magenta
            'Console.WriteLine("Programma Terminato con alcune attenzioni")
        End If
        Uscita = Uscita & "<table><body><html>"
        scriviSu(report, Uscita)
        'objShell.run report 'Lancia l'eseguibile definito per il tipo di file da leggere.
        mess = vbCrLf & "-------------------------Fine---------------------------------" & vbCrLf
        WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Cyan
        'Console.WriteLine(vbCrLf & "-------------------------Fine---------------------------------" & vbCrLf)
        mess = "Enter per Terminare"
        WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Yellow
        'Console.WriteLine("Enter per Terminare")
        Console.ForegroundColor = ConsoleColor.White
        Console.ReadLine()
        ofile.Close()
        End

    End Sub

    Function versioneDisp(ByRef ofile, ByRef oOut, Versione, dati_versione)
        WriteMia(ConsoleColor.Cyan, Versione, oOut, ofile)
        WriteMia(ConsoleColor.Cyan, dati_versione, oOut, ofile)
        Dim dummy = Console.ReadLine()
        ofile.Close()
        End
    End Function

    Function WriteMia(colore, messaggio, ByRef oOut, ByRef ofile)
        Console.ForegroundColor = colore
        Console.SetOut(oOut)
        Console.WriteLine(messaggio)
        Console.SetOut(ofile)
        Console.WriteLine(messaggio)
        Console.SetOut(oOut)
        WriteMia = 0
    End Function



    ' -----Fine Programma Inizio Funzioni ----------------------------------------------------------------------------------------------
    Function cambiaFileNameInLinked(ByRef n_fileIn, ByRef listaFile, ByRef objWorkbook, ByRef objWorkSecondbook, ByRef objExcel, ByRef Att, ByRef myPass,
                                    ByRef Uscita, ByRef ListaFlink, ByRef indice, ByRef listaLinkedFlink, ByRef indice_Lflink, ByRef Fdebug,
                                    ByRef file_master, ByRef fileMaster, ByRef ofile, ByRef oOut)
        Dim bReadOnly
        'vado a Rimettere a posto il nome del file master negli hyperlink nei file Linked 
        For y = 0 To n_fileIn - 1
            If (InStr(1, listaFile(y), "LinkSuSeStesso") > 0) Then 'Il link è locale sullo stesso file
                'Non faccio niente i link Tl e FL sono già raccolti
            Else
                On Error Resume Next
                objWorkSecondbook = objExcel.Workbooks.Open(listaFile(y), False, False)
                If (Err.Number <> 0) Then
                    Dim mess = "Errore nell'apertura file Linked Descrizione " & Err.Description & vbCrLf
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore nell'apertura file Linked Descrizione " & Err.Description & vbCrLf)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
                bReadOnly = objWorkSecondbook.ReadOnly
                If bReadOnly = True Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Errore apertura file " & listaFile(y) & ", File OCCUPATO")
                    Call objWorkSecondbook.Close
                    Call objWorkbook.Close
                    objExcel.Quit
                    Dim mess = "Programma Terminato a causa di file " & listaFile(y) & " Occupato"
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Programma Terminato a causa di file " & listaFile(y) & " Occupato")
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine("Enter per Terminare")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.ReadLine()
                    ofile.Close()
                    End
                End If
                settaPassword(objWorkSecondbook, myPass, Att, ofile, oOut, Fdebug)
                If (Err.Number <> 0) Then
                    Dim mess = "Errore nel settare la Password su " & listaFile(y) & " " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore nel settare la Password su " & listaFile(y) & " " & Err.Description)
                    Err.Clear()
                End If
                On Error GoTo 0
                CambiaFile(objWorkSecondbook, objWorkSecondbook.Sheets.Count, Att, Uscita, ListaFlink, indice, listaLinkedFlink,
                           indice_Lflink, Fdebug, file_master, fileMaster, ofile, oOut)
                objWorkSecondbook.Save
                objWorkSecondbook.Close
            End If
        Next
    End Function

    Function creaFlinkMancanti(ByRef indice_hyp, ByRef listahyp, ByRef scrittoLinkNuovi, ByRef Uscita, ByRef objWorkbook, ByRef objWorksheet, ByRef Att, ByVal folderMaster, ByVal file_master, ByRef aggiuntaLink,
                               ByRef objWorkSecondbook, ByRef objWorkSecondsheet, ByRef objExcel, ByRef myPass, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim ris As Integer
        Dim mess
        For y = 0 To indice_hyp - 1
            If (listahyp(14, y) = -1) Then
                If (listahyp(12, y) = -1) Then 'non c'è ancora l'HyFlink#TO# e quindi neanche HyFlink#FR sul file linkato#
                    If Not (scrittoLinkNuovi) Then
                        Uscita = Uscita & "<tr><td align='center' colspan='10'><font color ='Navy'>Link nuovi</font></td></tr>"
                        scrittoLinkNuovi = True
                    End If
                    On Error Resume Next
                    objWorksheet = objWorkbook.Worksheets(listahyp(0, y))
                    If (Err.Number <> 0) Then
                        mess = "Errore creazione oggetto sheet Master</font><br/> Descrizione " & Err.Description
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore creazione oggetto sheet Master</font><br/> Descrizione " & Err.Description)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    Console.WriteLine("<font color='orange'>Creazione Flink " & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y) & " -> " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y))
                    ris = scriviLinkTo(objWorksheet, listahyp, y, "TO", objWorkbook, Att, ofile, oOut, Fdebug)
                    If Not (ris) Then
                        mess = "Errore nella creazione di HyFlink su " & folderMaster & "\" & fileMaster
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore nella creazione di HyFlink su " & folderMaster & "\" & fileMaster)
                        Err.Clear()
                    End If
                    'scrivo il link HyFlink#FR# sul file linkato
                    mess = "Creazione linkedFlink " & listahyp(5, y) & " " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y) & "->" & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y)
                    WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Yellow
                    'Console.WriteLine("Creazione linkedFlink " & listahyp(5, y) & " " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y) & "->" & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y))
                    ris = scrivoLinkSuLinked(y, listahyp, True, "FR", objWorksheet, objWorkbook, Att, objWorkSecondbook, objWorkSecondsheet,
                                             objExcel, myPass, folderMaster, file_master, fileMaster, ofile, oOut, Fdebug)
                    If Not (ris) Then
                        Console.WriteLine("Errore nella creazione di HyFlink#FR# su " & listahyp(5, y))
                        Err.Clear()
                    End If
                    On Error GoTo 0
                    Uscita = Uscita & "<tr><td><font color ='maroon'>Aggiunto Link</font></td><td><font color ='maroon'>" & folderMaster & "\" & fileMaster & "</font></td><td><font color ='maroon'> di:" & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y) & "</font></td><td><font color ='maroon'> a:" & listahyp(5, y) & " " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y) & "</font></td></tr>"
                    aggiuntaLink = True
                End If
            End If
        Next
    End Function

    Function controlloHyp(ByRef indice_hyp, ByRef listahyp, ByRef indice_lhyp, ByRef listaLinkedhyp, ByRef ofile, ByRef oOut)
        controlloHyp = True
        For y = 0 To indice_hyp
            If (listahyp(15, y) = -1) Then 'so' cazzi questo hyper link non ha il corrispondente
                Dim mess = "L'HypLink " & listahyp(0, y) & ":" & listahyp(3, y) & listahyp(1, y) & "->" & listahyp(6, y) & " Non ha il ritorno"
                WriteMia(ConsoleColor.DarkYellow, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.DarkYellow
                'Console.WriteLine("L'HypLink " & listahyp(0, y) & ":" & listahyp(3, y) & listahyp(1, y) & "->" & listahyp(6, y) & " Non ha il ritorno")
                controlloHyp = False
            End If
        Next
        For y = 0 To indice_lhyp
            If (listaLinkedhyp(15, y) = -1) Then 'so' cazzi questo hyper link non ha il corrispondente
                Dim mess = "L'HypLink " & listaLinkedhyp(0, y) & ":" & listaLinkedhyp(3, y) & listaLinkedhyp(1, y) & "->" & listaLinkedhyp(6, y) & " Non ha il corrispettivo "
                WriteMia(ConsoleColor.DarkYellow, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.DarkYellow
                'Console.WriteLine("L'HypLink " & listaLinkedhyp(0, y) & ":" & listaLinkedhyp(3, y) & listaLinkedhyp(1, y) & "->" & listaLinkedhyp(6, y) & " Non ha il corrispettivo ")
                controlloHyp = False
            End If
        Next
    End Function
    Function controlloLinkPersi(ByRef indice_hyp, ByRef listahyp, ByRef indice, ByRef listaFlink, ByRef indice_Lflink, ByRef listaLinkedFlink, ByRef ofile, ByRef oOut)
        Dim mess = "Controllo Link persi"
        WriteMia(ConsoleColor.White, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.White
        'Console.WriteLine("Controllo Link persi")
        Dim z
        controlloLinkPersi = False
        For z = 0 To indice_hyp - 1
            If (((listahyp(12, z) = -1)) And (StrComp(listahyp(5, z), "LinkSuSeStesso") = -1)) Then
                mess = "L'HypLink " & listahyp(0, z) & ":" & listahyp(3, z) & listahyp(1, z) & "->" & listahyp(6, z) & " Non ha un Flink Corrispondente (controllare) "
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("L'HypLink " & listahyp(0, z) & ":" & listahyp(3, z) & listahyp(1, z) & "->" & listahyp(6, z) & " Non ha un Flink Corrispondente (controllare) " & vbCrLf)
                controlloLinkPersi = True
            End If
            If ((listahyp(13, z) = -1) And (StrComp(listahyp(5, z), "LinkSuSeStesso") = -1)) Then
                mess = "L'HypLink " & listahyp(0, z) & ":" & listahyp(3, z) & listahyp(1, z) & "->" & listahyp(6, z) & " Non ha un LinkedFlink  Corrispondente (controllare) "
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("L'HypLink " & listahyp(0, z) & ":" & listahyp(3, z) & listahyp(1, z) & "->" & listahyp(6, z) & " Non ha un LinkedFlink  Corrispondente (controllare) " & vbCrLf)
                controlloLinkPersi = True
            End If
        Next
        For z = 0 To indice - 1
            If (((listaFlink(12, z) = -1)) And (StrComp(listahyp(5, z), "LinkSuSeStesso") = -1)) Then
                mess = "Il Flink " & listaFlink(0, z) & ":" & listaFlink(3, z) & listaFlink(1, z) & "->" & listaFlink(6, z) & " Non ha un HyperLink Corrispondente (controllare) "
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Il Flink " & listaFlink(0, z) & ":" & listaFlink(3, z) & listaFlink(1, z) & "->" & listaFlink(6, z) & " Non ha un HyperLink Corrispondente (controllare) " & vbCrLf)
                controlloLinkPersi = True
            End If
            If ((listaFlink(13, z) = -1) And (StrComp(listaFlink(5, z), "LinkSuSeStesso") = -1)) Then
                mess = "Il Flink " & listaFlink(0, z) & ":" & listaFlink(3, z) & listaFlink(1, z) & "->" & listaFlink(6, z) & " Non ha un LinkedFlink  Corrispondente (controllare) "
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Il Flink " & listaFlink(0, z) & ":" & listaFlink(3, z) & listaFlink(1, z) & "->" & listaFlink(6, z) & " Non ha un LinkedFlink  Corrispondente (controllare) " & vbCrLf)
                controlloLinkPersi = True
            End If
        Next
        For z = 0 To indice_Lflink - 1
            If (((listaLinkedFlink(12, z) = -1)) And (StrComp(listaLinkedFlink(5, z), "LinkSuSeStesso") = -1)) Then
                mess = "Il LinkedFlink " & listaLinkedFlink(0, z) & ":" & listaLinkedFlink(3, z) & listaLinkedFlink(1, z) & "->" & listaLinkedFlink(6, z) & " Non ha un HyperLink Corrispondente (controllare) "
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                ' Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Il LinkedFlink " & listaLinkedFlink(0, z) & ":" & listaLinkedFlink(3, z) & listaLinkedFlink(1, z) & "->" & listaLinkedFlink(6, z) & " Non ha un HyperLink Corrispondente (controllare) " & vbCrLf)
                controlloLinkPersi = True
            End If
            If ((listaLinkedFlink(13, z) = -1) And (StrComp(listaLinkedFlink(5, z), "LinkSuSeStesso") = -1)) Then
                mess = "Il LinkedFlink " & listaLinkedFlink(0, z) & ":" & listaLinkedFlink(3, z) & listaLinkedFlink(1, z) & "->" & listaLinkedFlink(6, z) & " Non ha un Flink  Corrispondente (controllare) "
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Il LinkedFlink " & listaLinkedFlink(0, z) & ":" & listaLinkedFlink(3, z) & listaLinkedFlink(1, z) & "->" & listaLinkedFlink(6, z) & " Non ha un Flink  Corrispondente (controllare) " & vbCrLf)
                controlloLinkPersi = True
            End If
        Next
    End Function

    Function popolaListe(ByRef indice, ByRef n_fileIn, ByRef listaFile, ByRef listaFlink, ByRef indice_hyp, ByRef listahyp,
                         ByRef indice_lhyp, ByRef listaLinkedhyp, ByRef indice_Lflink, ByRef listaLinkedFlink, sheet_n,
                         ByRef objWorksheet, ByRef objWorkbook, ByRef objWorkSecondbook, ByRef objWorkSecondsheet,
                         ByRef Att, ByRef folderMaster, ByRef file_master, ByRef objExcel, ByRef myPass, ByRef listaSegnalazioni,
                         ByRef Uscita, ByRef fso, ByRef fileMaster, ofile, ByRef oOut, Fdebug) 'raccoglie tutte le liste tranne quella dei files
        Dim sheet_name As String
        If (Fdebug) Then
            Dim mess = "Ripopolo le liste"
            WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Yellow
            'Console.WriteLine("Ripopolo le liste")
        End If
        'vado a rileggere Flink e linkedFlink
        If (indice_hyp <> 0) Then
            clearLista(listahyp, indice_hyp, 14) 'setto a zero l'indice dopo perchè l'indicazione serva a sapere cosa cancellare
            indice_hyp = 0
        End If
        If (indice_lhyp <> 0) Then
            clearLista(listaLinkedhyp, indice_lhyp, 14) 'setto a zero l'indice dopo perchè l'indicazione serva a sapere cosa cancellare
            indice_lhyp = 0
        End If
        If (indice <> 0) Then
            clearLista(listaFlink, indice, 14) 'setto a zero l'indice dopo perchè l'indicazione serva a sapere cosa cancellare
            indice = 0
        End If
        If (indice_Lflink <> 0) Then
            clearLista(listaLinkedFlink, indice_Lflink, 14) 'setto a zero l'indice dopo perchè l'indicazione serva a sapere cosa cancellare
            indice_Lflink = 0
        End If

        For sheet = 1 To sheet_n Step 1
            On Error Resume Next
            objWorksheet = objWorkbook.Worksheets(sheet)
            If (Err.Number <> 0) Then
                Dim mess = "Errore creazione oggetto sheet file Master Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore creazione oggetto sheet file Master Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            On Error Resume Next
            sheet_name = objWorksheet.Name
            If (Err.Number <> 0) Then
                Dim mess = "Errore estrae Nome del sheet file Master Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore estrae Nome del sheet file Master Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            'popola la lista degli HypLink
            FindHyper(objWorksheet, listahyp, indice_hyp, sheet_name, folderMaster & "\" & file_master, listaSegnalazioni, Att,
                      Uscita, fso, file_master, folderMaster, fileMaster, ofile, oOut, Fdebug)
            ' popola la lista dei link Flink
            FindToFrom(objWorksheet, listaFlink, indice, sheet_name, 1, folderMaster & "\" & file_master,
                       Att, folderMaster, file_master, fso, Uscita, fileMaster, ofile, oOut, Fdebug) ' 1 = TO 
            FindToFrom(objWorksheet, listaFlink, indice, sheet_name, 3, folderMaster & "\" & file_master,
                       Att, folderMaster, file_master, fso, Uscita, fileMaster, ofile, oOut, Fdebug) ' 3= TL e FL
        Next
        raccogliLinkedFlink(n_fileIn, listaFile, indice_Lflink, listaLinkedFlink, indice_lhyp, listaLinkedhyp, objWorkbook,
                            objWorkSecondbook, objWorkSecondsheet, objExcel, Att, myPass, listaSegnalazioni,
                            Uscita, fso, file_master, folderMaster, fileMaster, ofile, oOut, Fdebug)
        incrociaFlinkHyp(indice_hyp, listahyp, indice, listaFlink, indice_Lflink, listaLinkedFlink, indice_lhyp, listaLinkedhyp)
    End Function
    '--------------------------- fine PopolaListe

    Function settaPassword(ByRef WK, ByRef myPass, ByRef Att, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim wsp
        If (StrComp(myPass, "") <> 0) Then
            For Each wsp In WK.Worksheets
                On Error Resume Next
                If (wsp.ProtectContents) Then
                    wsp.Protect(myPass, "True", "True", "True", "True")
                    If (Err.Number = 1004) Then
                        If (Fdebug) Then
                            Dim mess = "Password differente su " & WK.Name & " Foglio:" & wsp.Name
                            WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
                        End If
                        Err.Clear()
                        'Att = True
                    Else
                        If (Fdebug) Then
                            Dim mess = WK.Name & " Foglio:" & wsp.Name & " Non Protetto"
                            WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
                        End If
                        Err.Clear()
                    End If
                    If (wsp.ProtectionMode) Then
                        If (Fdebug) Then
                            Dim mess = WK.Name & " Foglio:" & wsp.Name & " Accesso consentito allo script"
                            WriteMia(ConsoleColor.DarkYellow, mess, oOut, ofile)
                        End If
                        Err.Clear()
                        On Error GoTo 0
                    End If
                End If
            Next
        Else
            Dim mess = "Apro " & WK.Name & " senza password"
            WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Magenta
            'Console.WriteLine("Apro " & WK.Name & " senza password" & vbCrLf)
        End If
        'The three protection properties of a worksheet are the following:
        '  Sheets(1).ProtectContents
        '  Sheets(1).ProtectDrawingObjects
        '  Sheets(1).ProtectScenarios
        'You can check whether both 3 are False. If this is the case, it is not protected.
        '.ProtectionMode
    End Function

    Function settaFontPerCommenti(ByRef objWorkbook, dimFont, ByRef Att, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim ws, comm, testo, test_pos, indirizzo
        For Each ws In objWorkbook.Worksheets
            For Each comm In ws.Comments
                indirizzo = comm.Parent.Address
                testo = comm.text
                test_pos = InStr(testo, "HyElink")
                If (test_pos <> 0) Then
                    On Error Resume Next
                    'Console.WriteLine( "Setto Font su "&indirizzo&" "&comm.text&""&vbCrLf
                    With comm.Shape.TextFrame.Characters.Font
                        .Name = "Arial"
                        .Size = dimFont
                    End With
                    If (Err.Number <> 0) Then
                        Dim mess = "Errore " & Err.Description & " nel settaggio del font su " & comm.Parent.Address & " " & comm.text
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore " & Err.Description & " nel settaggio del font su " & comm.Parent.Address & " " & comm.text)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    On Error Resume Next
                    comm.Shape.TextFrame.AutoSize = True
                    If (Err.Number <> 0) Then
                        Dim mess = "Errore " & Err.Description & " nel settaggio Autosize su " & comm.Parent.Address & " " & comm.text
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore " & Err.Description & " nel settaggio Autosize su " & comm.Parent.Address & " " & comm.text)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                End If
            Next
        Next
    End Function

    Function cercaContropartePOS(punt, ByRef indice, ByRef listaFlink)
        'cerco la controparte che abbia come locazione POS il mio puntamento
        Dim z
        For z = 0 To indice - 1
            If ((listaFlink(14, z) = listaFlink(7, punt)) And (listaFlink(15, z) = listaFlink(8, punt)) _
            And (listaFlink(16, z) = listaFlink(10, punt))) Then
                cercaContropartePOS = z
                Exit Function
            End If
        Next
        cercaContropartePOS = -1
    End Function

    Function cercalinkedFlinkPOS(punt, ByRef indice, ByRef listaLinkedFlink, ByRef listaFlink)
        'cerco la controparte che abbia come locazione POS il mio puntamento
        Dim z
        For z = 0 To indice - 1
            If ((listaLinkedFlink(14, z) = listaFlink(7, punt)) And (listaLinkedFlink(15, z) = listaFlink(8, punt)) _
            And (listaLinkedFlink(16, z) = listaFlink(10, punt))) Then
                cercalinkedFlinkPOS = z
                Exit Function
            End If
        Next
        cercalinkedFlinkPOS = -1
    End Function

    Function controlloLink(ByRef indice, ByRef listaFlink, ByRef objWorksheet, ByRef listaLinkedFlink, ByRef indice_Lflink, ByRef Att,
                           ByRef folderMaster, ByRef file_master, ByRef objWorkSecondbook, ByRef objWorkSecondsheet,
                           ByRef objWorkbook, ByRef objExcel, ByRef myPass, ByRef Uscita, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim ris, hypCorr, hypControparte, Controparte, foglio_s, col_s, riga_s
        Console.WriteLine("Start controllo link")
        'Inizio cercando i link che si sono spostati (0,3,1 <> 14,15,16 su listaflink)
        For y = 0 To indice - 1   'spazzola tutta la listaFlink in cerca di link spostati
            If ((listaFlink(0, y) <> listaFlink(14, y)) Or (listaFlink(3, y) <> listaFlink(15, y)) Or (listaFlink(1, y) <> listaFlink(16, y))) Then
                'si è spostato in questo caso devo modificare:
                '                               se link Locale : l'Flink e l'Hyplink di chi mi puntava
                '                               se remoto      : llinkedFlink e l'Hyperlink sul file remoto 
                If (StrComp(listaFlink(5, y), "LinkSuSeStesso") = 0) Then
                    'link Locale modificare l'Flink e l'Hyplink di chi mi puntava
                    Controparte = cercaContropartePOS(y, indice, listaFlink) 'Devo cercare il Flink che ha come POS il mio puntamento
                    ' modifico l'Flink sul file
                    ris = modificaLinkToLocal(objWorksheet, y, listaFlink(0, Controparte), listaFlink(3, Controparte),
                                              listaFlink(1, Controparte), objWorkbook, listaFlink, Att, ofile, oOut, Fdebug) 'serve a scrivere la giusta POS sul link spostato (gli passo il PUNT per poter scrivere il link verso l'eventuale nuova cella della controparte)
                    ' nel caso anche la controparte non sia piu alla sua vecchia posizione.
                    ris = modificaLinkToLocal(objWorksheet, Controparte, listaFlink(0, y), listaFlink(3, y), listaFlink(1, y),
                                              objWorkbook, listaFlink, Att, ofile, oOut, Fdebug) 'serve a scrivere il giusto link sulla controparte (gli passo il nuovo PUNT)
                    ' modifico l'hyplink sul file
                    modificaHypTo(Controparte, listaFlink, listaFlink(0, y), listaFlink(2, y), listaFlink(3, y),
                                  listaFlink(1, y), objWorkbook, Uscita, Att, ofile, oOut, Fdebug)
                Else
                    'remoto      : llinkedFlink e l'Hyperlink sul file remoto
                    hypCorr = listaFlink(12, y)
                    Controparte = cercalinkedFlinkPOS(y, indice, listaLinkedFlink, listaFlink)
                    If (Controparte <> -1) Then
                        'adesso devo modificare sia il linkedFlink che l'hyperlink perchè puntino alla mia nuova posizione
                        foglio_s = listaLinkedFlink(7, Controparte)
                        col_s = listaLinkedFlink(8, Controparte)
                        riga_s = listaLinkedFlink(10, Controparte)
                        ' modifico l'Flink sul file
                        '
                        ris = modificaLinkTo(objWorksheet, y, "TO", Controparte, listaFlink, Att, objWorkbook, listaLinkedFlink,
                                             ofile, oOut, Fdebug) 'serve a scrivere la giusta POS sul link spostato con il link verso la reale posizione della cella puntata
                        If Not (ris) Then
                            Dim mess = "Errore in modificaLinkTo su:" & listaFlink(5, y)
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Errore in modificaLinkTo su:" & listaFlink(5, y))
                        End If
                        ris = scrivoLinkSuLinkedPOS(Controparte, y, objWorkSecondbook, objWorkSecondsheet, objWorkbook,
                                                    objExcel, listaFlink, Att, myPass, listaLinkedFlink,
                                                    fileMaster, ofile, oOut, Fdebug) 'scrivo il nuovo Flink
                        If Not (ris) Then
                            Dim mess = "Errore in scrivoLinkSuLinkedPOS su:" & listaFlink(5, y)
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Errore in scrivoLinkSuLinkedPOS su:" & listaFlink(5, y))
                        End If
                        'Corregge l'Hyperlink dal linkedFlink verso di me partendo dalla posizione RPos
                        modificaHypFromPOS(y, Controparte, objWorkbook, objWorkSecondbook, objExcel, listaFlink,
                                           listaLinkedFlink, Att, myPass, Uscita, ofile, oOut, Fdebug)
                        'modificaHypFrom y, listaFlink, foglio_s, col_s, riga_s, false
                    Else
                        'Console.ForegroundColor = ConsoleColor.Red
                        Dim mess = "Errore Non trovo la controparte su listaLinkedFlink a:" & listaFlink(5, y) & " s:" & listaFlink(0, y) & " c:" & listaFlink(3, y) & " r:" & listaFlink(1, y)
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    End If
                End If
            End If
        Next
        'Attenzione a non fare l'errore di ricostruire le liste in mezzo al controllo e modifica
        For y = 0 To indice_Lflink - 1 'spazzola tutta la listaLinkedFlink in cerca di link spostati
            If ((listaLinkedFlink(0, y) <> listaLinkedFlink(14, y)) Or (listaLinkedFlink(3, y) <> listaLinkedFlink(15, y)) Or (listaLinkedFlink(1, y) <> listaLinkedFlink(16, y))) Then
                'cerca in listaFlink un link che punta al mio POS, gli passo: parte iniziale file, sheetName,riga,colonna del mio POS
                'Console.ForegroundColor = ConsoleColor.Cyan
                Dim mess = "Debug Trovato LinkedFlink spostato indice " & y & " S:" & listaLinkedFlink(0, y) & " C:" & listaLinkedFlink(3, y) & " R:" & listaLinkedFlink(1, y)
                WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
                On Error Resume Next
                Controparte = cercaFlinkDaHyp(listaLinkedFlink(17, y), listaLinkedFlink(14, y), listaLinkedFlink(16, y), listaLinkedFlink(15, y), listaFlink, indice)
                '                             file parziale di questo linkedFlink,sheet di POS          , riga del POS         ,colonna del POS, lista dove cercare, indice della lista
                If (Err.Number <> 0) Then
                    mess = "Errore " & Err.Description & " nella chiamata a cercaFlinkDaHyp indice:" & y
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore " & Err.Description & " nella chiamata a cercaFlinkDaHyp indice:" & y)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
                If (Controparte <> -1) Then
                    'scrivo il nuovo pos sul Flink sul file linked, non modifico la lista quindi passo i valori della posizione con listaLinkedFlink
                    ris = newPOSLinked(Controparte, y, objWorkbook, objWorkSecondbook, objExcel, listaFlink, Att, myPass,
                                       objWorkSecondsheet, listaLinkedFlink, fileMaster, ofile, oOut, Fdebug)
                    If Not (ris) Then
                        mess = "Errore nella modifica del POS su LinkedFlink su:" & listaFlink(5, y) & " s:" & listaLinkedFlink(0, y) & " c:" & listaLinkedFlink(3, y) & " r:" & listaLinkedFlink(1, y)
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.WriteLine("Errore nella modifica del POS su LinkedFlink su:" & listaFlink(5, y) & " s:" & listaLinkedFlink(0, y) & " c:" & listaLinkedFlink(3, y) & " r:" & listaLinkedFlink(1, y))
                        Att = True
                    End If
                    'modifica del Flink della controparte che punti su di me Reale
                    ris = newFlinkHypLink(objWorksheet, listaFlink, Controparte, listaLinkedFlink(0, y), listaLinkedFlink(2, y),
                                          listaLinkedFlink(3, y), listaLinkedFlink(1, y), objWorkbook, Att, ofile, oOut, Fdebug)
                    If Not (ris) Then
                        mess = "Errore newFlinkHypLink modifica HyFlink#TO# su:" & folderMaster & "\" & fileMaster & " s:" & listaFlink(0, Controparte) & " c:" & listaFlink(3, Controparte) & " r:" & listaFlink(1, Controparte)
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore newFlinkHypLink modifica HyFlink#TO# su:" & folderMaster & "\" & fileMaster & " s:" & listaFlink(0, Controparte) & " c:" & listaFlink(3, Controparte) & " r:" & listaFlink(1, Controparte))
                        Att = True
                    End If
                    'Corregge l'Hyperlink della controparte perchè punti su di me reale
                    If (Fdebug) Then
                        mess = "Debug Vado a Correggere L'Hyperlink indice " & y & " S:" & listaLinkedFlink(0, y) & " C:" & listaLinkedFlink(3, y) & " R:" & listaLinkedFlink(1, y)
                        WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Cyan
                        'Console.WriteLine("Debug Vado a Correggere L'Hyperlink indice " & y & " S:" & listaLinkedFlink(0, y) & " C:" & listaLinkedFlink(3, y) & " R:" & listaLinkedFlink(1, y))
                    End If

                    modificaHypTo(Controparte, listaFlink, listaLinkedFlink(0, y), listaLinkedFlink(2, y), listaLinkedFlink(3, y),
                                  listaLinkedFlink(1, y), objWorkbook, Uscita, Att, ofile, oOut, Fdebug)
                Else
                    mess = "Errore In controlloLink Non trovo la controparte su listaFlink a:" & listaLinkedFlink(11, y) & " s:" & listaLinkedFlink(0, y) & " c:" & listaLinkedFlink(3, y) & " r:" & listaLinkedFlink(1, y)
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore In controlloLink Non trovo la controparte su listaFlink a:" & listaLinkedFlink(11, y) & " s:" & listaLinkedFlink(0, y) & " c:" & listaLinkedFlink(3, y) & " r:" & listaLinkedFlink(1, y))
                End If
            End If
        Next
    End Function
    'Questa funzione non viene usata ---------------Attenzione
    'Function creaHypLinkLocali(ByRef indice_hyp, ByRef listahyp, ByRef hyp_link_new, ByRef objWorkbook, ByRef objExcel, ByRef Att)
    'For y = 0 To indice_hyp - 1
    'If ((listahyp(14, y) = -1) And (StrComp(listahyp(5, y), "LinkSuSeStesso") = 0)) Then   'e un link locale e non si trova il corrispettivo
    'Console.WriteLine("<font color='red'>Andrei a creare l'hyp "&listahyp(5,y)&"->"&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font><br/>"&vbCrLf
    'on error goto 0
    'Wscript.Echo "<font color='red'>Andrei a creare l'hyp "&listahyp(5,y)&"->"&listahyp(7,y)&"!"&listahyp(8,y)&listahyp(10,y)&"</font><br/>"&vbCrLf
    '           creaHypXl(y, listahyp, objWorkbook, objExcel, Att)
    '           hyp_link_new = True
    'End If
    'Next
    'End Function


    Function creaFlinkLocali(ByRef indice_hyp, ByRef listahyp, ByRef scrittoLinkNuovi, ByRef Uscita, ByRef Att, ByRef objWorksheet,
                             ByRef objWorkbook, ByRef folderMaster, ByRef file_master, ByRef aggiuntaLink, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim ris As Integer
        Dim mess
        For y = 0 To indice_hyp - 1 Step 1
            If Not (listahyp(14, y) = -1) Then 'Significa che è linkato con un altro hyp
                If (listahyp(15, y) = -1) Then 'Signofica che non è stato ancora creato l'Flink
                    If Not (scrittoLinkNuovi) Then
                        Uscita = Uscita & "<tr><td align='center' colspan='10'><font color ='Navy'>Link nuovi</font></td></tr>"
                        scrittoLinkNuovi = True
                    End If
                    On Error Resume Next
                    objWorksheet = objWorkbook.Worksheets(listahyp(0, y))
                    If (Err.Number <> 0) Then
                        mess = "Errore creazione oggetto sheet Master</font><br/> Descrizione " & Err.Description
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore creazione oggetto sheet Master</font><br/> Descrizione " & Err.Description)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    mess = "Creazione Flink Local " & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y)
                    WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Yellow
                    'Console.WriteLine("Creazione Flink Local " & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y))
                    ris = scriviLinkTo(objWorksheet, listahyp, y, "XL", objWorkbook, Att, ofile, oOut, Fdebug)
                    If Not (ris) Then
                        mess = "Errore nella creazione di HyFlink su " & folderMaster & "\" & file_master
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore nella creazione di HyFlink su " & folderMaster & "\" & file_master)
                        Err.Clear()
                        Att = True
                    End If
                    Uscita = Uscita & "<tr><td><font color='maroon'>Aggiunto Link Local </font></td><td><font color='maroon'>" & folderMaster & "\" & file_master & "</font></td><td><font color ='maroon'> di:" & listahyp(0, y) & "!" & listahyp(3, y) & listahyp(1, y) & "</font></td><td><font color ='maroon'> a:" & listahyp(5, y) & " " & listahyp(7, y) & "!" & listahyp(8, y) & listahyp(10, y) & "</font></td></tr>"
                    aggiuntaLink = True
                    listahyp(15, y) = 8888
                End If
            End If
        Next

    End Function

    Function raccogliLinkedFlink(ByRef n_fileIn, ByRef listaFile, ByRef indice_Lflink, ByRef listaLinkedFlink, ByRef indice_lhyp, ByRef listaLinkedhyp, ByRef objWorkbook, ByRef objWorkSecondbook,
                                 ByRef objWorkSecondsheet, ByRef objExcel, ByRef Att, ByRef myPass, ByRef listaSegnalazioni, ByRef Uscita,
                                 ByRef fso, ByRef file_master, ByRef folderMaster, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim sheet_nf, bReadOnly, sheet_name As String
        Dim sheet As Integer
        Dim mess
        On Error GoTo 0
        For y = 0 To n_fileIn - 1
            If (InStr(1, listaFile(y), "LinkSuSeStesso") > 0) Then 'Il link è locale sullo stesso file
                'Non faccio niente i link Tl e FL sono già raccolti
            Else
                On Error Resume Next
                objWorkSecondbook = objExcel.Workbooks.Open(listaFile(y), False, False)
                If (Err.Number <> 0) Then
                    mess = "Errore nell'apertura file Linked Descrizione " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore nell'apertura file Linked Descrizione " & Err.Description)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
                bReadOnly = objWorkSecondbook.ReadOnly
                If bReadOnly = True Then
                    mess = "Errore apertura file " & listaFile(y) & ", File OCCUPATO</font><br/>"
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore apertura file " & listaFile(y) & ", File OCCUPATO</font><br/>" & vbCrLf)
                    Call objWorkSecondbook.Close
                    Call objWorkbook.Close
                    objExcel.Quit
                    mess = "Programma Terminato a causa di file Occupato"
                    WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Magenta
                    'Console.WriteLine("Programma Terminato a causa di file Occupato" & vbCrLf)
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine("Enter per Terminare")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.ReadLine()
                    ofile.Close()
                    End
                End If
                settaPassword(objWorkSecondbook, myPass, Att, ofile, oOut, Fdebug)
                sheet_nf = objWorkSecondbook.Sheets.Count
                For sheet = 1 To sheet_nf Step 1
                    On Error Resume Next
                    objWorkSecondsheet = objWorkSecondbook.Worksheets(sheet)
                    If (Err.Number <> 0) Then
                        mess = "Errore creazione oggetto sheet Linked Descrizione " & Err.Description
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.WriteLine("Errore creazione oggetto sheet Linked Descrizione " & Err.Description)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    On Error Resume Next
                    sheet_name = objWorkSecondsheet.Name
                    If (Err.Number <> 0) Then
                        mess = "Errore Name da oggetto sheet Linked Descrizione " & Err.Description
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore Name da oggetto sheet Linked Descrizione " & Err.Description)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    'raccoglie gli hyp remoti (LinkedHyp)
                    FindHyper(objWorkSecondsheet, listaLinkedhyp, indice_lhyp, sheet_name, listaFile(y), listaSegnalazioni,
                              Att, Uscita, fso, file_master, folderMaster, fileMaster, ofile, oOut, Fdebug)
                    'Set objRange = objWorkSecondsheet.UsedRange 'DAFARE verificare se serve
                    FindToFrom(objWorkSecondsheet, listaLinkedFlink, indice_Lflink, sheet_name, 2, listaFile(y), Att,
                               folderMaster, file_master, fso, Uscita, fileMaster, ofile, oOut, Fdebug) ' 2 = FROM
                Next 'Loop su tutti gli sheet di un file Input
                On Error Resume Next
                objWorkSecondbook.Close(False, listaFile(y))
                If (Err.Number <> 0) Then
                    mess = "Errore RaccogliLinkedFlink Close File Descrizione " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore RaccogliLinkedFlink Close File Descrizione " & Err.Description)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
            End If
        Next
    End Function

    Function CambiaFile(ByRef objWb, ByVal sheet_n, ByRef Att, ByRef Uscita, ByRef ListaFlink, ByRef indice, ByRef listaLinkedFlink,
                        ByRef indice_Lflink, ByRef Fdebug, ByRef file_master, ByRef fileMaster, ByRef ofile, ByRef oOut) 'Modifica il nome su tutti gli hyperlink per il objWorkbook passato
        Dim objWsh, sheet_name, hyp, lo_riga, lo_colonna, index_lf, posF, pos1, file
        Dim Flink = ""
        Dim file_new = ""
        Dim folder, info, sub_add, s_name, punt_flink, Part_ini

        For sheet = 1 To sheet_n Step 1
            On Error Resume Next
            objWsh = objWb.Worksheets(sheet)
            If (Err.Number <> 0) Then
                Dim mess = "Errore select sheet  Descrizione: " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore select sheet  Descrizione: " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            On Error Resume Next
            sheet_name = objWsh.Name
            If (Err.Number <> 0) Then
                Dim mess = "Errore Nome sheet  Descrizione: " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore Nome sheet  Descrizione: " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            For Each hyp In objWsh.Hyperlinks
                'bisogna estrarre il SubAddress e cercarlo nella listaFlink da questa estrarre il nuovo nome file e proseguire
                If Not (Len(hyp.Address) = 0) Then 'controlla che non sia un link su se stesso in questo caso non e necessario fare niente
                    lo_riga = 0
                    lo_colonna = 0
                    s_name = ""
                    SeparaRigheColonne(hyp.Parent.Address(0, 0), lo_riga, lo_colonna)
                    On Error Resume Next
                    file = hyp.Address
                    If (Err.Number <> 0) Then
                        Dim mess = "Errore estrazione Address di Hyperlink Descrizione: " & Err.Description
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore estrazione Address di Hyperlink Descrizione: " & Err.Description)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    On Error Resume Next
                    sub_add = hyp.SubAddress
                    If (Err.Number <> 0) Then
                        Dim mess = "Errore estrazione SubAddress di Hyperlink Descrizione: " & Err.Description
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore estrazione SubAddress di Hyperlink Descrizione: " & Err.Description)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    If (InStr(1, hyp.SubAddress, "!") = 0) Then
                        'Console.WriteLine("CambiaFile: hyplink estraneo al programma in " & sheet_name & "!"& hyp.Parent.Address(0, 0) &""&vbCrLf
                        Uscita = Uscita & "<tr><td><font color ='maroon'>-->hyplink estraneo al programma in " & sheet_name & "!" & hyp.Parent.Address(0, 0) & "</font></td></tr>"
                    Else
                        separSheetCR(Replace(sub_add, "'", ""), s_name, lo_riga, lo_colonna, ofile, oOut, Fdebug)
                        'INSERITA per diagnosi Console.WriteLine("Da 838 hyp.a:" & hyp.Address & " Sub: " & hyp.SubAddress & " Parent: " & hyp.Parent.Address(0, 0) & " r:"& lo_riga &" c:"&lo_colonna&"
                        Part_ini = estraiParteIniziale(hyp.Address, hyp, "HYPER", fileMaster, ofile, oOut, Fdebug)
                        'devo cercare anche per file parziale, altrimenti trovo un altro entry chee può semigliarci
                        On Error Resume Next
                        punt_flink = cercaFlinkDaHyp(Part_ini, s_name, lo_riga, lo_colonna, ListaFlink, indice) ' cerco l'entry di listaFlink corrispondente a questo hyplink
                        If (Err.Number <> 0) Then
                            Dim mess = "Errore punt_flink con ListaFlink " & punt_flink & " Descrizione: " & Err.Description
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Errore punt_flink con ListaFlink " & punt_flink & " Descrizione: " & Err.Description)
                            Err.Clear()
                            Att = True
                        End If
                        On Error GoTo 0
                        If (punt_flink = -1) Then
                            ' verifico se non si tratta di un file linkato
                            On Error Resume Next
                            punt_flink = cercaFlinkDaHyp(Part_ini, s_name, lo_riga, lo_colonna, listaLinkedFlink, indice_Lflink) ' cerco l'entry di linkedFlink corrispondente a questo hyplink
                            If (Err.Number <> 0) Then
                                Dim mess = "Errore punt_flink con listaLinkedFlink " & punt_flink & " Descrizione: " & Err.Description & "" & " Errore:" & Err.Description
                                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                                'Console.ForegroundColor = ConsoleColor.Red
                                'Console.WriteLine("Errore punt_flink con listaLinkedFlink " & punt_flink & " Descrizione: " & Err.Description & "" & " Errore:" & Err.Description & vbCrLf)
                                Err.Clear()
                                Att = True
                            End If
                            On Error GoTo 0
                            If (punt_flink <> -1) Then
                                file_new = listaLinkedFlink(5, punt_flink)
                            Else
                                Dim mess = "CambiaFile: Errore NON Grave non trovo la controparte Del collegamento: Parent:" & hyp.Parent.Address & " file:" & Part_ini & " foglio:" & s_name & " riga:" & lo_riga & " colonna:" & lo_colonna
                                WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
                                'Console.ForegroundColor = ConsoleColor.Magenta
                                'Console.WriteLine("CambiaFile: Errore NON Grave non trovo la controparte Del collegamento: Parent:" & hyp.Parent.Address & " file:" & Part_ini & " foglio:" & s_name & " riga:" & lo_riga & " colonna:" & lo_colonna)
                            End If
                        Else
                            file_new = ListaFlink(5, punt_flink)
                        End If
                        If Not (punt_flink = -1) Then
                            If Not (InStr(1, UCase(hyp.Address), UCase(file_new)) > 0) Then
                                info = "Address            :" & hyp.Address & vbCrLf &
                                   "SubAddress      :" & hyp.SubAddress & vbCrLf &
                                   "ScreenTip       :" & hyp.ScreenTip & vbCrLf &
                                   "TextToDisplay     :" & hyp.TextToDisplay & vbCrLf &
                                   "Flink             :" & Flink & vbCrLf &
                                   "Nuovo file         :" & file_new
                                hyp.Address = file_new
                            End If
                        End If
                    End If
                End If 'controlla che non sia un link su se stesso
            Next
        Next
    End Function

    Function scrivoLinkSuLinkedPOS(ByVal y, ByVal idy, ByRef objWorkSecondbook, ByRef objWorkSecondsheet, ByRef objWorkbook, ByRef objExcel,
                                   ByRef listaFlink, ByRef Att, ByRef myPass, ByRef listalinkedFlink, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug) 'scrive il Flink sul file remoto usando la giusta posizione e non quella precedente
        'scrivoLinkSuLinkedPOS(Controparte, y) y è il puntatore di listaFlink per ottenere il nuovo puntamento
        Dim commento_esiste, comm
        Dim commento = ""
        Dim bReadOnly = 0
        On Error Resume Next
        objWorkSecondbook = objExcel.Workbooks.Open(listaFlink(5, idy), False, False)
        If (Err.Number <> 0) Then
            Dim mess = "Errore scrivoLinkSuLinkedPOS: apertura Linked " & listaFlink(5, idy) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore scrivoLinkSuLinkedPOS: apertura Linked " & listaFlink(5, idy) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        bReadOnly = objWorkSecondbook.ReadOnly
        If bReadOnly = True Then
            Dim mess = "Errore apertura file " & listaFlink(5, idy) & ", File OCCUPATO"
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore apertura file " & listaFlink(5, idy) & ", File OCCUPATO" & vbCrLf)
            Call objWorkSecondbook.Close
            Call objWorkbook.Close
            objExcel.Quit
            mess = "Programma Terminato a causa di file Occupato"
            WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Magenta
            'Console.WriteLine("Programma Terminato a causa di file Occupato" & vbCrLf)
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.WriteLine("Enter per Terminare")
            Console.ForegroundColor = ConsoleColor.White
            Console.ReadLine()
            ofile.Close()
            End
        End If
        settaPassword(objWorkSecondbook, myPass, Att, ofile, oOut, Fdebug)
        On Error Resume Next
        objWorkSecondsheet = objWorkSecondbook.Worksheets(listalinkedFlink(0, y)) 'mi setto sul giusto sheet
        If (Err.Number <> 0) Then
            Dim mess = "Errore scrivoLinkSuLinkedPOS: set sheet " & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore scrivoLinkSuLinkedPOS: set sheet " & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error Resume Next
        commento_esiste = True
        comm = objWorkSecondsheet.Cells(listalinkedFlink(1, y), listalinkedFlink(2, y)).Comment.Text
        If (Err.Number <> 0) Then
            'Il commento non esiste ancora
            comm = ""
            commento_esiste = False
            Err.Clear()
        End If
        On Error GoTo 0
        If (commento_esiste) Then
            If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
                'rimuovo il link
                commento = rimuoviLink(comm)
            Else
                commento = comm
            End If
        End If
        commento = commento & vbCrLf & "HyFlink#FR#Pos=PS=" & listalinkedFlink(0, y) & "#PC=" & listalinkedFlink(3, y) & "#PR=" & listalinkedFlink(1, y) & "#cartella=.#file=" & fileMaster & "#S=" & listaFlink(0, idy) & "#C=" & listaFlink(3, idy) & "#R=" & listaFlink(1, idy) & "#HyElink"
        commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
        If (commento_esiste) Then
            On Error Resume Next
            objWorkSecondsheet.Cells(listalinkedFlink(1, y), listalinkedFlink(2, y)).ClearComments
            If (Err.Number <> 0) Then
                Console.WriteLine("Errore scrivoLinkSuLinkedPOS: add comment " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
        End If
        On Error Resume Next
        objWorkSecondsheet.Cells(listalinkedFlink(1, y), listalinkedFlink(2, y)).AddComment(commento)
        If (Err.Number <> 0) Then
            Dim mess = "Errore scrivoLinkSuLinkedPOS: add comment " & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore scrivoLinkSuLinkedPOS: add comment " & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objWorkSecondbook.Save
        If (Err.Number <> 0) Then
            Dim mess = "Errore scrivoLinkSuLinkedPOS: Save linked Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore scrivoLinkSuLinkedPOS: Save linked Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objWorkSecondbook.Close
        If (Err.Number <> 0) Then
            Dim mess = "Errore scrivoLinkSuLinkedPOS: Close Linked Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore scrivoLinkSuLinkedPOS: Close Linked Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        scrivoLinkSuLinkedPOS = True
    End Function

    Function newPOSLinked(ByVal idlf, ByVal idll, ByRef objWorkbook, ByRef objWorkSecondbook, ByRef objExcel, ByRef listaFlink,
                          ByRef Att, ByRef myPass, ByRef objWorkSecondsheet, ByRef listaLinkedFlink, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug)
        'newPOSLinked(Controparte,y)
        'Vado a scrivere la nuova POS sul link remoto
        Dim commento_esiste, comm
        Dim commento = ""
        Dim bReadOnly = 0
        On Error Resume Next
        objWorkSecondbook = objExcel.Workbooks.Open(listaFlink(5, idlf), False, False) 'apro il file lincato
        If (Err.Number <> 0) Then
            Dim mess = "Errore newPOSLinked: apertura Linked " & listaFlink(5, idlf) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore newPOSLinked: apertura Linked " & listaFlink(5, idlf) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        bReadOnly = objWorkSecondbook.ReadOnly
        If bReadOnly = True Then
            Dim mess = "Errore apertura file " & listaFlink(5, idlf) & ", File OCCUPATO</font><br/>" & vbCrLf
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore apertura file " & listaFlink(5, idlf) & ", File OCCUPATO</font><br/>" & vbCrLf)
            Call objWorkSecondbook.Close
            Call objWorkbook.Close
            objExcel.Quit
            mess = "Programma Terminato a causa di file Occupato"
            WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Magenta
            'Console.WriteLine("Programma Terminato a causa di file Occupato" & vbCrLf)
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.WriteLine("Enter per Terminare")
            Console.ForegroundColor = ConsoleColor.White
            Console.ReadLine()
            ofile.Close()
            End
        End If
        settaPassword(objWorkSecondbook, myPass, Att, ofile, oOut, Fdebug)
        On Error Resume Next
        objWorkSecondsheet = objWorkSecondbook.Worksheets(listaLinkedFlink(0, idll)) 'mi setto sul giusto sheet
        If (Err.Number <> 0) Then
            Dim mess = "Errore newPOSLinked: set sheet " & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore newPOSLinked: set sheet " & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        commento_esiste = True
        comm = objWorkSecondsheet.Cells(listaLinkedFlink(1, idll), listaLinkedFlink(2, idll)).Comment.Text
        If (Err.Number <> 0) Then
            'Il commento non esiste ancora
            comm = ""
            commento_esiste = False
            Err.Clear()
        End If
        On Error GoTo 0
        If (commento_esiste) Then
            If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
                'rimuovo il link
                commento = rimuoviLink(comm)
            Else
                commento = comm
            End If
        End If
        commento = commento & vbCrLf & "HyFlink#FR#Pos=PS=" & listaLinkedFlink(0, idll) & "#PC=" & listaLinkedFlink(3, idll) & "#PR=" & listaLinkedFlink(1, idll) & "#cartella=.#file=" & fileMaster & "#S=" & listaFlink(0, idlf) & "#C=" & listaFlink(3, idlf) & "#R=" & listaFlink(1, idlf) & "#HyElink"
        commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
        If (commento_esiste) Then
            On Error Resume Next
            objWorkSecondsheet.Cells(listaLinkedFlink(1, idll), listaLinkedFlink(2, idll)).ClearComments
            If (Err.Number <> 0) Then
                Dim mess = "Errore newPOSLinked: add comment " & Err.Number & " Description " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore newPOSLinked: add comment " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
        End If
        On Error Resume Next
        objWorkSecondsheet.Cells(listaLinkedFlink(1, idll), listaLinkedFlink(2, idll)).AddComment(commento)
        If (Err.Number <> 0) Then
            Dim mess = "Errore newPOSLinked: add comment " & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore newPOSLinked: add comment " & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objWorkSecondbook.Save
        If (Err.Number <> 0) Then
            Dim mess = "Errore newPOSLinked: Save linked Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore newPOSLinked: Save linked Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objWorkSecondbook.Close
        If (Err.Number <> 0) Then
            Dim mess = "Errore newPOSLinked: Close Linked Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore newPOSLinked: Close Linked Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        newPOSLinked = True
    End Function


    Function scrivoLinkSuLinked(ByVal y, ByRef lista_ref, ByVal hyp_yn, ByVal chiave, ByRef objWorksheet, ByRef objWorkbook,
                                ByRef Att, ByRef objWorkSecondbook, ByRef objWorkSecondsheet, ByRef objExcel, ByRef myPass,
                                ByRef folderMaster, ByRef file_master, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug) 'hyp se deve fare anche l'hyperlink
        Dim idx = 0
        Dim objSheet
        Dim commento = ""
        Dim objLink, comm, commento_esiste
        Dim cartella = ""
        Dim bReadOnly = 0
        'Inizio la scrittura del HyFlink#FR# sul file linkato
        If (InStr(1, lista_ref(5, y), "LinkSuSeStesso") > 0) Then 'Il link è locale sullo stesso file
            'devo lavorare sul file master aperto
            chiave = "XL"
            On Error Resume Next
            objWorksheet = objWorkbook.Worksheets(lista_ref(7, y)) 'mi setto sul giusto sheet
            If (Err.Number <> 0) Then
                Dim mess = "Errore set sheet scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore set sheet scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            'objWorksheet.Cells(lista_ref(10,y), lista_ref(9,y)).ClearComments
            On Error Resume Next
            comm = objWorksheet.Cells(lista_ref(10, y), lista_ref(9, y)).Comment.Text
            If (Err.Number <> 0) Then
                Dim mess = "Errore clear comment scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore clear comment scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
                'rimuovo il link
                commento = rimuoviLink(comm)
            Else
                commento = comm
            End If

            'HyFlink#XL#Pos=S=xx#C=jj#R=rr#Punt=S=xx#C=jj#R=rr#HyElink
            commento = commento & vbCrLf & "HyFlink#" & chiave & "#Pos=PS=" & lista_ref(7, y) & "#PC=" & lista_ref(8, idx) & "#PR=" & lista_ref(10, y) & "#Punt=#S=" & lista_ref(0, y) & "#C=" & lista_ref(3, y) & "#R=" & lista_ref(1, y) & "#HyElink"
            commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
            On Error Resume Next
            objWorksheet.Cells(lista_ref(10, y), lista_ref(9, y)).ClearComments
            If (Err.Number <> 0) Then
                Dim mess = "Errore clear commento 3 scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.WriteLine("Errore clear commento 3 scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            On Error Resume Next
            objWorksheet.Cells(lista_ref(10, y), lista_ref(9, y)).AddComment(commento)
            If (Err.Number <> 0) Then
                Dim mess = "Errore add comment scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore add comment scrivoLinkSuLinked: " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            If (hyp_yn) Then
                On Error Resume Next
                'Ho modificato da (objWorksheet.Name) a (lista_ref(7,y))
                objLink = objWorksheet.Hyperlinks.Add(objWorkbook.Worksheets(lista_ref(7, y)).Range("'" & lista_ref(7, y) & "'!" & lista_ref(8, y) & lista_ref(10, y)),
                    "",
                    "'" & lista_ref(0, y) & "'!" & lista_ref(3, y) & lista_ref(1, y),
                    "hypCreato")
                If (Err.Number <> 0) Then
                    Dim mess = "Errore scrivoLinkSuLinked: di aggiunta" & lista_ref(7, y) & "!" & lista_ref(8, y) & lista_ref(10, y) & "Descrizione: " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.WriteLine("Errore scrivoLinkSuLinked: di aggiunta" & lista_ref(7, y) & "!" & lista_ref(8, y) & lista_ref(10, y) & "Descrizione: " & Err.Description)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
            End If
            'Fine scrittura HyFlink#FR#
        Else
            On Error Resume Next
            objWorkSecondbook = objExcel.Workbooks.Open(lista_ref(5, y), False, False)
            If (Err.Number <> 0) Then
                Dim mess = "Errore scrivoLinkSuLinked: apertura Linked " & lista_ref(5, y) & " Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore scrivoLinkSuLinked: apertura Linked " & lista_ref(5, y) & " Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            bReadOnly = objWorkSecondbook.ReadOnly
            If bReadOnly = True Then
                Dim mess = "Errore apertura file " & lista_ref(5, y) & ", File Occupato"
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore apertura file " & lista_ref(5, y) & ", File OCCUPATO" & vbCrLf)
                Call objWorkSecondbook.Close
                Call objWorkbook.Close
                objExcel.Quit
                mess = "Programma Terminato a causa di file Occupato"
                WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Magenta
                'Console.WriteLine("Programma Terminato a causa di file Occupato" & vbCrLf)
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.WriteLine("Enter per Terminare")
                Console.ForegroundColor = ConsoleColor.White
                Console.ReadLine()
                ofile.Close()
                End
            End If
            settaPassword(objWorkSecondbook, myPass, Att, ofile, oOut, Fdebug)
            On Error Resume Next
            objWorkSecondsheet = objWorkSecondbook.Worksheets(lista_ref(7, y)) 'mi setto sul giusto sheet
            If (Err.Number <> 0) Then
                Dim mess = "Errore scrivoLinkSuLinked: set sheet " & Err.Number & " Description " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore scrivoLinkSuLinked: set sheet " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            On Error Resume Next
            commento_esiste = True
            comm = objWorkSecondsheet.Cells(lista_ref(10, y), lista_ref(9, y)).Comment.Text
            If (Err.Number <> 0) Then
                'Il commento non esiste ancora
                comm = ""
                commento_esiste = False
                Err.Clear()
            End If
            On Error GoTo 0
            If (commento_esiste) Then
                If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
                    'rimuovo il link
                    commento = rimuoviLink(comm)
                Else
                    commento = comm
                End If
            End If
            'Qui devo usare il relativo nella costruzione del commento 
            commento = commento & vbCrLf & "HyFlink#FR#Pos=PS=" & lista_ref(7, y) & "#PC=" & lista_ref(8, y) & "#PR=" & lista_ref(10, y) & "#cartella=.#file=" & fileMaster & "#S=" & lista_ref(0, y) & "#C=" & lista_ref(3, y) & "#R=" & lista_ref(1, y) & "#HyElink"
            commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
            If (commento_esiste) Then
                On Error Resume Next
                objWorkSecondsheet.Cells(lista_ref(10, y), lista_ref(9, y)).ClearComments
                If (Err.Number <> 0) Then
                    Dim mess = "Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
            End If
            On Error Resume Next
            objWorkSecondsheet.Cells(lista_ref(10, y), lista_ref(9, y)).AddComment(commento)
            If (Err.Number <> 0) Then
                Dim mess = "Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore scrivoLinkSuLinked: add comment " & Err.Number & " Description " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            If (hyp_yn) Then
                'Crea L'Hyperlink
                On Error Resume Next
                objLink = objWorkSecondsheet.Hyperlinks.Add(objWorkSecondbook.Worksheets(objWorkSecondsheet.name).Range("'" & lista_ref(7, y) & "'!" & lista_ref(8, y) & lista_ref(10, y)),
                    folderMaster & "\" & file_master,
                    "'" & lista_ref(0, y) & "'!" & lista_ref(3, y) & lista_ref(1, y),
                    "hyplink=" & folderMaster & "\" & file_master & "-" & lista_ref(0, y) & "'!" & lista_ref(3, y) & lista_ref(1, y))
                If (Err.Number <> 0) Then
                    Dim mess = "Errore scrivoLinkSuLinked: di aggiunta" & lista_ref(7, y) & "!" & lista_ref(8, y) & lista_ref(10, y) & "Descrizione: " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.WriteLine("Errore scrivoLinkSuLinked: di aggiunta" & lista_ref(7, y) & "!" & lista_ref(8, y) & lista_ref(10, y) & "Descrizione: " & Err.Description)
                    Err.Clear()
                    Att = True
                End If
                On Error GoTo 0
                'Fine della creazione dell'Hyperlink
            End If
            On Error Resume Next
            objWorkSecondbook.Save
            If (Err.Number <> 0) Then
                Dim mess = "Errore scrivoLinkSuLinked: Save linked Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore scrivoLinkSuLinked: Save linked Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            On Error Resume Next
            objWorkSecondbook.Close
            If (Err.Number <> 0) Then
                Dim mess = "Errore scrivoLinkSuLinked: Close Linked Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore scrivoLinkSuLinked: Close Linked Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
        End If
        'Fine scrittura HyFlink#FR#
        scrivoLinkSuLinked = True ' in realtà se non è andata bene sono già uscito dal programma
    End Function

    Function inilistaSegnalazioni(ByRef listaSegnalazioni)
        Dim h
        For h = 0 To 99
            listaSegnalazioni(h) = "#"
        Next
        inilistaSegnalazioni = 0
    End Function

    Function nomeFileSegnalato(ByVal file, ByRef listaSegnalazioni)
        Dim h
        nomeFileSegnalato = False
        For h = 0 To 99
            If (InStr(1, listaSegnalazioni(h), file) > 0) Then
                nomeFileSegnalato = True
                Exit Function
            Else
                If (InStr(1, listaSegnalazioni(h), "#") > 0) Then
                    listaSegnalazioni(h) = file
                    Exit Function
                End If
            End If
        Next
    End Function

    Function FindHyper(sheet, lista, ByRef idx, s_loc_name, fileInUso, ByRef listaSegnalazioni, ByRef Att, ByRef Uscita, ByRef fso,
                       ByRef file_master, ByRef folderMaster, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim lo_riga = ""
        Dim lo_colonna = ""
        Dim li_riga = ""
        Dim li_colonna = ""
        Dim li_col_num = 0
        Dim s_name = ""
        Dim file_p_linked, folder, file_name
        For Each hyp In sheet.Hyperlinks
            SeparaRigheColonne(hyp.Parent.Address(0, 0), lo_riga, lo_colonna)
            If (InStr(1, hyp.SubAddress, "!") = 0) Then
                'Console.WriteLine("FindHyper: hyplink estraneo al programma in " & hyp.Parent.Address(0, 0) &""&vbCrLf
            Else
                separSheetCR(Replace(hyp.SubAddress, "'", ""), s_name, li_riga, li_colonna, ofile, oOut, Fdebug)
                s_name = Replace(s_name, "'", "")
                li_col_num = calcolaColonna(li_colonna, Att, ofile, oOut, Fdebug)
                file_p_linked = estraiParteIniziale(hyp.Address, hyp, "HYPER", fileMaster, ofile, oOut, Fdebug) 'Se vuoto = LinkSuSeStesso
                folder = estraiFolderDaAddress(hyp.Address, file_p_linked, folderMaster) 'Se vuoto = LinkSuSeStesso
                folder = Replace(folder, "/", "\")
                If Not (InStr(1, folder, "LinkSuSeStesso") > 0) Then
                    file_name = cercaFile(file_p_linked, folder, fso)
                    If (StrComp(file_name, "NULLA") = 0) Then
                        On Error Resume Next
                        Dim mess = "Errore cercaFile folder da hyp" & hyp.address & " NULLO"
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Errore cercaFile folder da hyp" & hyp.address & " NULLO" & vbCrLf)
                        Err.Clear()
                        Att = True
                    End If
                    On Error GoTo 0
                    If Not (InStr(1, hyp.Address, file_name) > 0) Then
                        If Not (nomeFileSegnalato(file_name, listaSegnalazioni)) Then
                            Uscita = Uscita & "<tr><td><font color ='orange'>Nome File modificato</font></td><td><font color ='orange'>da:" & hyp.Address & "</font></td><td></td><td><font color ='orange'> a:" & file_name & "</td></tr>"
                        End If
                    End If
                Else
                    file_name = "LinkSuSeStesso"
                End If
                lista(0, idx) = s_loc_name                              ' 0	sheet (dove è registrato il link)
                lista(1, idx) = lo_riga                                 ' 1	riga (dove è registrato il link)
                lista(2, idx) = calcolaColonna(lo_colonna, Att, ofile, oOut, Fdebug)              ' 2	colonna numero (dove è registrato il link)
                lista(3, idx) = lo_colonna                              ' 3	colonna lettere (dove è registrato il link)
                lista(4, idx) = hyp.Address                             ' 4 Address = path relativo e nome del file linkato (relativo alla directory del file excel)
                If (InStr(1, file_name, "LinkSuSeStesso") > 0) Then
                    lista(5, idx) = "LinkSuSeStesso"
                Else
                    lista(5, idx) = folder & "\" & file_name
                End If                                                      ' 5 	Path completo del file linkato
                lista(6, idx) = hyp.SubAddress                          ' 6 	SubAddress nome_dello_sheet!colonna_lettereRiga
                lista(7, idx) = s_name                                  ' 7 	Nome sheet linked
                lista(8, idx) = li_colonna                              ' 8		Colonna in lettere linked
                lista(9, idx) = li_col_num                              ' 9 	Colonna numero linked
                lista(10, idx) = li_riga                                    ' 10	Riga linked
                lista(11, idx) = file_p_linked                            ' 11    file linked parte iniziale
                lista(12, idx) = -1                                       ' 12    link con listaFlink
                lista(13, idx) = -1                                       ' 13    link con listaLinkedFlink
                lista(14, idx) = -1                                       ' 14    link locale su lista
                lista(15, idx) = -1                                       ' 15    indica la situazione del Flink
                lista(16, idx) = -1                                       ' 16    indica se anche il puntamento di ritorno è giusto (spostamento)
                idx = idx + 1
            End If
        Next
    End Function

    Function creaHypXl(ByVal punt, ByRef lista, ByRef objWorkbook, ByRef Att, ByRef objExcel, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim ws, objSheet, objLink, objRange, mess
        On Error Resume Next
        objSheet = objWorkbook.Worksheets(lista(7, punt)) 'lo sheet dove risiede la controparte
        If (Err.Number <> 0) Then
            mess = "Errore creaHypXl: Sheet " & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore creaHypXl: Sheet " & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objRange = objExcel.Range("'" & lista(7, punt) & "'!" & lista(8, punt) & lista(10, punt))
        If (Err.Number <> 0) Then
            mess = "Errore creaHypXl: Range " & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore creaHypXl: Range " & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objLink = objSheet.Hyperlinks.Add(objRange, "", "'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt), "hypCreato")
        If (Err.Number <> 0) Then
            mess = "Errore creaHypXl: di aggiunta" & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore creaHypXl: di aggiunta" & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description)
            Err.Clear()
            Att = True
        End If
        mess = "creaHypXl: CREATO Link locale " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt)
        WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Yellow
        'Console.WriteLine("creaHypXl: CREATO Link locale " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt))
        On Error GoTo 0
    End Function

    Function modificaHypTo(ByVal punt, ByRef lista, p_sh, p_cn, p_c, p_r, ByRef objWorkbook, ByRef Uscita, ByRef Att, ByRef ofile, ByRef oOut, ByRef Fdebug)
        'modificaHypTo Controparte,listaFlink, listaFlink(0,y), listaFlink(2,y), listaFlink(3,y), listaFlink(1,y)
        Dim ws, objSheet, subb, subbTest, mess
        objSheet = objWorkbook.Worksheets(lista(0, punt))
        subb = objSheet.range("'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt)).Hyperlinks(1).SubAddress
        If (InStr(lista(6, punt), "'") <> 0) Then
            subbTest = subb
        Else
            subbTest = Replace(subb, "'", "")
        End If
        ' TOLTO in quanto se vado a modificarlo e non lo modifico nella lista ovviamente è diverso
        'if ((subbTest = lista(6,punt)) OR (spostamentoDalPos))  then
        On Error Resume Next
        'setto il link della controparte (reale) su di me (reale)
        objSheet.range("'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt)).Hyperlinks(1).SubAddress = "'" & p_sh & "'!" & p_c & p_r
        If (Err.Number <> 0) Then
            mess = "Errore modificaHypTo: modifica SubAddress " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore modificaHypTo: modifica SubAddress " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        mess = "modificaHypTo: Ipertesto modificato da:" & subb & " a:" & objSheet.range("'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt)).Hyperlinks(1).SubAddress
        WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Yellow
        'Console.WriteLine("modificaHypTo: Ipertesto modificato da:" & subb & " a:" & objSheet.range("'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt)).Hyperlinks(1).SubAddress)
        Uscita = Uscita & "<tr><td colspan=10><font color ='orange'>Ipertesto modificato da:" & subb & " a:" & objSheet.range("'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt)).Hyperlinks(1).SubAddress & "</font></td></tr>"
        'else
        '	On Error Resume Next
        '	Console.WriteLine("modificaHypTo: L'Hyperlink non corrisponde :"&objSheet.range("'"&lista(0,punt)&"'!"&lista(3,punt)&lista(1,punt)).Hyperlinks(1).SubAddress&" diverso da:"&lista(6,punt)&""&vbCrLf
        '	on error goto 0
        '	Err.Clear
        'end if
    End Function

    Function modificaHypFromPOS(ByVal punt, ByVal idllf, ByRef objWorkbook, ByRef objWorkSecondbook, ByRef objExcel, ByRef listaFlink,
                                ByRef listaLinkedFlink, ByRef Att, ByRef myPass, ByRef Uscita, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim bReadOnly
        Dim mess
        'modificaHypFromPOS y, Controparte
        'funzione che va a modificare l'hyperlink remoto puntando su listaFlink selezionata ma scrivendo sulla posizione reale della linkedFlink
        'quindi gli passo la listaFlink selezionata ma gli passo anche la RPos della linkedFlink e il foglio,colonna,righa di dove ero prima
        Dim objSheet, subb, objRange, objLink, testSubb
        On Error Resume Next
        objWorkSecondbook = objExcel.Workbooks.Open(listaFlink(5, punt), False, False)
        If (Err.Number <> 0) Then
            mess = "Errore modificaHypFrom: apertura Linked " & listaFlink(5, punt) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.WriteLine("Errore modificaHypFrom: apertura Linked " & listaFlink(5, punt) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        bReadOnly = objWorkSecondbook.ReadOnly
        If bReadOnly = True Then
            mess = "Errore apertura file " & listaFlink(5, punt) & ", File OCCUPATO"
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore apertura file " & listaFlink(5, punt) & ", File OCCUPATO")
            Call objWorkSecondbook.Close
            Call objWorkbook.Close
            objExcel.Quit
            mess = "Programma Terminato a causa di file Occupato"
            WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Magenta
            'Console.WriteLine("Programma Terminato a causa di file Occupato" & vbCrLf)
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.WriteLine("Enter per Terminare")
            Console.ForegroundColor = ConsoleColor.White
            Console.ReadLine()
            ofile.Close()
            End
        End If
        settaPassword(objWorkSecondbook, myPass, Att, ofile, oOut, Fdebug)
        On Error Resume Next
        objSheet = objWorkSecondbook.Worksheets(listaLinkedFlink(0, idllf))
        If (Err.Number <> 0) Then
            mess = "Errore modificaHypFrom: setta sheet Linked " & listaLinkedFlink(0, idllf) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.WriteLine("Errore modificaHypFrom: setta sheet Linked " & listaLinkedFlink(0, idllf) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        subb = objSheet.range("'" & listaLinkedFlink(0, idllf) & "'!" & listaLinkedFlink(3, idllf) & listaLinkedFlink(1, idllf)).Hyperlinks(1).SubAddress
        If (Err.Number <> 0) Then
            mess = "Errore modificaHypFrom: Copia SubAddress Linked " & listaLinkedFlink(0, idllf) & "'!" & listaLinkedFlink(3, idllf) & listaLinkedFlink(1, idllf) & "su:" & listaFlink(5, punt) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore modificaHypFrom: Copia SubAddress Linked " & listaLinkedFlink(0, idllf) & "'!" & listaLinkedFlink(3, idllf) & listaLinkedFlink(1, idllf) & "su:" & listaFlink(5, punt) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        If (InStr(1, subb, "'") > 0) Then
            testSubb = Replace(subb, "'", "")
        End If
        'if (testSubb = listaLinkedFlink(6,idllf)) then
        'Se sto modificando il link sul reale in modo diverso da quello sulle liste non posso controllarlo sulle liste
        On Error Resume Next
        objSheet.range("'" & listaLinkedFlink(0, idllf) & "'!" & listaLinkedFlink(3, idllf) & listaLinkedFlink(1, idllf)).Hyperlinks(1).SubAddress = "'" & listaFlink(0, punt) & "'!" & listaFlink(3, punt) & listaFlink(1, punt)
        If (Err.Number <> 0) Then
            mess = "Errore modificaHypFrom: Modifica SubAddress Linked " & listaLinkedFlink(0, idllf) & "'!" & listaLinkedFlink(3, idllf) & listaLinkedFlink(1, idllf) & "su:" & listaFlink(5, punt) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore modificaHypFrom: Modifica SubAddress Linked " & listaLinkedFlink(0, idllf) & "'!" & listaLinkedFlink(3, idllf) & listaLinkedFlink(1, idllf) & "su:" & listaFlink(5, punt) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        mess = "modifica Ipertesto da: " & subb & " a: " & "'" & listaFlink(0, punt) & "'!" & listaFlink(3, punt) & listaFlink(1, punt)
        WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
        'Console.ForegroundColor = ConsoleColor.Yellow
        'Console.WriteLine("modifica Ipertesto da: " & subb & " a: " & "'" & listaFlink(0, punt) & "'!" & listaFlink(3, punt) & listaFlink(1, punt))
        Uscita = Uscita & "<tr><td colspan=10><font color ='Purple'>modifica Ipertesto da: " & subb & " a: " & "'" & listaFlink(0, punt) & "'!" & listaFlink(3, punt) & listaFlink(1, punt) & "</font></td></tr>"
        On Error Resume Next
        objWorkSecondbook.Save
        If (Err.Number <> 0) Then
            mess = "Errore modificaHypFromPOS: Salva Linked " & listaFlink(5, punt) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore modificaHypFromPOS: Salva Linked " & listaFlink(5, punt) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objWorkSecondbook.Close
        If (Err.Number <> 0) Then
            mess = "Errore modificaHypFromPOS: Chiudi Linked " & listaFlink(5, punt) & " Descrizione " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore modificaHypFromPOS: Chiudi Linked " & listaFlink(5, punt) & " Descrizione " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
    End Function

    Function modificaHypFrom(ByVal punt, ByRef lista, ByVal foglio_s, ByVal col_s, ByVal riga_s, aggiungi, ByRef objWorkbook,
                             ByRef objWorkSecondbook, ByRef objExcel, ByRef Att, ByRef myPass, ByRef folderMaster,
                             ByRef file_master, ByRef Uscita, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim objSheet As Object, subb, objRange, objLink, bReadOnly As Integer
        If (InStr(1, lista(5, punt), "LinkSuSeStesso") > 0) Then 'e un link su se stesso
            On Error Resume Next
            objSheet = objWorkbook.Worksheets(lista(7, punt))
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: setta sheet Linked Su " & lista(7, punt) & " Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("modificaHypFrom: Errore setta sheet Linked SuSe " & lista(7, punt) & " Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
        Else  'e un link normale
            On Error Resume Next
            objWorkSecondbook = objExcel.Workbooks.Open(lista(5, punt), False, False)
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: apertura Linked " & lista(5, punt) & " Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: apertura Linked " & lista(5, punt) & " Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            bReadOnly = objWorkSecondbook.ReadOnly
            If bReadOnly = True Then
                Dim mess = "Errore apertura file " & lista(5, punt) & ", File OCCUPATO"
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore apertura file " & lista(5, punt) & ", File OCCUPATO")
                Call objWorkSecondbook.Close
                Call objWorkbook.Close
                objExcel.Quit
                mess = "Programma Terminato a causa di file Occupato"
                WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Magenta
                'Console.WriteLine("Programma Terminato a causa di file Occupato" & vbCrLf & vbCrLf)
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.WriteLine("Enter per Terminare")
                Console.ForegroundColor = ConsoleColor.White
                Console.ReadLine()
                ofile.Close()
                End
            End If
            settaPassword(objWorkSecondbook, myPass, Att, ofile, oOut, Fdebug) 'objWorkSecondbook
            On Error Resume Next
            objSheet = objWorkSecondbook.Worksheets(lista(7, punt))
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: setta sheet Linked " & lista(7, punt) & " Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: setta sheet Linked " & lista(7, punt) & " Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
        End If  'e un link su se stesso o normale
        If (aggiungi) Then 'Aggiunta
            'WScript.Echo "Vado in aggiunta" & lista(7,punt) & "!" & lista(8,punt) & lista(10,punt)
            'n error resume next
            On Error Resume Next
            objRange = objExcel.Range("'" & lista(7, punt) & "'!" & lista(8, punt) & lista(10, punt))
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: Range " & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: Range " & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            On Error Resume Next
            objLink = objSheet.Hyperlinks.Add(objRange, folderMaster & "\" & file_master, "'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt), "hypCreato")
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: di aggiunta" & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: di aggiunta" & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "Descrizione: " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            'WScript.Echo "Finita aggiunta" & lista(7,punt) & "!" & lista(8,punt) & lista(10,punt)
        Else 'Modifica
            On Error Resume Next
            subb = objSheet.range("'" & lista(7, punt) & "'!" & lista(8, punt) & lista(10, punt)).Hyperlinks(1).SubAddress
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: Copia SubAddress Linked " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "su:" & lista(5, punt) & " Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: Copia SubAddress Linked " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "su:" & lista(5, punt) & " Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            If ((subb = "'" & foglio_s & "'!" & col_s & riga_s) Or (subb = foglio_s & "!" & col_s & riga_s)) Then
                On Error Resume Next
                objSheet.range("'" & lista(7, punt) & "'!" & lista(8, punt) & lista(10, punt)).Hyperlinks(1).SubAddress = "'" & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt)
                If (Err.Number <> 0) Then
                    Dim mess = "Errore modificaHypFrom: Modifica SubAddress Linked " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "su:" & lista(5, punt) & " Descrizione " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore modificaHypFrom: Modifica SubAddress Linked " & lista(7, punt) & "!" & lista(8, punt) & lista(10, punt) & "su:" & lista(5, punt) & " Descrizione " & Err.Description)
                    Err.Clear()
                    Att = True
                Else
                    Dim mess = "modificaHyp da: " & subb & " a: " & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt)
                    WriteMia(ConsoleColor.Yellow, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Yellow
                    'Console.WriteLine("modificaHyp da: " & subb & " a: " & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt))
                    Uscita = Uscita & "<tr><td colspan=10><font color ='Fuchsia'> modifica Ipertesto da: " & subb & " a: " & lista(0, punt) & "'!" & lista(3, punt) & lista(1, punt) & "</font></td></tr>"
                End If
            Else
                Dim mess = "Errore modificaHypFrom: L'Hyperlink in modifica non corrisponde a quello atteso " & subb
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: L'Hyperlink in modifica non corrisponde a quello atteso " & subb)
            End If
            On Error GoTo 0
        End If
        If Not (InStr(1, lista(5, punt), "LinkSuSeStesso") > 0) Then 'e un link su se stesso
            On Error Resume Next
            objWorkSecondbook.Save
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: Salva Linked " & lista(5, punt) & " Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: Salva Linked " & lista(5, punt) & " Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
            On Error Resume Next
            objWorkSecondbook.Close
            If (Err.Number <> 0) Then
                Dim mess = "Errore modificaHypFrom: Chiudi Linked " & lista(5, punt) & " Descrizione " & Err.Description
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore modificaHypFrom: Chiudi Linked " & lista(5, punt) & " Descrizione " & Err.Description)
                Err.Clear()
                Att = True
            End If
            On Error GoTo 0
        End If  'e un link su se stesso
    End Function

    Function FindToFrom(ByRef sheet, ByRef lis, ByRef index, ByVal sh_name, ByVal typo, ByVal file, ByRef Att,
                        ByRef folderMaster, ByRef file_master, ByRef fso, ByRef Uscita, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug)
        'sheet = oggetto, lis Lista su cui registrare, index indice nella lista, sheet_name nome dell0 sheet, typo = 1 = TO 2 = FROM
        Dim cmt, colLocNum
        Dim rigaLocale = ""
        Dim colonnaLocale = ""
        Dim rigaLinked = ""
        Dim colLinkedNum = 0
        Dim colonnaLinked = ""
        Dim sheetLinked = ""
        Dim PosI, PosF, PosC, PosFi, PosSh, PosCo, PosRi, PosPSh, PosPCo, PosPRi, intermedio, temp, chiave
        Dim fileLinked = ""
        Dim fileCompletoLinked = ""
        Dim cartellaLinked = ""
        Dim cartellaRel = ""
        Dim file_completo = ""
        Dim pri = ""
        Dim pco = ""
        Dim psh = ""
        Dim indirizzo, hypAddress
        'HyFlink#TO#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink tipo 1
        'HyFlink#FR#Pos=PS=xx#PC=jj#PR=rr#cartella=C:\xxxxxx\yyyyy#file=LivelliAutonomia#S=xx#C=jj#R=rr#HyElink tipo 2
        'HyFlink#XL#Pos=PS=xx#PC=jj#PR=rr#Punt=S=xx#C=jj#R=rr#HyElink tipo 3
        '12345678901234567890
        If (typo = 1) Then
            chiave = "O" 'TO
        Else
            If (typo = 3) Then
                chiave = "L" ' TL and FL
            Else
                chiave = "R" ' FR
            End If
        End If
        For Each cmt In sheet.Comments
            'WScript.Echo "Loop:"  &  index
            PosI = InStr(1, cmt.text, "HyFlink#", 1)
            PosF = InStr(1, cmt.text, "#HyElink", 1)
            If (PosI > 0) Then 'è un link
                'indirizzo = cmt.Parent.Address
                'indirizzo = replace(indirizzo,"$","")
                'hypAddress = sheet.Range(indirizzo).Hyperlinks(1).Address
                SeparaRigheColonne(cmt.Parent.Address(0, 0), rigaLocale, colonnaLocale)
                colLocNum = calcolaColonna(colonnaLocale, Att, ofile, oOut, Fdebug)
                If (Mid(cmt.text, PosI + 9, 1) = chiave) Then  'è un link chiave
                    PosC = InStr(1, cmt.text, "#cartella=", 1)
                    PosFi = InStr(1, cmt.text, "#file=", 1)
                    If (chiave = "L") Then
                        PosSh = InStr(1, cmt.text, "=S=", 1)
                        cartellaLinked = "LinkSuSeStesso"
                    Else
                        PosSh = InStr(1, cmt.text, "#S=", 1)
                    End If
                    PosCo = InStr(1, cmt.text, "#C=", 1)
                    PosRi = InStr(1, cmt.text, "#R=", 1)
                    PosPSh = InStr(1, cmt.text, "=PS=", 1)
                    PosPCo = InStr(1, cmt.text, "#PC=", 1)
                    PosPRi = InStr(1, cmt.text, "#PR=", 1)
                    If (PosPRi > 0) Then
                        intermedio = Mid(cmt.text, PosPRi + 4, Len(cmt.text) - (PosPRi + 4))
                        pri = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link lo sheet di posizione
                    Else
                        If (Fdebug) Then
                            Dim mess = "Attenzione HyFlink senza sheet di posizione " & cmt.text
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'onsole.WriteLine("Attenzione HyFlink senza sheet di posizione " & cmt.text)
                        End If
                        Att = True
                    End If
                    If (PosPSh > 0) Then
                        intermedio = Mid(cmt.text, PosPSh + 4, Len(cmt.text) - (PosPSh + 4))
                        psh = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link lo sheet di posizione
                    Else
                        If (Fdebug) Then
                            Dim mess = "Attenzione HyFlink senza sheet di posizione " & cmt.text
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Attenzione HyFlink senza sheet di posizione " & cmt.text)
                        End If
                    End If
                    If (PosPCo > 0) Then
                        intermedio = Mid(cmt.text, PosPCo + 4, Len(cmt.text) - (PosPCo + 4))
                        pco = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link lo sheet di posizione
                    Else
                        If (Fdebug) Then
                            Dim mess = "Attenzione HyFlink senza sheet di posizione " & cmt.text
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Attenzione HyFlink senza Colonna di posizione " & cmt.text)
                        End If
                    End If
                    If Not (chiave = "L") Then
                        If (PosC > 0) Then
                            intermedio = Mid(cmt.text, PosC + 10, Len(cmt.text) - (PosC + 10))
                            cartellaLinked = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link il nome della cartellaLinked
                            'controllo se indirizzamento relativo e converto
                            If (InStr(1, cartellaLinked, "..") > 0) Then
                                cartellaRel = cartellaLinked
                                cartellaLinked = creaPathCompleto(cartellaRel, folderMaster)
                                If (InStr(1, cartellaLinked, "\\") > 0) Then
                                    cartellaLinked = "\\" + Replace(cartellaLinked, "\\", "\", 3)
                                End If
                            Else
                                If (StrComp(cartellaLinked, ".") = 0) Then
                                    'la cartella è relativa ed è quella di folderMaster
                                    cartellaRel = cartellaLinked
                                    cartellaLinked = creaPathCompleto(cartellaRel, folderMaster)
                                End If
                            End If
                        Else
                            If (Fdebug) Then
                                Dim mess = "Attenzione HyFlink senza cartella" & cmt.text
                                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                                'Console.ForegroundColor = ConsoleColor.Red
                                'Console.WriteLine("Attenzione HyFlink senza cartella" & cmt.text)
                            End If
                            Att = True
                        End If

                        If (PosFi > 0) Then
                            intermedio = Mid(cmt.text, PosFi + 6, Len(cmt.text) - (PosFi + 6))
                            fileLinked = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link il nome del fileLinked
                        Else
                            If (Fdebug) Then
                                Dim mess = "Attenzione HyFlink senza file" & cmt.text
                                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                                'Console.ForegroundColor = ConsoleColor.Red
                                'Console.WriteLine("Attenzione HyFlink senza file" & cmt.text)
                            End If
                            Att = True
                        End If
                    End If
                    If (PosSh > 0) Then
                        intermedio = Mid(cmt.text, PosSh + 3, Len(cmt.text) - (PosSh + 3))
                        sheetLinked = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link il nome del sheetLinked
                    Else
                        If (Fdebug) Then
                            Dim mess = "Attenzione HyFlink senza Sheet" & cmt.text
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Attenzione HyFlink senza Sheet" & cmt.text)
                        End If
                        Att = True
                    End If
                    If (PosCo > 0) Then
                        intermedio = Mid(cmt.text, PosCo + 3, Len(cmt.text) - (PosCo + 3))
                        colonnaLinked = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link la colonnaLinked
                        colLinkedNum = calcolaColonna(colonnaLinked, Att, ofile, oOut, Fdebug)
                    Else
                        If (Fdebug) Then
                            Dim mess = "Attenzione HyFlink senza Colonna" & cmt.text
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Attenzione HyFlink senza Colonna" & cmt.text)
                        End If
                        Att = True
                    End If
                    If (PosRi > 0) Then
                        intermedio = Mid(cmt.text, PosRi + 3, Len(cmt.text) - (PosRi + 3))
                        rigaLinked = Mid(intermedio, 1, InStr(1, intermedio, "#", 1) - 1) 'estrae dal link la rigaLinked
                    Else
                        If (Fdebug) Then
                            Dim mess = "Attenzione HyFlink senza Riga " & cmt.text
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Attenzione HyFlink senza Riga " & cmt.text)
                        End If
                        Att = True
                    End If
                    If (StrComp(cartellaLinked, "LinkSuSeStesso") = 0) Then
                        cartellaLinked = "LinkSuSeStesso"
                        file_completo = "LinkSuSeStesso"
                    Else
                        If (fso.FolderExists(cartellaLinked)) Then
                            On Error Resume Next
                            fileCompletoLinked = cercaFile(fileLinked, cartellaLinked, fso)
                            If (Err.Number <> 0) Then
                                Dim mess = "Errore cercaFile Descrizione " & Err.Description
                                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                                'Console.ForegroundColor = ConsoleColor.Red
                                'Console.WriteLine("Errore cercaFile Descrizione " & Err.Description)
                                Err.Clear()
                                Att = True
                            End If
                            On Error GoTo 0
                            file_completo = cartellaLinked & "\" & fileCompletoLinked
                        Else
                            Att = True
                            Dim mess = "Attenzione DIR-NON-VALIDA: " & cartellaLinked & " Link:" & cmt.text & " da:" & cmt.Parent.Address(0, 0)
                            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                            'Console.ForegroundColor = ConsoleColor.Red
                            'Console.WriteLine("Attenzione DIR-NON-VALIDA: " & cartellaLinked & " Link:" & cmt.text & " da:" & cmt.Parent.Address(0, 0))
                            Uscita = Uscita & "<tr><td colspan=10><font color=red> Attenzione la cartella " & cartellaLinked & " Link:" & cmt.text & " da:" & cmt.Parent.Address(0, 0) & " non esiste</font></td></tr>"
                        End If
                    End If
                    lis(0, index) = sh_name
                    lis(1, index) = rigaLocale
                    lis(2, index) = colLocNum
                    lis(3, index) = colonnaLocale
                    lis(4, index) = cartellaRel
                    lis(5, index) = file_completo
                    lis(6, index) = sheetLinked & "!" & colonnaLinked & rigaLinked
                    lis(7, index) = sheetLinked 'devo usare il nome
                    lis(8, index) = colonnaLinked
                    lis(9, index) = colLinkedNum
                    lis(10, index) = rigaLinked
                    lis(11, index) = fileLinked
                    lis(12, index) = -1   'link con listahyp
                    lis(13, index) = -1   'link con listaFlink
                    lis(14, index) = psh  'Posizione sheet      La posizione serve ad evidenziar lo spostamento
                    lis(15, index) = pco  'Posizione Colonna (lettere)
                    lis(16, index) = pri  'Posizione riga
                    lis(17, index) = estraiParteIniziale(file, cmt, "COMMENTO", fileMaster, ofile, oOut, Fdebug)  'parte iniziale file di appartenenza Il cmt al posto di hyp non so se funziona
                    index = index + 1
                End If 'è un link HyFlink#TO#
            End If 'è un link
        Next
    End Function

    Function incrociaFlinkHyp(ByRef indice_hyp, ByRef listahyp, ByRef indice, ByRef listaFlink, ByRef indice_Lflink, ByRef listaLinkedFlink, ByRef indice_lhyp, ByRef listaLinkedhyp)
        'Crea i link fra la listahyp e le liste listaFlink e listaLinkedFlink
        'Cerca anche incroci interni a listahyp nel caso di link su se stesso
        Dim y, i, x
        For y = 0 To indice_hyp - 1
            For x = 0 To indice_hyp - 1
                If ((StrComp(listahyp(5, y), "LinkSuSeStesso") = 0) And (StrComp(listahyp(5, x), "LinkSuSeStesso") = 0)) Then
                    If ((listahyp(14, x) = -1) And (listahyp(14, y) = -1)) Then
                        If ((listahyp(0, y) = listahyp(7, x)) And (listahyp(1, y) = listahyp(10, x)) And (listahyp(2, y) = listahyp(9, x)) And (listahyp(3, y) = listahyp(8, x)) _
                        And (listahyp(7, y) = listahyp(0, x)) And (listahyp(10, y) = listahyp(1, x)) And (listahyp(9, y) = listahyp(2, x)) And (listahyp(8, y) = listahyp(3, x))) Then
                            'sono effettivamente uno la controparte dell'altro
                            listahyp(14, x) = y
                            listahyp(14, y) = x
                            listahyp(16, x) = y
                            listahyp(16, y) = x
                        End If
                    End If
                End If
            Next
        Next
        For y = 0 To indice_hyp - 1 'Trovo i Flink che corrispondono con gli hyp (Anche quelli locali)
            For i = 0 To indice - 1
                If ((listahyp(0, y) = listaFlink(0, i)) And (listahyp(1, y) = listaFlink(1, i)) And (listahyp(2, y) = listaFlink(2, i))) Then
                    listaFlink(12, i) = y
                    listahyp(12, y) = i
                    If (listahyp(14, y) <> -1) Then
                        If Not (listahyp(15, y) <> -1) Then ' se ha già 9999 o 8888 glieli lascio
                            listahyp(15, y) = 7777 ' setto che l'Flink è già creato
                        End If
                    End If
                End If
                'end if
            Next
        Next
        For y = 0 To indice_hyp - 1 'Trovo i LinkedFlink che corrispondono con gli hyp
            For x = 0 To indice_Lflink - 1
                If (listahyp(14, y) = -1) Then
                    If ((listahyp(0, y) = listaLinkedFlink(7, x)) And (listahyp(1, y) = listaLinkedFlink(10, x)) And (listahyp(2, y) = listaLinkedFlink(9, x))) Then
                        listaLinkedFlink(12, x) = y
                        listahyp(13, y) = x
                    End If
                End If
            Next
        Next
        '------------------------------------------------------------------------------------------
        '-----Incrocia LinkedHyp
        For y = 0 To indice_lhyp - 1 'Trovo gli incroci tra hyp e LinkedHyp
            For x = 0 To indice_hyp - 1
                'if ((StrComp(listaLinkedhyp(5,y),"LinkSuSeStesso") = 0) and (StrComp(listaLinkedhyp(5,x),"LinkSuSeStesso") = 0)) then
                'if ((listaLinkedhyp(14,x) = -1) and (listaLinkedhyp(14,y) = -1)) then
                If ((listaLinkedhyp(0, y) = listahyp(7, x)) And (listaLinkedhyp(1, y) = listahyp(10, x)) And (listaLinkedhyp(2, y) = listahyp(9, x)) And (listaLinkedhyp(3, y) = listahyp(8, x)) _
                        And (listaLinkedhyp(7, y) = listahyp(0, x)) And (listaLinkedhyp(10, y) = listahyp(1, x)) And (listaLinkedhyp(9, y) = listahyp(2, x)) And (listaLinkedhyp(8, y) = listahyp(3, x))) Then
                    'sono effettivamente uno la controparte dell'altro
                    listaLinkedhyp(15, y) = x
                    listahyp(15, x) = y
                End If
                'end if
                'end if
            Next
        Next
        For y = 0 To indice_lhyp - 1 'Trovo i LinkedFlink che corrispondono con i LinkedHyp
            For i = 0 To indice_Lflink - 1
                If ((listaLinkedhyp(0, y) = listaLinkedFlink(0, i)) And (listaLinkedhyp(1, y) = listaLinkedFlink(1, i)) And (listaLinkedhyp(2, y) = listaLinkedFlink(2, i))) Then
                    'listaLinkedFlink(12,i) = y
                    listaLinkedhyp(13, y) = i
                    If (listaLinkedhyp(14, y) <> -1) Then
                        If (listaLinkedhyp(15, y) = -1) Then ' se ha già 9999 o 8888 glieli lascio
                            listaLinkedhyp(15, y) = 7777 ' setto che l'Flink è già creato
                        End If
                    End If
                End If
                'end if
            Next
        Next
        For y = 0 To indice_lhyp - 1 'Trovo i Flink che corrispondono con gli Linkedhyp
            For x = 0 To indice - 1
                If (listaLinkedhyp(14, y) = -1) Then
                    If ((listaLinkedhyp(0, y) = listaFlink(7, x)) And (listaLinkedhyp(1, y) = listaFlink(10, x)) And (listaLinkedhyp(2, y) = listaFlink(9, x))) Then
                        'listaLinkedFlink(13,x) = y
                        listaLinkedhyp(12, y) = x
                    End If
                End If
            Next
        Next
        '------------------------------------------------------------------------------------------
        For y = 0 To indice - 1 'Trovo i LinkedFlink che corrispondono con i Flink
            For x = 0 To indice_Lflink - 1
                If ((listaFlink(0, y) = listaLinkedFlink(7, x)) And (listaFlink(1, y) = listaLinkedFlink(10, x)) And (listaFlink(2, y) = listaLinkedFlink(9, x))) Then
                    listaLinkedFlink(13, x) = y
                    listaFlink(13, y) = x
                End If
            Next
        Next
        incrociaFlinkHyp = 0
    End Function

    Function cercaControparte(ByVal punt, ByRef indice_hyp, ByRef listahyp)
        'cerco la controparte che abbia come locazione il mio puntamento
        Dim z
        For z = 0 To indice_hyp - 1
            If ((listahyp(14, z) = -1) And (listahyp(16, z) = -1)) Then
                If ((listahyp(0, z) = listahyp(7, punt)) And (listahyp(1, z) = listahyp(10, punt)) _
                And (listahyp(2, z) = listahyp(9, punt)) And (listahyp(3, z) = listahyp(8, punt))) Then
                    cercaControparte = z
                    Exit Function
                End If
            End If
        Next
        cercaControparte = -1
    End Function

    Function cercalinkedFlink(y, ByRef indice_Lflink, ByRef listaLinkedFlink, ByRef listaFlink)
        Dim k
        cercalinkedFlink = -1
        For k = 0 To indice_Lflink - 1
            If ((listaLinkedFlink(0, k) = listaFlink(7, y)) And
            (listaLinkedFlink(3, k) = listaFlink(8, y)) And
            (listaLinkedFlink(2, k) = listaFlink(9, y)) And
            (listaLinkedFlink(1, k) = listaFlink(10, y))) Then
                cercalinkedFlink = k
                Exit For
            End If
        Next
    End Function


    Function cercaFlink(ByVal sh, ByVal r, ByVal c, ByRef indice, ByRef listaFlink)
        Dim k
        cercaFlink = -1
        For k = 0 To indice - 1
            If ((listaFlink(0, k) = sh) And
            (listaFlink(1, k) = r) And
            (listaFlink(3, k) = c)) Then
                cercaFlink = k
                Exit For
            End If
        Next
    End Function

    Function cercaFlinkDaHyp(ByVal p_i, ByVal sh, ByVal r, ByVal c, lista, ByVal index)
        Dim k
        cercaFlinkDaHyp = -1
        For k = 0 To index - 1
            If ((lista(7, k) = sh) And
            (lista(10, k) = r) And
            (UCase(lista(11, k)) = UCase(p_i)) And
            (lista(8, k) = c)) Then
                cercaFlinkDaHyp = k
                Exit For
            End If
        Next
    End Function

    Function clearLista(ByRef lista, ByVal ind, ByVal k)
        Dim z, x
        For z = 0 To ind
            For x = 0 To k
                lista(x, z) = ""
            Next
        Next
        clearLista = 0
    End Function

    Function estraiParteIniziale(address, hyp, tipo, ByRef fileMaster, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim pos, temp, mess

        If (Len(address) = 0) Then
            estraiParteIniziale = "LinkSuSeStesso"
        Else
            pos = InStrRev(address, "\")
            If (pos = 0) Then
                pos = InStrRev(address, "/")
            End If
            If (pos <> 0) Then
                temp = Mid(address, pos + 1, Len(address) - pos)
            Else
                If (InStr(1, address, fileMaster) > 0) Then
                    estraiParteIniziale = "LinkSuSeStesso"
                    Exit Function
                Else
                    If (StrComp(tipo, "HYPER") = 0) Then
                        mess = "Attenzione Indirizzo del hyperlink non secondo standard : " & address & " Posizione: " & hyp.Parent.Address(0, 0)
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Attenzione Indirizzo del hyperlink non secondo standard : " & address & " Posizione: " & hyp.Parent.Address(0, 0))
                    Else
                        mess = "Attenzione Indirizzo del Commento non secondo standard : " & address & " Posizione: " & hyp.Parent.Address(0, 0)
                        WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                        'Console.ForegroundColor = ConsoleColor.Red
                        'Console.WriteLine("Attenzione Indirizzo del Commento non secondo standard : " & address & " Posizione: " & hyp.Parent.Address(0, 0))
                    End If
                    mess = "uscita dal programma per errore"
                    WriteMia(ConsoleColor.Magenta, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Magenta
                    'Console.WriteLine("uscita dal programma per errore")
                    Console.ForegroundColor = ConsoleColor.White
                    ofile.Close()
                    End
                End If
            End If
            pos = InStr(1, temp, "'")
            If (pos > 0) Then 'se non c'è l'apice mantengo tutto il nome del file
                estraiParteIniziale = Mid(temp, 1, pos - 1)
            Else
                estraiParteIniziale = Mid(temp, 1, InStr(1, temp, ".") - 1)
            End If
        End If
    End Function

    Function outListe(ByRef out, ByRef indice_hyp, ByRef indice_lhyp, ByRef indice, ByRef listahyp,
                      ByRef listaLinkedhyp, ByRef listaFlink, ByRef listaLinkedFlink, ByRef indice_Lflink, ByRef n_fileIn, ByRef listaFile)
        out = "##################Lista HyperLink########################" & vbCrLf
        For y = 0 To indice_hyp - 1 Step 1
            out = out & "hhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhh" & vbCrLf
            out = out & "--  Indice  -------------------:" & y & vbCrLf
            out = out & "0)  Sheet locale               :" & listahyp(0, y) & vbCrLf
            out = out & "1)  Riga locale                :" & listahyp(1, y) & vbCrLf
            out = out & "2)  Colonna num                :" & listahyp(2, y) & vbCrLf
            out = out & "3)  Colonna lett               :" & listahyp(3, y) & vbCrLf
            out = out & "4)  File link (rel)            :" & listahyp(4, y) & vbCrLf
            out = out & "5)  File link comp             :" & listahyp(5, y) & vbCrLf
            out = out & "6)  SubAddress                 :" & listahyp(6, y) & vbCrLf
            out = out & "7)  Sheet name link            :" & listahyp(7, y) & vbCrLf
            out = out & "8)  Col link lett              :" & listahyp(8, y) & vbCrLf
            out = out & "9)  Col link num               :" & listahyp(9, y) & vbCrLf
            out = out & "10) Riga linked                :" & listahyp(10, y) & vbCrLf
            out = out & "11) File link parziale         :" & listahyp(11, y) & vbCrLf
            out = out & "12) Link su listaFlink         :" & listahyp(12, y) & vbCrLf
            out = out & "13) Link su listaLinkedFlink   :" & listahyp(13, y) & vbCrLf
            out = out & "14) Link su se stesso          :" & listahyp(14, y) & vbCrLf
            out = out & "15) Punt corr. hyp o LinkedHyp :" & listahyp(15, y) & vbCrLf
            out = out & "16) situaz link                :" & listahyp(16, y) & vbCrLf
        Next
        out = out & vbCrLf
        out = out & "##################Lista LinkedHyperLink########################" & vbCrLf
        For y = 0 To indice_lhyp - 1 Step 1
            out = out & "HHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH" & vbCrLf
            out = out & "--  Indice  -------------------:" & y & vbCrLf
            out = out & "0)  Sheet locale               :" & listaLinkedhyp(0, y) & vbCrLf
            out = out & "1)  Riga locale                :" & listaLinkedhyp(1, y) & vbCrLf
            out = out & "2)  Colonna num                :" & listaLinkedhyp(2, y) & vbCrLf
            out = out & "3)  Colonna lett               :" & listaLinkedhyp(3, y) & vbCrLf
            out = out & "4)  File link (rel)            :" & listaLinkedhyp(4, y) & vbCrLf
            out = out & "5)  File link comp             :" & listaLinkedhyp(5, y) & vbCrLf
            out = out & "6)  SubAddress                 :" & listaLinkedhyp(6, y) & vbCrLf
            out = out & "7)  Sheet name link            :" & listaLinkedhyp(7, y) & vbCrLf
            out = out & "8)  Col link lett              :" & listaLinkedhyp(8, y) & vbCrLf
            out = out & "9)  Col link num               :" & listaLinkedhyp(9, y) & vbCrLf
            out = out & "10) Riga linked                :" & listaLinkedhyp(10, y) & vbCrLf
            out = out & "11) File link parziale         :" & listaLinkedhyp(11, y) & vbCrLf
            out = out & "12) Link su listaFlink         :" & listaLinkedhyp(12, y) & vbCrLf
            out = out & "13) Link su listaLinkedFlink   :" & listaLinkedhyp(13, y) & vbCrLf
            out = out & "14) Link su se stesso          :" & listaLinkedhyp(14, y) & vbCrLf
            out = out & "15) Punt corr. hyp o LinkedHyp :" & listaLinkedhyp(15, y) & vbCrLf
            out = out & "16) situaz link                :" & listaLinkedhyp(16, y) & vbCrLf
        Next
        out = out & vbCrLf
        out = out & "##################Lista Flink########################" & vbCrLf
        For y = 0 To indice - 1 Step 1
            out = out & "fffffffffffffffffffffffffffffffffffffffffffffffffffffff" & vbCrLf
            out = out & "--  Indice  -------------------:" & y & vbCrLf
            out = out & "0)  Sheet locale               :" & listaFlink(0, y) & vbCrLf
            out = out & "1)  Riga locale                :" & listaFlink(1, y) & vbCrLf
            out = out & "2)  Colonna num                :" & listaFlink(2, y) & vbCrLf
            out = out & "3)  Colonna lett               :" & listaFlink(3, y) & vbCrLf
            out = out & "4)  Cartella Link (rel)        :" & listaFlink(4, y) & vbCrLf
            out = out & "5)  File link comp             :" & listaFlink(5, y) & vbCrLf
            out = out & "6)  sim subAddress             :" & listaFlink(6, y) & vbCrLf
            out = out & "7)  Sheet name link            :" & listaFlink(7, y) & vbCrLf
            out = out & "8)  Col link lett              :" & listaFlink(8, y) & vbCrLf
            out = out & "9)  Col link num               :" & listaFlink(9, y) & vbCrLf
            out = out & "10) Riga linked                :" & listaFlink(10, y) & vbCrLf
            out = out & "11) File link parziale         :" & listaFlink(11, y) & vbCrLf
            out = out & "12) Link su listahyp           :" & listaFlink(12, y) & vbCrLf
            out = out & "13) Link su listaLinkedFlink   :" & listaFlink(13, y) & vbCrLf
            out = out & "14) Posizione sheet            :" & listaFlink(14, y) & vbCrLf
            out = out & "15) Posizione colonna          :" & listaFlink(15, y) & vbCrLf
            out = out & "16) Posizione riga             :" & listaFlink(16, y) & vbCrLf
            out = out & "17) ParteIniNomeFileAppartenen.:" & listaFlink(17, y) & vbCrLf
        Next
        out = out & vbCrLf
        out = out & "##################Lista LinkedFlink########################" & vbCrLf
        For y = 0 To indice_Lflink - 1 Step 1
            out = out & "lflflflflflflflflflflflflflflflflflflflflflflflf" & vbCrLf
            out = out & "--  Indice  -------------------:" & y & vbCrLf
            out = out & "0)  Sheet locale               :" & listaLinkedFlink(0, y) & vbCrLf
            out = out & "1)  Riga locale                :" & listaLinkedFlink(1, y) & vbCrLf
            out = out & "2)  Colonna num                :" & listaLinkedFlink(2, y) & vbCrLf
            out = out & "3)  Colonna lett               :" & listaLinkedFlink(3, y) & vbCrLf
            out = out & "4)  Cartella Link  (rel)       :" & listaLinkedFlink(4, y) & vbCrLf
            out = out & "5)  File link comp             :" & listaLinkedFlink(5, y) & vbCrLf
            out = out & "6)  sim subAddress             :" & listaLinkedFlink(6, y) & vbCrLf
            out = out & "7)  Sheet name link            :" & listaLinkedFlink(7, y) & vbCrLf
            out = out & "8)  Col link lett              :" & listaLinkedFlink(8, y) & vbCrLf
            out = out & "9)  Col link num               :" & listaLinkedFlink(9, y) & vbCrLf
            out = out & "10) Riga linked                :" & listaLinkedFlink(10, y) & vbCrLf
            out = out & "11) File link parziale         :" & listaLinkedFlink(11, y) & vbCrLf
            out = out & "12) Link su listahyp			:" & listaLinkedFlink(12, y) & vbCrLf
            out = out & "13) Link su listaFlink			:" & listaLinkedFlink(13, y) & vbCrLf
            out = out & "14) Posizione sheet            :" & listaLinkedFlink(14, y) & vbCrLf
            out = out & "15) Posizione colonna          :" & listaLinkedFlink(15, y) & vbCrLf
            out = out & "16) Posizione riga             :" & listaLinkedFlink(16, y) & vbCrLf
            out = out & "17) ParteIniNomeFileAppartenen.:" & listaLinkedFlink(17, y) & vbCrLf
        Next
        out = out & vbCrLf
        out = out & "################Lista Files Linked########################" & vbCrLf
        For y = 0 To n_fileIn - 1 Step 1
            out = out & "--------------------------------------------------------" & vbCrLf
            out = out & "---Indice  --------------------:" & y & vbCrLf
            out = out & "   File                        :" & listaFile(y) & vbCrLf
        Next
        outListe = out
    End Function

    Function separSheetCR(scr, ByRef s_name, ByRef riga, ByRef colonna, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim piece
        If (InStr(1, scr, "!") > 0) Then
            piece = Split(scr, "!")
            s_name = piece(0)
            SeparaRigheColonne(piece(1), riga, colonna)
        Else
            Dim mess = "separSheetCR: hyplink estraneo al programma " & scr
            WriteMia(ConsoleColor.Cyan, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Cyan
            'Console.WriteLine("separSheetCR: hyplink estraneo al programma " & scr)
        End If
        separSheetCR = 0
    End Function

    Function creaPathCompleto(address, master)
        Dim i, id_n, piece, out, p, temp, temp1, n_back, pos
        Dim lastPos = ""
        pos = InStr(1, address, "..")
        If (InStr(1, address, "..") > 0) Then
            ' devo vedere quante dir devo risalire
            n_back = 1
            Do While pos > 0
                pos = InStr(pos + 2, address, "..")
                If (pos > 0) Then
                    n_back = n_back + 1
                    lastPos = pos
                End If
            Loop
            piece = Split(master, "\")
            id_n = UBound(piece)
            out = piece(0)
            For i = 1 To (id_n - n_back)
                out = out & "\" & piece(i)
            Next
            temp1 = Mid(address, lastPos + 2)
            p = InStr(1, temp1, "\")
            If (p <= 0) Then
                p = InStr(1, temp1, "/")
            End If
            temp = Mid(temp1, p, Len(temp1) - (p - 1))
            out = out & temp
        Else

            If ((InStr(1, address, ":\") > 0) Or (InStr(1, address, "\\") > 0)) Then
                out = address
            Else
                If (StrComp(address, ".") = 0) Then
                    out = master
                Else
                    out = master & "\" & address
                End If
            End If
        End If
        creaPathCompleto = Replace(out, "/", "\")
    End Function

    Function leggiINI(ByRef folderOutput, ByRef fileOutput, ByRef repOutput, ByRef debug, ByRef settaFont, ByRef dimFont, ByRef fileINI)
        Dim objFileToRead, linea, x, debug_st, temp
        objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(fileINI, 1)
        For x = 1 To 6
            linea = objFileToRead.ReadLine()
            If (InStr(1, linea, "cartella=", 1) > 0) Then
                folderOutput = Mid(linea, 10, Len(linea) - 9)
            Else
                If (InStr(1, linea, "file=", 1) > 0) Then
                    fileOutput = Mid(linea, 6, Len(linea) - 5)
                Else
                    If (InStr(1, linea, "rapporto=", 1) > 0) Then
                        repOutput = Mid(linea, 10, Len(linea) - 9)
                    Else
                        If (InStr(1, linea, "debug=", 1) > 0) Then
                            debug_st = Mid(linea, 7, Len(linea) - 6)
                            If ((InStr(1, LCase(debug_st), "si", 1) > 0) Or (InStr(1, LCase(debug_st), "yes", 1))) Then
                                debug = True
                            Else
                                debug = False
                            End If
                        Else
                            If (InStr(1, linea, "settaFont=", 1) > 0) Then
                                temp = Mid(linea, 11, Len(linea) - 10)
                                If (StrComp(temp, "si") = 0) Then
                                    settaFont = True
                                End If
                            Else
                                If (InStr(1, linea, "dimFont=", 1) > 0) Then
                                    dimFont = Mid(linea, 9, Len(linea) - 8)
                                End If 'dimFont
                            End If 'settaFont
                        End If 'debug

                    End If 'rapporto
                End If 'file
            End If 'cartelle
        Next
        objFileToRead.Close
        objFileToRead = Nothing
        leggiINI = 0
    End Function

    Function calcolaColonna(ByVal nome_colonna, ByRef Att, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim tot, ci, c
        tot = 0
        ci = 0
        If (Len(nome_colonna) > 1) Then
            If (Len(nome_colonna) > 2) Then
                If (Len(nome_colonna) > 3) Then
                    Dim mess = "Considero le colonne solo fino alla terza lettera"
                    WriteMia(ConsoleColor.DarkGreen, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.DarkGreen
                    'Console.WriteLine("Considero le colonne solo fino alla terza lettera")
                    Att = True
                    calcolaColonna = "+ZZZ"
                    Exit Function
                End If
                'WScript.Echo "siamo a 3"
                c = Mid(nome_colonna, 1, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                ci = ci * 676
                tot = tot + ci
                'WScript.Echo "1tot:" & tot
                '---------------------------
                c = Mid(nome_colonna, 2, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                ci = ci * 26
                tot = tot + ci
                'WScript.Echo "2tot:" & tot
                '---------------------------
                c = Mid(nome_colonna, 3, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                tot = tot + ci
                'WScript.Echo "3tot:" & tot
            Else
                c = Mid(nome_colonna, 1, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                ci = ci * 26
                tot = tot + ci
                '---------------------------
                c = Mid(nome_colonna, 2, 1)
                ci = CInt(Asc(UCase(c)) - 64)
                tot = tot + ci
            End If
        Else
            If (nome_colonna = "") Then
                Dim mess = "Errore Apertura Nome colonna vuoto"
                WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                'Console.ForegroundColor = ConsoleColor.Red
                'Console.WriteLine("Errore Apertura Nome colonna vuoto")
                Att = True
            End If
            tot = CInt(Asc(UCase(nome_colonna)) - 64)
        End If
        calcolaColonna = tot
    End Function

    Function popolaListaFile(ByRef lf, ByVal ind, ByRef listahyp) 'Crea la lista dei file di input per evitare di leggerli due volte
        Dim x, y, ce
        y = 0
        For x = 0 To ind - 1 Step 1
            ce = thereis(listahyp(5, x), lf, y)
            If Not (ce) Then
                lf(y) = listahyp(5, x)
                y = y + 1
            End If
        Next
        popolaListaFile = y ' esporto il livello al quale è arrivata la listaFile
    End Function

    Function thereis(ByVal ff, ByRef lis, ByVal upto)
        Dim i
        For i = 0 To upto Step 1
            If (ff = lis(i)) Then 'se la dir c'è
                thereis = True
                Exit Function
            End If
        Next
        thereis = False
    End Function

    Function SeparaRigheColonne(ByRef indirizzo, ByRef riga, ByRef colonna)
        Dim c, i

        For i = 1 To Len(indirizzo)
            c = Mid(indirizzo, i, 1)
            If (IsNumeric(c)) Then
                Exit For
            End If
        Next
        'WScript.Echo "Numerico da "  &  i
        colonna = Mid(indirizzo, 1, i - 1)
        riga = Mid(indirizzo, i, Len(indirizzo))
        SeparaRigheColonne = 0
    End Function

    Function cercaFile(ByVal patt, ByVal folder, ByVal fso) 'pattern da cercare e directory
        Dim filenamecompleto, list
        Dim f, parte
        Dim objFolder
        filenamecompleto = "NULLA"
        list = CreateObject("ADOR.Recordset")
        list.Fields.Append("name", 200, 255)
        list.Fields.Append("date", 7)
        list.Open

        'list.MoveFirst
        'Do Until list.EOF
        '  WScript.Echo list("date").Value  &  vbTab  &  list("name").Value
        '  list.MoveNext
        'Loop

        'WScript.Echo "Folder in cerca:"  &  folder
        objFolder = fso.GetFolder(folder)
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
            parte = Left(LCase(list("name").Value), Len(patt)) 'preleva i primi caratteri
            If Not (InStr(parte, patt) = 0) Then
                filenamecompleto = LCase(list("name").Value)
                Exit Do
            End If
            list.MoveNext
        Loop
        cercaFile = filenamecompleto
    End Function

    Function estraiFolderDaAddress(ByVal add, ByVal iniz, ByRef folderMaster)
        Dim Pos, i, temp, temp2, sotto_dir
        If (Len(add) = 0) Then
            estraiFolderDaAddress = "LinkSuSeStesso"
            Exit Function
        End If
        If (StrComp(iniz, "LinkSuSeStesso") = 0) Then
            estraiFolderDaAddress = "LinkSuSeStesso"
            Exit Function
        End If
        If ((InStr(1, add, ":\") > 0) Or (InStr(1, add, "\\") > 0)) Then
            'dir completa
            Pos = InStr(1, add, iniz)
            estraiFolderDaAddress = Mid(add, 1, Pos - 2) 'toglie anche la \
        Else
            If (InStr(1, add, "..") > 0) Then
                Pos = InStr(1, add, "\")
                If (Pos = 0) Then
                    Pos = InStr(1, add, "/")
                End If
                temp = Mid(add, Pos + 1, Len(add) - Pos)
                sotto_dir = 0
                While (InStr(1, temp, "..") > 0)
                    Pos = InStr(1, temp, "\")
                    If (Pos = 0) Then
                        Pos = InStr(1, temp, "/")
                    End If
                    temp = Mid(temp, Pos + 1, Len(temp) - Pos)
                    sotto_dir = sotto_dir + 1
                End While
                Pos = InStr(1, temp, iniz)
                temp = Mid(temp, 1, Pos - 2) 'toglie anche la \
                temp2 = folderMaster
                For i = 0 To sotto_dir
                    Pos = InStrRev(temp2, "\")
                    If (Pos = 0) Then
                        Pos = InStrRev(temp2, "/")
                    End If
                    temp2 = Mid(temp2, 1, Pos - 1)
                Next
                estraiFolderDaAddress = temp2 & "\" & temp
            Else
                If Not (InStr(1, add, "\") > 0) Then
                    'contiene solo il nome file la dir è quella del Master
                    estraiFolderDaAddress = folderMaster
                End If
            End If
        End If
    End Function

    Function scriviSu(ByVal nome, ByVal dato)
        Const toWrite = 2
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.OpenTextFile(nome, toWrite, True, 0)
        f.Write(dato)
        f.close
        f = Nothing
        fs = Nothing
        scriviSu = 0
    End Function

    Function appendiA(ByVal nome, ByVal dato)
        Const toAppend = 8
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.OpenTextFile(nome, toAppend, True, 0)
        f.Write(dato)
        f.close
        f = Nothing
        fs = Nothing
        appendiA = 0
    End Function

    Function newFlinkHypLink(objSheet, lis, idx, p_sh, p_cn, p_c, p_r, ByRef objWorkbook, ByRef Att, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim cartella, file, commento, comm, objCommento

        If (InStrRev(lis(4, idx), "..") > 0) Then
            'Relativo
            cartella = lis(4, idx)
            If (InStrRev(lis(4, idx), ".xls") > 0) Then
                cartella = Mid(cartella, 1, InStrRev(cartella, "\") - 1)
            End If
        Else
            If (InStrRev(lis(4, idx), "\") = 0) Then
                cartella = Mid(lis(4, idx), 1, InStrRev(lis(4, idx), "/") - 1)
            Else
                cartella = Mid(lis(4, idx), 1, InStrRev(lis(4, idx), "\") - 1)
            End If
        End If
        On Error Resume Next
        objSheet = objWorkbook.Worksheets(lis(0, idx)) 'mi setto sul giusto sheet
        If (Err.Number <> 0) Then
            Dim mess = "Errore select sheet newFlinkHypLink:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore select sheet newFlinkHypLink:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        '                  (riga,col)
        'objSheet.Cells(lis(1,idx), lis(2,idx)).ClearComments
        On Error Resume Next
        comm = objSheet.Cells(lis(1, idx), lis(2, idx)).Comment.Text
        If (Err.Number <> 0) Then
            Dim mess = "Errore clear comment newFlinkHypLink:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore clear comment newFlinkHypLink:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
            'rimuovo il link
            commento = rimuoviLink(comm)
        Else
            commento = comm
        End If
        commento = commento & vbCrLf & "HyFlink#TO#Pos=PS=" & lis(0, idx) & "#PC=" & lis(3, idx) & "#PR=" & lis(1, idx) & "#cartella=" & cartella & "#file=" & lis(11, idx) & "#S=" & p_sh & "#C=" & p_c & "#R=" & p_r & "#HyElink"
        commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
        On Error Resume Next
        objSheet.Cells(lis(1, idx), lis(2, idx)).ClearComments
        If (Err.Number <> 0) Then
            Dim mess = "Errore clear comment newFlinkHypLink:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore clear comment newFlinkHypLink:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objSheet.Cells(lis(1, idx), lis(2, idx)).AddComment(commento)
        If (Err.Number <> 0) Then
            Dim mess = "Errore add comment newFlinkHypLink:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore add comment newFlinkHypLink:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        newFlinkHypLink = True
    End Function



    Function modificaLinkTo(objSheet, idx, chiave, contro, ByRef listaFlink, ByRef Att, ByRef objWorkbook,
                            ByRef listaLinkedFlink, ByRef ofile, ByRef oOut, ByRef Fdebug)
        'modificaLinkTo(objWorksheet, y, Controparte, "XL")
        Dim cartella, file, commento, comm, objCommento

        If (StrComp(listaFlink(5, idx), "LinkSuSeStesso") = 0) Then
            cartella = "LinkSuSeStesso"
            chiave = "XL" 'Il controllo viene fatto soltanto sulla seconda lettera che sia TL o FL fa poca differenza
        Else
            If (InStrRev(listaFlink(4, idx), "..") > 0) Then
                'Relativo
                cartella = listaFlink(4, idx)
                If (InStrRev(listaFlink(4, idx), ".xls") > 0) Then
                    cartella = Mid(cartella, 1, InStrRev(cartella, "\") - 1)
                End If
            Else
                If (InStrRev(listaFlink(4, idx), "\") = 0) Then
                    cartella = Mid(listaFlink(4, idx), 1, InStrRev(listaFlink(4, idx), "/") - 1)
                Else
                    cartella = Mid(listaFlink(4, idx), 1, InStrRev(listaFlink(4, idx), "\") - 1)
                End If
            End If
        End If
        On Error Resume Next
        objSheet = objWorkbook.Worksheets(listaFlink(0, idx)) 'mi setto sul giusto sheet
        If (Err.Number <> 0) Then
            Dim mess = "Errore select sheet modificaLinkTo:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore select sheet modificaLinkTo:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        '                  (riga,col)
        'objSheet.Cells(lis(1,idx), lis(2,idx)).ClearComments
        On Error Resume Next
        comm = objSheet.Cells(listaFlink(1, idx), listaFlink(2, idx)).Comment.Text
        If (Err.Number <> 0) Then
            Dim mess = "Errore clear comment modificaLinkTo:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore clear comment modificaLinkTo:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
            'rimuovo il link
            commento = rimuoviLink(comm)
        Else
            commento = comm
        End If
        If (StrComp(listaFlink(5, idx), "LinkSuSeStesso") = 0) Then
            'HyFlink#XL#Pos=S=xx#C=jj#R=rr#Punt=S=xx#C=jj#R=rr#HyElink
            'ATTENZIONE il puntamento che vado a scrivere su di me (reale) deve essere verso la posizione reale adesso della controparte cioè (0,controparte) (3,controparte) (1,controparte)
            commento = commento & vbCrLf & "HyFlink#" & chiave & "#Pos=PS=" & listaFlink(0, idx) & "#PC=" & listaFlink(3, idx) & "#PR=" & listaFlink(1, idx) & "#Punt=S=" & listaLinkedFlink(0, contro) & "#C=" & listaLinkedFlink(3, contro) & "#R=" & listaLinkedFlink(1, contro) & "#HyElink"
        Else
            commento = commento & vbCrLf & "HyFlink#" & chiave & "#Pos=PS=" & listaFlink(0, idx) & "#PC=" & listaFlink(3, idx) & "#PR=" & listaFlink(1, idx) & "#cartella=" & cartella & "#file=" & listaLinkedFlink(17, contro) & "#S=" & listaLinkedFlink(0, contro) & "#C=" & listaLinkedFlink(3, contro) & "#R=" & listaLinkedFlink(1, contro) & "#HyElink"
        End If
        commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
        On Error Resume Next
        objSheet.Cells(listaFlink(1, idx), listaFlink(2, idx)).ClearComments
        If (Err.Number <> 0) Then
            Dim mess = "Errore clear comment modificaLinkTo:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore clear comment modificaLinkTo:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        objSheet.Cells(listaFlink(1, idx), listaFlink(2, idx)).AddComment(commento)
        If (Err.Number <> 0) Then
            Dim mess = "Errore add comment modificaLinkTo:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore add comment modificaLinkTo:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        modificaLinkTo = True
    End Function



    Function modificaLinkToLocal(objSheet, idx, Rsh, Rco, Rri, ByRef objWorkbook, ByRef listaFlink, ByRef Att, ByRef ofile, ByRef oOut, ByRef Fdebug)
        'modificaLinkToLocal(objWorksheet, y, Controparte)
        Dim cartella, file, commento, comm, objCommento
        On Error Resume Next
        objSheet = objWorkbook.Worksheets(listaFlink(0, idx)) 'mi setto sul giusto sheet
        If (Err.Number <> 0) Then
            Dim mess = "Errore select sheet modificaLinkToLocal:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore select sheet modificaLinkToLocal:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        '                  (riga,col)
        'objSheet.Cells(lis(1,idx), lis(2,idx)).ClearComments
        On Error Resume Next
        comm = objSheet.Cells(listaFlink(1, idx), listaFlink(2, idx)).Comment.Text
        If (Err.Number <> 0) Then
            Dim mess = "Errore clear comment modificaLinkToLocal:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore clear comment modificaLinkToLocal:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
            'rimuovo il link
            commento = rimuoviLink(comm)
        Else
            commento = comm
        End If
        'HyFlink#XL#Pos=S=xx#C=jj#R=rr#Punt=S=xx#C=jj#R=rr#HyElink
        commento = commento & vbCrLf & "HyFlink#XL#Pos=PS=" & listaFlink(0, idx) & "#PC=" & listaFlink(3, idx) & "#PR=" & listaFlink(1, idx) & "#Punt=S=" & Rsh & "#C=" & Rco & "#R=" & Rri & "#HyElink"
        ' uso come POS posizione reale di idx e come puntatore la posizione reale della controparte S=Rsh C=Rco R=Rri
        commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
        On Error Resume Next
        objSheet.Cells(listaFlink(1, idx), listaFlink(2, idx)).ClearComments
        If (Err.Number <> 0) Then
            Dim mess = "Errore clear comment modificaLinkToLocal:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore clear comment modificaLinkToLocal:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        On Error Resume Next
        ' setto nella posizione reale il commento
        objSheet.Cells(listaFlink(1, idx), listaFlink(2, idx)).AddComment(commento)
        If (Err.Number <> 0) Then
            Dim mess = "Errore add comment modificaLinkToLocal:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore add comment modificaLinkToLocal:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        modificaLinkToLocal = True
    End Function


    Function scriviLinkTo(objSheet, lis, idx, chiave, ByRef objWorkbook, ByRef Att, ByRef ofile, ByRef oOut, ByRef Fdebug)
        Dim cartella, file, commento, comm, objCommento

        If (StrComp(lis(5, idx), "LinkSuSeStesso") = 0) Then
            cartella = "LinkSuSeStesso"
            chiave = "XL" 'Il controllo viene fatto soltanto sulla seconda lettera che sia TL o FL fa poca differenza
        Else
            If (InStrRev(lis(4, idx), "..") > 0) Then
                'Relativo
                cartella = lis(4, idx)
                If (InStrRev(lis(4, idx), ".xls") > 0) Then
                    If (InStrRev(cartella, "\") = 0) Then
                        cartella = Mid(cartella, 1, InStrRev(cartella, "/") - 1)
                    Else
                        cartella = Mid(cartella, 1, InStrRev(cartella, "\") - 1)
                    End If
                End If
            Else
                If (InStrRev(lis(4, idx), "\") = 0) Then
                    cartella = Mid(lis(4, idx), 1, InStrRev(lis(4, idx), "/") - 1)
                Else
                    cartella = Mid(lis(4, idx), 1, InStrRev(lis(4, idx), "\") - 1)
                End If
            End If
        End If
        On Error Resume Next
        objSheet = objWorkbook.Worksheets(lis(0, idx)) 'mi setto sul giusto sheet
        If (Err.Number <> 0) Then
            Dim mess = "Errore select sheet scriviLinkTo:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore select sheet scriviLinkTo:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        '                  (riga,col)
        'objSheet.Cells(lis(1,idx), lis(2,idx)).ClearComments
        On Error Resume Next
        comm = objSheet.Cells(lis(1, idx), lis(2, idx)).Comment.Text
        If (Err.Number <> 0) Then
            If (Err.Number <> 424) Then
                If (Err.Number <> 9) Then
                    comm = ""
                Else
                    Dim mess = "Errore clear comment scriviLinkTo s:" & lis(0, idx) & " r:" & lis(1, idx) & " c:" & lis(2, idx) & " err:" & Err.Number & " Descrizione: " & Err.Description
                    WriteMia(ConsoleColor.Red, mess, oOut, ofile)
                    'Console.ForegroundColor = ConsoleColor.Red
                    'Console.WriteLine("Errore clear comment scriviLinkTo s:" & lis(0, idx) & " r:" & lis(1, idx) & " c:" & lis(2, idx) & " err:" & Err.Number & " Descrizione: " & Err.Description)
                    Err.Clear()
                    Att = True
                End If
            Else
                Err.Clear()
                'Console.WriteLine("Errore non grave: su commento scriviLinkTo s:" & lis(0,idx) & " r:" & lis(1,idx) & " c:" & lis(2,idx) & " err:" &  Err.Number  &  " Descrizione: "  &  Err.Description & "" & vbCrLf
                'erore soltanto perchè il commento non c'è ancora.
            End If
        End If
        On Error GoTo 0
        If (InStr(1, comm, "HyFlink#") > 0) Then 'c'è già un link devo toglierlo
            'rimuovo il link
            commento = rimuoviLink(comm)
        Else
            commento = comm
        End If
        If (StrComp(lis(5, idx), "LinkSuSeStesso") = 0) Then
            'HyFlink#XL#Pos=S=xx#C=jj#R=rr#Punt=S=xx#C=jj#R=rr#HyElink
            commento = commento & vbCrLf & "HyFlink#" & chiave & "#Pos=PS=" & lis(0, idx) & "#PC=" & lis(3, idx) & "#PR=" & lis(1, idx) & "#Punt=S=" & lis(7, idx) & "#C=" & lis(8, idx) & "#R=" & lis(10, idx) & "#HyElink"
        Else
            commento = commento & vbCrLf & "HyFlink#" & chiave & "#Pos=PS=" & lis(0, idx) & "#PC=" & lis(3, idx) & "#PR=" & lis(1, idx) & "#cartella=" & cartella & "#file=" & lis(11, idx) & "#S=" & lis(7, idx) & "#C=" & lis(8, idx) & "#R=" & lis(10, idx) & "#HyElink"
        End If
        commento = commento & vbCrLf & "FcomTime#" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "#FcomTime"
        On Error Resume Next
        objSheet.Cells(lis(1, idx), lis(2, idx)).ClearComments
        If (Err.Number <> 0) Then
            Dim mess = "Errore clear comment scriviLinkTo:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore clear comment scriviLinkTo:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error Resume Next
        objSheet.Cells(lis(1, idx), lis(2, idx)).AddComment(commento)
        If (Err.Number <> 0) Then
            Dim mess = "Errore add comment scriviLinkTo:" & Err.Number & " Description " & Err.Description
            WriteMia(ConsoleColor.Red, mess, oOut, ofile)
            'Console.ForegroundColor = ConsoleColor.Red
            'Console.WriteLine("Errore add comment scriviLinkTo:" & Err.Number & " Description " & Err.Description)
            Err.Clear()
            Att = True
        End If
        On Error GoTo 0
        scriviLinkTo = True
    End Function

    Function rimuoviLink(ByRef cc)
        Dim PosI, PosF, tmp, tmp1, temporaneo
        PosI = InStr(1, cc, "HyFlink#", 1)
        PosF = InStr(1, cc, "#FcomTime", 1) + 8
        If ((PosF = Len(cc)) And PosI = 1) Then
            rimuoviLink = ""
            Exit Function
        Else
            tmp = Mid(cc, 1, PosI - 1)
            tmp1 = Mid(cc, PosF + 1, Len(cc) - PosF)
        End If
        temporaneo = rimuoviLineeVuoteIniziali(tmp & tmp1)
        rimuoviLink = temporaneo
    End Function

    Function rimuoviLineeVuoteIniziali(ByVal str)
        Do
            If (InStr(1, str, vbCrLf) = 1) Then
                str = Mid(str, 3, Len(str) - 2)
            End If
        Loop While (InStr(1, str, vbCrLf) = 1)
        rimuoviLineeVuoteIniziali = str
    End Function

End Module
