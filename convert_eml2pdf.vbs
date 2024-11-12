'script		convert_eml2pdf
'author 	daniel brenk
'date		04.07.2018
'changes	change history
'1			V4	2022		use 8.1-name/shortpath to prevent unsupported characters in filenames to break the process
'2			V5	2024-11-12	cleanup
'remark		this script uses 2 open-source tools that both are under Apache V2 License
'1			https://www.whitebyte.info/publications/eml-to-pdf-converter, https://github.com/nickrussler/eml-to-pdf-converter
'2			https://pdfbox.apache.org/

Option Explicit

Const mailconverterdir = "C:\Program Files (x86)\EMLtoPDFConverter"
Const mailconverterjar = "emailconverter.jar"
Const inputdir = "C:\SERDATA\Programmierung_Skripte\eml2pdf\input"
Const outputdir = "C:\SERDATA\Programmierung_Skripte\eml2pdf\output"
Const backupdir = "C:\SERDATA\Programmierung_Skripte\eml2pdf\backup"
Const pdfboxdir = "C:\SERDATA\Programmierung_Skripte\eml2pdf"
Const pdfboxjar = "pdfbox-app-2.0.9.jar"

Const logfile = "C:\SERDATA\Programmierung_Skripte\eml2pdf\convert_eml2pdf_V5.log"


'loop through eml files in a folder
'rename .eml to _mailbody.eml
'convert _mailbody.eml to _mailbody.pdf
'search the attachment .pdf files that belong to the _mailbody.pdf
'collate those to one _completemail.pdf
'put the resulting file in the output directory
'move input files to backup directory


work(inputdir)

Private Function work(sFolder)
	Dim oFile, oFSO, sEmlname
	Dim sFilename, sTypeEnding
	Dim iLen
	Dim oList
	Dim sOutFilePath, sNewFilePath, mailbodypdfpath, sOrigEmlPath, sOrigEmlPathTypeless, sAttachmentDir
	Dim entry
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oList = CreateObject("System.Collections.ArrayList")

	WriteLogFileLine logfile, "DEBUG - - - work(" & sFolder & ") started - - -"

	if doesFolderExist(mailconverterdir) and doesFolderExist(inputdir) and doesFolderExist(outputdir) and doesFolderExist(backupdir) and doesFolderExist(pdfboxdir) Then
		WriteLogFileLine logfile, "DEBUG folders ok"
	Else
		WriteLogFileLine logfile, "ERROR folders missing"
		exit function
	end if
	
	For Each oFile In oFSO.GetFolder(sFolder).Files
	  If UCase(oFSO.GetExtensionName(oFile.Name)) = "EML" Then
		sEmlname = oFile.Name
		sOrigEmlPath = oFile.Path
		sOrigEmlPathTypeless = Left(sOrigEmlPath, InstrRev(sOrigEmlPath, ".") -1)
		iLen = InstrRev(sEmlname, ".")
		sFilename = Left(sEmlname, iLen - 1)
		sTypeEnding = Right(sEmlname, Len(sEmlname) - iLen)
		sNewFilePath = inputdir & "\" & sFilename & "_mailbody" & ".eml"
		oFSO.MoveFile oFile.Path, sNewFilePath
		Set oFile = oFSO.GetFile(sNewFilePath)
		renderEml2Pdf(oFile.Path)
		subSleep 3
		mailbodypdfpath = Left(oFile.Path, InstrRev(oFile.Path, ".")-1) & ".pdf"
		'msgbox "mailodypdfpath: " & mailbodypdfpath
		oList.Add mailbodypdfpath
		Dim oFsoSearch, oFile2
		Set oFsoSearch = CreateObject("Scripting.FileSystemObject")
			sAttachmentDir = inputdir & "\" & sFilename & "_mailbody" & "-attachments"
			'msgbox sAttachmentDir
			if oFsoSearch.FolderExists(sAttachmentDir) then
				For Each oFile2 in oFsoSearch.GetFolder(sAttachmentDir).Files
					'msgbox "checking: " & Left(oFile2.Name, InstrRev(oFile2.Name, ".")-1) & "=" & sFilename
					If UCase(oFsoSearch.GetExtensionName(oFile2)) = "PDF" Then
						if oFile2.Path <> mailbodypdfpath then
							oList.Add oFile2.ShortPath 'SER dab 2022 ShortPath Gibt den kurzen Pfad einer angegebenen Datei zurück (die 8.3-Benennungskonvention)
							'oList.Add oFile2.Path
						End If
					End If
				Next
			end if
		sOutFilePath = outputdir & "\" & sFilename & "_completemail_" & timestamp() & ".pdf"
		'msgbox "collatingTo: " & sOutFilePath & CHR(13) & "Files:" & oList.Count
		
		'if more than one PDF is to be processed into one "complete" PDF then Collate else Move
		if oList.Count > 1 then
			collatePDFFiles oList, sOutFilePath
		else
			For Each entry in oList
				oFsoSearch.MoveFile entry, sOutFilePath
			Next
		end if
		
		oList.Clear()
		
		If oFSO.FileExists(sOutFilePath) Then
				'move to backup
				if oFsoSearch.FileExists(sNewFilePath) then
					oFsoSearch.MoveFile sNewFilePath, backupdir & "\" & oFile.Name & timeStamp()
				end if
				
				if oFsoSearch.FolderExists(inputdir) then
					For Each oFile2 in oFsoSearch.GetFolder(inputdir).Files
						If Left(oFile2.Name, Len(sFilename)) = sFilename Then
							oFsoSearch.MoveFile oFile2.Path, backupdir & "\" & oFile2.Name & timeStamp()
						End If
					Next
				end if
				
				if oFsoSearch.FolderExists(sAttachmentDir) then
					For Each oFile2 in oFsoSearch.GetFolder(sAttachmentDir).Files
						oFsoSearch.MoveFile oFile2.Path, backupdir & "\" & oFile2.Name & timeStamp()
					Next
				End If
				
				if oFsoSearch.FolderExists(sAttachmentDir) then
					oFsoSearch.DeleteFolder sAttachmentDir
				End if
				
				Set oFsoSearch = nothing
			  End if
		Else
			WriteLogFileLine logfile, "ERROR file does not exist outfile[" & sOutFilePath & "] originalEML[" & sOrigEmlPath & "]"
			
		End If
		
	
	Next

	Set oFSO = Nothing
	
End Function


Private Function renderEml2Pdf(emlfile)
	WriteLogFileLine logfile, "DEBUG renderEml2Pdf[" & emlfile & "]"
	Dim oShell
	Dim sCmd
	Dim oReturn
		Set oShell = CreateObject("WScript.Shell")
		oShell.CurrentDirectory = mailconverterdir
		sCmd = "java -jar" & " " & """" & mailconverterjar & """"&  " --debug --extract-attachments " & """" & emlfile & """"
		WriteLogFileLine logfile, "DEBUG renderEml2Pdf[" & sCmd & "]"
		oReturn = oShell.Run(sCmd, 6, True)
		
		WriteLogFileLine logfile, "DEBUG renderEml2Pdf.Return[" & oReturn & "]"
	set oShell = Nothing
	
	'doesFileExist sOutFilePath
End Function


Private Function collatePDFFiles(oList, sOutFilePath)
	Dim oShell
	Dim sCmd
	Dim oReturn
	Dim entry
	
	For Each entry in oList
		doesFileExist entry
	Next
		Set oShell = CreateObject("WScript.Shell")
		oShell.CurrentDirectory = pdfboxdir
		'java -jar pdfbox-app-2.0.9.jar PDFMerger File1.pdf File2.pdf outFile.pdf
		'sCmd = "java -jar" & " " & """" & mailconverterjar & """"&  " -d " & """" & emlfile & """"
		sCmd = "java -jar" & " " & """" & pdfboxjar & """" & " PDFMerger"
		For Each entry in oList
			sCmd = sCmd & " " & """" & entry & """"
		Next
		sCmd = sCmd & " " & """" & sOutFilePath & """"
		'msgbox sCmd
		WriteLogFileLine logfile, "DEBUG collatePDFFiles[" & sCmd & "]"

		oReturn = oShell.Run(sCmd, 6, True)
	set oShell = Nothing
	doesFileExist sOutFilePath
End Function

Function timeStamp()
    Dim t 
    t = Now
    timeStamp = Year(t) & "-" & _
    Right("0" & Month(t),2)  & "-" & _
    Right("0" & Day(t),2)  & "_" & _  
    Right("0" & Hour(t),2) & _
    Right("0" & Minute(t),2) & _    
	Right("0" & Second(t),2) 
End Function

Function WriteLogFileLine(sLogFileName,sLogFileLine)
	Dim objFsoLog, logOutput
    Set objFsoLog = CreateObject("Scripting.FileSystemObject")
	Set logOutput = objFsoLog.OpenTextFile(sLogFileName, 8, True)
    logOutput.WriteLine(cstr(timeStamp) + " -" + vbTab + sLogFileLine)
	logOutput.Close
    Set logOutput = Nothing
	Set objFsoLog = Nothing
End Function

Function doesFileExist(sFilePath)
	Dim obFSO1
    Set obFSO1 = CreateObject("Scripting.FileSystemObject")
	If obFSO1.FileExists(sFilePath) Then
		WriteLogFileLine logfile, "DEBUG file does exist file[" & sFilePath & "]"
		doesFileExist = true
	Else
		WriteLogFileLine logfile, "ERROR file does not exist file[" & sFilePath & "]"
		doesFileExist = false
	End If
	Set obFSO1 = Nothing
End Function

Function doesFolderExist(sDirPath)
	Dim obFSO2
    Set obFSO2 = CreateObject("Scripting.FileSystemObject")
	If obFSO2.FolderExists(sDirPath) Then
		doesFolderExist = true
	Else
		WriteLogFileLine logfile, "ERROR folder does not exist folder[" & sDirPath & "]"
		doesFolderExist = false
	End If
	Set obFSO2 = Nothing
End Function

Sub subSleep(strSeconds) ' subSleep(2)
    Dim objShell
    Dim strCmd
    set objShell = CreateObject("wscript.Shell")
    strCmd = "%COMSPEC% /c ping -n " & strSeconds & " 127.0.0.1>nul"     
    objShell.Run strCmd,0,1 
End Sub 