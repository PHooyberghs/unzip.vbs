option explicit

'''Credits''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' I've learned a lot from scripts by Rob van der Woude
' Rob van der Woude's Scripting Pages
' In particular, these pages
' http://www.robvanderwoude.com/vbstech_databases_access.php
' http://www.robvanderwoude.com/vbstech_files_zip.php
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'variable declaration
dim iSplit, sSource, sSourcePath, sSourcefile, sSourceExtension, sDestination, sLogPath, sTmpPath, lMax_Size, sExtensionList, blnSelectIn, sDestinationFin
dim objShell, objFileSystem, objRegExp
dim sPathToCheck, sFolderShort1, sFolderShort2
dim sLogFile, foLogFile
dim sSourcefile_FIN
dim objFolderStart

'check command line arguments
Get_Arguments

'set/intialize base objects
Base_Objects_Initialize

' check/create destination folder and subfolders
checkpath sDestination
checkpath sDestinationFin
' checkpath sLogPath 'no longer required since sLogPath is no longer used
checkpath sTmpPath

'initialize/start log
Log_Initialize

'check sourcefile
SourceFile_Check

'modify sourcefile extension if not zip
Extension_ZIP

'loop trough namespace objectitems
set objFolderStart=objShell.Namespace(sSourcefile_FIN)
Loop_Source_Folder objFolderStart, sDestinationFin, "[" & sSourcefile & "]"

' remove temporary file when if applicable
if sSourcefile_FIN<>sSource then objFileSystem.deletefile(sSourcefile_FIN)

'extraction attempt completed
foLogFile.writeline("Extraction attempt completed")

'normal ending
Normal_Exit

'subs
'normal exit
function Normal_Exit
	foLogFile.writeline("Logging stopped")
	set foLogFile=nothing
	set objRegExp = nothing	
	Set objFileSystem = nothing
	Set objShell = nothing
	wscript.quit(0)
End function

' check/create folder
Sub CheckPath(sPathToCheck)
	objRegExp.pattern="(.+)(\\[\w\d\s]+)$"
	do while not (objFileSystem.folderexists(sPathToCheck))
		sFolderShort1=sPathToCheck
		do while not (objFileSystem.folderexists(sFolderShort1))
			sFolderShort2=objRegExp.Replace(sFolderShort1,"$1")
			if (objFileSystem.folderexists(sFolderShort2)) then
				objFileSystem.createfolder(sFolderShort1)
				wscript.echo "folder '" & sFolderShort1 & "' created"
				sFolderShort1=sPathToCheck
				exit do
			else
				sFolderShort1=sFolderShort2
			end if
		loop
	loop
end sub

'check command line arguments
sub Get_Arguments
	with wscript.arguments
		'wscript.echo .unnamed.count
		if .unnamed.count<>2 then
			Syntax_Error
		else
			sSource = ucase(trim(.unnamed(0)))
			sDestination = ucase(trim(.unnamed(1)))
		end if
		if .named.count>0 then
			if .Named.Exists("MSB") then
				lMax_Size=clng(.named.item("MSB"))
			else
				lMax_Size=clng(0)
			end if
			if .Named.Exists("FEF") then
				sExtensionList=.named.item("FEF")
			else
				sExtensionList=""
				blnSelectIn=0
			end if
			if .Named.Exists("FO") then
				blnSelectIn=0
			else
				blnSelectIn=1
			end if
		end if
	end with
	iSplit=InStrRev(sSource,"\")
	sSourcePath=left(sSource,iSplit)
	sSourcefile=right(sSource,len(sSource)-iSplit)
	isplit=InStrRev(sSourcefile,".")
	sSourceExtension=right(sSourcefile,len(sSourcefile)-iSplit)
	sSourcefile=left(sSourcefile,iSplit-1)
	if right(sDestination,1)="\" then sDestination=left(sDestination,len(sDestination)-1)
	sDestinationFin=sDestination & "\" & sSourcefile
	'sLogPath=sDestination & "\log"
	sTmpPath=sDestination & "\tmp"
end sub

'set/intialize base objects
sub Base_Objects_Initialize
	Set objShell = CreateObject("Shell.Application")
	Set objFileSystem = CreateObject( "Scripting.FileSystemObject" )
	set objRegExp = new regexp
	objRegExp.IgnoreCase=-1
	objRegExp.Global=-1
end sub

'initialize/start log
sub Log_Initialize
	'sLogFile=SLogPath &  "\" & sSourcefile &"_vbs_unzip.log"
	sLogFile=sDestinationFin &  "\" & sSourcefile &"_vbs_unzip.log"
	set foLogFile=objFileSystem.CreateTextFile(sLogFile,true)
	foLogFile.writeline("Logging started")
end sub

'check sourcefile
sub SourceFile_Check
	if objFileSystem.Fileexists(sSource)=false then
		foLogFile.writeline("Sourcefile not found:")
		foLogFile.writeline(sSource)
		foLogFile.writeline("")
		wscript.echo "ERROR: Source file does not exist"
		wscript.echo
		WScript.Quit(2)
	else
		foLogFile.writeline("Extraction attempt from:")
		foLogFile.writeline(sSource)
	end if
end sub

'modify sourcefile extension if not zip
sub Extension_ZIP
	wscript.echo "sSource: " & sSource
	if ucase(sSourceExtension)<>"ZIP" then
		sSourcefile_FIN=sTmpPath & "\" & sSourcefile & ".ZIP"
		objFileSystem.copyfile sSource, sSourcefile_FIN
	else
		sSourcefile_FIN=sSource
	end if
	wscript.echo "sSourcefile_FIN: " & sSourcefile_FIN
	wscript.echo
end sub

'loop trough namespace objectitems
		'objectitem=folder ==> 'loop trough namespace objectitems
		'objectitem=file ==> 
				'create new subfolder within destinationfolder with same name as parentfolder (when not existing)
				'copy from source to this subfolder
Sub Loop_Source_Folder(objFolder,sDestination_local,sBasePath)
	dim nsFolder, objFolderItems, intFolderItems_Cnt, sMsg, reTest, lFileSize
	dim cnt, objItem, sFileExtension, sDestinationTmp, objDestination, sTargetFile,blnCopy_OK, nAttempt,nAttempt_sub
	dim iCopyParameters, blnExtract, sNoUnzipLog, foTarget
	iCopyParameters=4+8+512+1024
	set nsFolder=objShell.Namespace(objFolder)
	set objFolderItems=nsFolder.items
	checkpath sDestination_local
	set objDestination=objShell.Namespace(sDestination_local)
	intFolderItems_Cnt=objFolderItems.count
	objRegExp.pattern=sExtensionList
	'folders
	for cnt=0 to intFolderItems_Cnt-1
		set objItem=objFolderItems.item(cnt)
		If objItem.IsFolder Then
			sDestinationTmp=sDestination_local & "\" & objItem.name
			Loop_Source_Folder objItem.getfolder,sDestinationTmp, sBasePath & "\" & objItem.name
		else
		end if
	next
	'files
	for cnt=0 to intFolderItems_Cnt-1
		set objItem=objFolderItems.item(cnt)
		If objItem.IsFolder Then
		else
			blnExtract=1
			lFileSize=clng(objItem.size)
			sTargetFile=sDestination_local & "\" & objItem.Name
			sFileExtension=objFileSystem.GetExtensionName(sTargetFile)
			sMsg=""
			sMsg=sMsg & vbCrLf & "Item: " & sBasePath & "\" & objItem.name
			sMsg=sMsg & vbCrLf & "Size: " & cstr(lFileSize)
			if (lMax_Size>0 and abs(lFileSize) > abs(lMax_Size))=true then
				sMsg=sMsg & "Not unzipped due to maximum size filter (" & lMax_Size & ")"
				blnExtract=0
			end if
			if ((blnSelectIn=0 and objRegExp.test(trim(sFileExtension))=-1) or (blnSelectIn=1 and objRegExp.test(trim(sFileExtension))=0)) then
				sMsg=sMsg & vbCrLf & "Not unzipped due to extension filter (" & objRegExp.pattern & ")"
				blnExtract=0
			end if
			if blnExtract>0 then
				blnExtract=0
				nAttempt=0
				blnCopy_OK=0
				do while blnCopy_OK=0 and nAttempt<10
					nAttempt=nAttempt+1
					blnCopy_OK=1
					objDestination.copyhere objItem,&HC1C
					nAttempt_sub=0
					do while objFileSystem.fileexists(sTargetFile)=false and nAttempt_sub<10
						blnCopy_OK=0
						nAttempt_sub=nAttempt_sub+1
						wscript.sleep(3000)
					loop
					if blnCopy_OK>0 then
						blnCopy_OK=1
						nAttempt_sub=0
						set foTarget=objFileSystem.GetFile(sTargetFile)
						do while foTarget.Size<lFileSize and nAttempt_sub<10
							blnCopy_OK=0
							nAttempt_sub=nAttempt_sub+1
							wscript.sleep(3000)
						loop
					end if
				loop
				if blnCopy_OK>0 then
					sMsg=sMsg & vbCrLf & "Extracted to:"
					sMsg=sMsg & vbCrLf & sTargetFile
				else
					sMsg=sMsg & vbCrLf & "Extrax-ction failed !!!!"
				end if
			end if
			wscript.echo sMsg
			wscript.echo
			foLogFile.writeline(sMsg)
			foLogFile.writeline("")
		end if
	next
End sub

Sub Syntax_Error
    Dim sMsg_SE
	sMsg_SE=""
    sMsg_SE = sMsg_SE & vbCrLf _
			& "windowsunzip.vbs" & vbCrLf _
			& "Extract folders and files from compressed folder"  & vbCrLf _
			& "where extension is found in given extension list" & vbCrLf _
			& vbCrLf & vbCrLf _
			& "Usage:" & vbCrLf _
			& "CSCRIPT  //NOLOGO  " & vbCrLf _
			& "		[path_to_script\]windowsunzip.VBS" & vbCrLf _
			& "		Source_Folder" & vbCrLf _
			& "		Target_Folder" & vbCrLf _
			& "		MAX_SIZE_BYTES" & vbCrLf _
			& "		Extension_List" & vbCrLf _
			& "		[/FO]" & vbCrLf _
			& vbCrLf & vbCrLf _
			& "Where:" & vbCrLf _
			& """Source_Folder"" is a compressed filesystem folder"   & vbCrLf _
			& """Target_Folder"" is the filesystem folder in which the folder structure and file names must be extracted"   & vbCrLf _
			& """MAX_SIZE_BYTES"" is maximum size in bytes of files that will be extracted"   & vbCrLf _
			& """Extension_List"" is the list of extensions for which file extraction has to be performed"  & vbCrLf _
			& "  /FO			select out extension list (default is selecting in)" & vbCrLf _
			& vbCrLf & vbCrLf _
			& "Written by HSP" &vbCrLf _
			& "Sligthly based on:" &vbCrLf _
			& "Rob van der Woude" & vbCrLf _
			& "http://www.robvanderwoude.com" & vbCrLf _
			& "http://www.robvanderwoude.com/vbstech_databases_access.php"
    WScript.Echo sMsg_SE
    WScript.Quit(1)
End Sub



