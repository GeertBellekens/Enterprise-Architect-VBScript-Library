'[path=\Framework\Utils]
'[group=Utils]

'
' Script Name: FileSystem
' Author: Geert Bellekens
' Purpose: A collection of useful functions related to the file system
' Date: 2016-06-18
'
function getTempFilename()
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim tfolder, tname
	Const TemporaryFolder = 2
	Set tfolder = fso.GetSpecialFolder(TemporaryFolder)
	tname = fso.GetTempName    
	getTempFilename = tfolder &"\"& tname
End Function

function unzip (zipfile)
	'The folder the contents should be extracted to.
	dim extractTo, fso, filename, foldername
	Set fso = CreateObject("Scripting.FileSystemObject")
	filename = fso.GetFileName(zipfile)
	foldername = Replace(FileName, ".zip", "")
	extractTo = fso.GetParentFolderName(zipfile) & "\" & foldername
	
	'If the extraction location does not exist create it.
	If NOT fso.FolderExists(extractTo) Then
	   fso.CreateFolder(extractTo)
	End If

	'Extract the contents of the zip file.
	set objShell = CreateObject("Shell.Application")
	dim filesInZip
	set FilesInZip = objShell.NameSpace(zipfile).items
	objShell.NameSpace(extractTo).CopyHere(filesInZip)
	
	'clear objects
	Set fso = Nothing
	Set objShell = Nothing
	
	'return folder name
	unzip = extractTo
end function

Function ChooseFile (ByVal initialDir, filter)

	dim shel, fso, tempdir, tempfile, powershellfile, powershellOutputFile,psScript, textFile
	Set shell = CreateObject("WScript.Shell")

	Set fso = CreateObject("Scripting.FileSystemObject")

	tempDir = shell.ExpandEnvironmentStrings("%TEMP%")

	tempFile = tempDir & "\" & fso.GetTempName

	' temporary powershell script file to be invoked
	powershellFile = tempFile & ".ps1"

	' temporary file to store standard output from command
	powershellOutputFile = tempFile & ".txt"

	'if the filter is empty we use all files
	if len(filter) = 0 then
	filter = "All Files (*.*)|*.*"
	end if

	'input script
	psScript = psScript & "[System.Reflection.Assembly]::LoadWithPartialName(""System.windows.forms"") | Out-Null" & vbCRLF
	psScript = psScript & "$dlg = New-Object System.Windows.Forms.OpenFileDialog" & vbCRLF
	psScript = psScript & "$dlg.initialDirectory = """ &initialDir & """" & vbCRLF
	'psScript = psScript & "$dlg.filter = ""ZIP files|*.zip|Text Documents|*.txt|Shell Scripts|*.*sh|All Files|*.*""" & vbCRLF
	psScript = psScript & "$dlg.filter = """ & filter & """" & vbCRLF
	' filter index 4 would show all files by default
	' filter index 1 would should zip files by default
	psScript = psScript & "$dlg.FilterIndex = 1" & vbCRLF
	psScript = psScript & "$dlg.Title = ""Select a file""" & vbCRLF
	psScript = psScript & "$dlg.ShowHelp = $True" & vbCRLF
	psScript = psScript & "$dlg.ShowDialog() | Out-Null" & vbCRLF
	psScript = psScript & "Set-Content """ &powershellOutputFile & """ $dlg.FileName" & vbCRLF
	'MsgBox psScript

	Set textFile = fso.CreateTextFile(powershellFile, True)
	textFile.WriteLine(psScript)
	textFile.Close
	Set textFile = Nothing

	' objShell.Run (strCommand, [intWindowStyle], [bWaitOnReturn]) 
	' 0 Hide the window and activate another window.
	' bWaitOnReturn set to TRUE - indicating script should wait for the program 
	' to finish executing before continuing to the next statement

	Dim appCmd
	appCmd = "powershell -ExecutionPolicy unrestricted &'" & powershellFile & "'"
	'MsgBox appCmd
	shell.Run appCmd, 0, TRUE

	' open file for reading, do not create if missing, using system default format
	Set textFile = fso.OpenTextFile(powershellOutputFile, 1, 0, -2)
	ChooseFile = textFile.ReadLine
	textFile.Close
	Set textFile = Nothing
	fso.DeleteFile(powershellFile)
	fso.DeleteFile(powershellOutputFile)

End Function

Function cleanFileName(fileName)
	Dim regEx
	Set regEx = CreateObject("VBScript.RegExp")

	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.Pattern = "[(?*"",\\<>&#~%{}+@:\/!;]+"
	cleanFileName = regEx.Replace(fileName, "-")
end function