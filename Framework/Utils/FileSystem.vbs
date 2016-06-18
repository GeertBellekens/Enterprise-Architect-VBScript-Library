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