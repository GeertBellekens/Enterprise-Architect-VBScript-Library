option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Author: Geert Bellekens
' Purpose: Saves all scripts in a given folder on the file system
' Date: 2015-12-07
'
sub main
	dim script
	set script = New Script
	dim allScripts
	set allScripts = script.getAllScripts()
	'get the folder from the user
	dim folder, shell
	Set shell  = CreateObject( "Shell.Application" )
    Set folder = shell.BrowseForFolder( 0, "Select Folder", 0, "C:\Users\wij\Documents\BellekensIT\Development\Enterprise-Architect-VBScript-Library" )
	if not folder is nothing then
		Session.Output "folder.Self.Path: " & folder.Self.Path
		set allScripts = script.getAllScripts()
	end if
	for each script in allScripts
		Session.Output "filename: " & folder.Self.Path & "\" & script.Group.Name & "\" & script.Name & ".vbs"
		dim file
		set file = New TextFile
		file.Folder = folder.Self.Path & "\" & script.Group.Name
		file.FileName = script.Name & ".vbs"
		file.Contents = script.Code
		file.Save
	next
end sub

main