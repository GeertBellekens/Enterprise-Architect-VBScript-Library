'[path=\Framework\Tools\Script Management]
'[group=Script Management]

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
	dim allScripts, allGroups
	set allGroups = Nothing
	'get the folder from the user
	dim folder
    set folder = new FileSystemFolder
	set folder = folder.getUserSelectedFolder("")
	if not folder is nothing then
		set allScripts = script.getAllScripts(allGroups)
		Session.Output "allGroups.Count: " & allGroups.Count
	end if
	for each script in allScripts
		Session.Output "filename: " & folder.FullPath & script.Path & "\" & script.Name & ".vbs"
		dim file
		set file = New TextFile
		file.FullPath = folder.FullPath & script.Path & "\" & script.Name & ".vbs"
		'first make sure the code indicator is added to the code
		script.addGroupToCode
		'then save the script with the group indicator
		file.Contents = script.Code
		file.Save
	next
end sub

main