'[path=\Framework\Tools\UML Profile]
'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Synchronize stereotypes
' Author: Geert Bellekens
' Purpose: Synchronizes all stereotypes in the list
' Date: 2016-01-12
'
dim outputTabName
outputTabName = "Synchronize stereotypes"

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'tell the user we are starting
	Repository.WriteOutput outputTabName, now() & ": Starting synchronizing stereotypes",0
	dim stereotypes, stereotype
	dim profile 
	set stereotypes = CreateObject("System.Collections.ArrayList")
	'*********** Set profile and stereotype names here Start *********************
	profile = "MyProfileName"
	stereotypes.Add "Stereotype1"
	stereotypes.Add "Stereotype2"
'	stereotypes.Add "Stereotype3"
	stereotypes.Add "Stereotype4"
	'*********** Add all stereotypes here End ***********************
	for each stereotype in stereotypes
		SynchronizeSteretoype profile, stereotype
	next
	'tell the user we are finished
	Repository.WriteOutput outputTabName, now() & ": Finished synchronizing stereotypes",0
end sub

function SynchronizeSteretoype(profile, stereotype)
	Repository.WriteOutput outputTabName, "Processing stereotype " & profile & "::" & stereotype,0
	Repository.CustomCommand "Repository", "SynchProfile", "Profile=" & profile & ";Stereotype=" & stereotype & ";" 
end function

main