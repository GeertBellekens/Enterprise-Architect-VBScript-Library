'[path=\Projects\Project B\Model Management]
'[group=Model Management]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Remove All From Version Control
' Author: Geert Bellekens
' Purpose: Removes ALL packages from version control. Use only on backup models.
' Date: 2019-04-03
'
sub main
	'do not execute on SQL server databases
	if Repository.RepositoryType = "SQLSVR" then
		Msgbox "Script is disabled on SQL server repositories", vbExclamation, "Not allowed!"
		exit sub
	end if
	'ask for confirmation
	dim userIsSure
	userIsSure = Msgbox("Are you sure you want to remove ALL packages from version control?" & nvbNewLine & "This should NOT be used on production models!", vbYesNo+vbExclamation, "Remove ALL from TFS?")
	if userIsSure = vbYes then
		Repository.Execute "update t_package set IsControlled = 0"
		MsgBox "Finished!"
	end if
end sub

main