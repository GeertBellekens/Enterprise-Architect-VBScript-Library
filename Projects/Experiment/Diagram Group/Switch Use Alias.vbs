'[path=\Projects\Experiment\Diagram Group]
'[group=Diagram Group]
'[group_type=DIAGRAM]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Switch Use Alias
' Author: Geert Bellekens
' Purpose: Switches the option Use Alias on the currently selected diagram
' Date: 2017-03-16
'
sub main
	dim diagram as EA.Diagram
	set diagram = Repository.GetCurrentDiagram()
	if not diagram is nothing then
		if instr(diagram.ExtendedStyle, "UseAlias=0") > 0 then
			diagram.ExtendedStyle = replace(diagram.ExtendedStyle,"UseAlias=0","UseAlias=1")
		elseif instr(diagram.ExtendedStyle, "UseAlias=1") > 0 then
			diagram.ExtendedStyle = replace(diagram.ExtendedStyle,"UseAlias=1","UseAlias=0")
		else
			diagram.ExtendedStyle = diagram.ExtendedStyle & "UseAlias=1;"
		end if
		diagram.Update
		Repository.ReloadDiagram(diagram.DiagramID)
	end if
end sub

main