'[path=\Projects\Project B\Project Browser Group]
'[group=Project Browser Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Baloise Scripts.LinkToCRMain

'This script only calls the function defined in the main script.
'Ths script is to be copied in Diagram, Search and Project Browser groups

'Execute main function defined in LinkToCRMain
sub main
	dim treeSelectedElements
	set treeSelectedElements = Repository.GetTreeSelectedElements()
	if treeSelectedElements.Count > 0 then
		linkItemToCR nothing, treeSelectedElements
	else
		dim selectedItem
		set selectedItem = Repository.GetTreeSelectedObject
		linkItemToCR selectedItem, nothing
	end if
end sub

main
