'[path=\Projects\Project A\Rationalisation Data Models]
'[group=Rationalisation Data Models]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: Hide LDM Stereotypes items
' Author: Geert Bellekens
' Purpose: Hide all elements, attributes and associations that have the LDM stereotype
' Date: 2016-09-09
'
sub main
	'select source logical
	dim DMPackage as EA.Package
	msgbox "select the DM package"
	set DMPackage = selectPackage()
	dim DMPackageIDString, DMPackageTree
	set DMPackageTree = getPackageTree(DMPackage)
	dmPackageIDString = makePackageIDString(DMPackageTree)
	'hide attributes with LDM stereotype
	dim updateDiagramObjectsSQL
	 = "update do set do.ObjectStyle = do.ObjectStyle + 'HideStype=LDM;' " & _
		" from t_object o  " & _
		" join t_diagramobjects do on do.Object_ID = o.Object_ID " & _
		" join t_diagram d on d.Diagram_ID = do.Diagram_ID " & _
		" where d.Package_ID in (" & dmPackageIDString & ") " & _
		" and do.ObjectStyle not like '%HideStype=%' " 
	'hide associations with LDM stereotype
	dim 
end sub

main