'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.DocGenUtil

' Script Name: Get Element impact
' Author: Matthias Van der Elst
' Purpose: Get impact for the selected element
' Date: 09/02/2017
'

'
' Project Browser Script main function
'
sub OnProjectBrowserElementScript()
	' Get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject()
	getImpact(selectedElement)
end sub

sub getImpact(selectedElement)
	dim eType 'Element type
	dim eStereoType 'Element stereotype
	dim eGUID 'Element GUID
	dim sqlGetImpact 'Query
	dim eID 'Element ID (= Object_ID)
	dim eName 'Element name
	eType = selectedElement.Type
	eStereoType = selectedElement.Stereotype
	eGUID = selectedElement.ElementGUID
	eID = selectedElement.ElementID
	eName = selectedElement.Name
	
	'Based on metamodel
	dim metamodel as EA.Diagram
	set metamodel = Repository.GetDiagramByGuid("{F6E96ABF-4D46-465c-851F-A83172B31786}")
	Session.Output metamodel.name
	Session.Output metamodel.DiagramID
	
	'Returns al the elements on the metamodel
	'select count(*) from t_diagramobjects where diagram_id = '8298'
	'select o.name, o.object_type from t_diagramobjects do
	'inner join t_object o on do.object_id = o.object_id
	'where do.diagram_id = '8298' and o.object_type <> 'Note'
	
	'Geselecteerde element van diagram situeren op het metamodel
	
	'De relaties ophalen die mogelijk zijn op basis van het metamodel
	
	
	dim lev_1 'Get the selected element based on the GUID
	set lev_1 = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact = 	"select lev_1.Object_ID " & _
					"from t_object lev_1 " & _
					"where  lev_1.ea_guid like '" & eGUID & "'"
	set lev_1 = getElementsFromQuery(sqlGetImpact)
	
	dim lev_2 'Get the elements that has a connection with the lev_1 element
	set lev_2 = CreateObject("System.Collections.ArrayList")
	sqlGetImpact = 	"select lev_2.Object_ID " & _
					"from t_object lev_2 " & _
					"inner join t_connector con_1 " & _
					"on lev_2.object_id = con_1.start_object_id or lev_2.object_id = con_1.end_object_id " & _
					"where (con_1.start_object_id = '" & eID & "' " & _
					"or con_1.end_object_id = '" & eID & "') " & _
					"and lev_2.Object_ID <> '" & eID & "' "
	
	set lev_2 = getElementsFromQuery(sqlGetImpact)		

	dim lev_3  'Get the elements that has a connection with the lev_2 elements
	set lev_3 = CreateObject("System.Collections.ArrayList")
	sqlGetImpact = 	"select lev_3.Object_ID " & _
					"from (((t_object lev_2 " & _
					"inner join t_connector con_1 " & _
					"on lev_2.object_id = con_1.start_object_id or lev_2.object_id = con_1.end_object_id) " & _
					"inner join t_connector con_2 " & _
					"on con_2.start_object_id = lev_2.object_id or con_2.end_object_id = lev_2.object_id) " & _
					"inner join t_object lev_3 " & _
					"on lev_3.object_id = con_2.start_object_id or lev_3.object_id = con_2.end_object_id) " & _
					"where (con_1.start_object_id = '" & eID & "' " & _
					"or con_1.end_object_id = '" & eID & "') " & _
					"and lev_2.Object_ID <> '" & eID & "' " & _
					"and lev_3.Object_ID <> lev_2.Object_ID "
					
					
	set lev_3 = getElementsFromQuery(sqlGetImpact)					
	
	dim lev_4 'Get the elements that has a connection with the lev_3 elements
	set lev_4 = CreateObject("System.Collections.ArrayList")
	sqlGetImpact = 	"select lev_4.Object_ID " & _
					"from (((((t_object lev_2 " & _
					"inner join t_connector con_1 " & _
					"on lev_2.object_id = con_1.start_object_id or lev_2.object_id = con_1.end_object_id) " & _
					"inner join t_connector con_2 " & _
					"on con_2.start_object_id = lev_2.object_id or con_2.end_object_id = lev_2.object_id) " & _
					"inner join t_object lev_3 " & _
					"on lev_3.object_id = con_2.start_object_id or lev_3.object_id = con_2.end_object_id) " & _
					"inner join t_connector con_3 " & _
					"on con_3.start_object_id = lev_3.object_id or con_3.end_object_id = lev_3.object_id) " & _
					"inner join t_object lev_4 " & _
					"on lev_4.object_id = con_3.start_object_id or lev_4.object_id = con_3.end_object_id) " & _
					"where (con_1.start_object_id = '" & eID & "' " & _
					"or con_1.end_object_id = '" & eID & "') " & _
					"and lev_2.Object_ID <> '" & eID & "' " & _
					"and lev_3.Object_ID <> lev_2.Object_ID " & _
					"and lev_4.Object_ID <> lev_3.Object_ID "
					
	set lev_4 = getElementsFromQuery(sqlGetImpact)					
	
	'exportResults eName, lev_1, lev_2, lev_3, lev_4  NOG UIT COMMENTAAR TE HALEN
	
	' dim result
	'result = MsgBox ("Export results to document?", vbYesNo, "Impact Analysis")

	'Select Case result
	'	Case vbYes
	'		exportResults()
	'	Case vbNo
	'		MsgBox("You chose No")
	'End Select
	
end sub

sub exportResults (eName, lev_1, lev_2, lev_3, lev_4)
	'Add master document
	dim packageGUID
	dim documentName
	dim masterDocument as EA.Package
	dim eTemplate

	
	packageGUID = "{21E715ED-25B2-4255-AF61-1EEAA8EE6305}"
	documentName = "Impact for " & eName
	eTemplate = "IMP_Element"
	set masterDocument = addMasterdocument(packageGUID, documentName) 'without further details
	
	dim element as EA.Element
	dim i
	i = 1
	'Model documents
	
	'Level 1
	addModelDocumentWithSearch masterDocument, "IMP_Level_1", "Level 1", "", i, ""
	Session.Output i & " - Level 1"
	i = i + 1
	for each element in lev_1
		addModelDocument masterDocument, eTemplate, element.Name, element.ElementGUID, i
		Session.Output i & " - " & element.Name 
		i = i + 1
	next
		
		Session.Output vbnewline 
		
	'Level 2
	addModelDocumentWithSearch masterDocument, "IMP_Level_2", "Level 2", "", i, ""
	Session.Output i & " - Level 2"
	i = i + 1	
	for each element in lev_2
		addModelDocument masterDocument, eTemplate, element.Name, element.ElementGUID, i
		Session.Output i & " - " & element.Name 
		i = i + 1
	next
	
		Session.Output vbnewline 
	
	'Level 3
	addModelDocumentWithSearch masterDocument, "IMP_Level_3", "Level 3", "", i, ""
	Session.Output i & " - Level 3"
	i = i + 1
	for each element in lev_3
		addModelDocument masterDocument, eTemplate, element.Name, element.ElementGUID, i
		Session.Output i & " - " & element.Name 
		i = i + 1
	next
	
		Session.Output vbnewline 
		
	'Level 4
	addModelDocumentWithSearch masterDocument, "IMP_Level_4", "Level 4", "", i, ""
	Session.Output i & " - Level 4" 
	i = i + 1
	for each element in lev_4
		addModelDocument masterDocument, eTemplate, element.Name, element.ElementGUID, i
		Session.Output i & " - " & element.Name 
		i = i + 1
	next
	

	'reload the package to show the correct order
	'Repository.RefreshModelView(masterDocument.PackageID)
	

end sub





	

OnProjectBrowserElementScript