'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: 
' Author: 
' Purpose: Remove the duplicate generalizations
' Date: 
'
sub main
	dim subElements
	dim subElement as EA.Element
	'get the subElements that have more then one 
	dim getSubElementsQuery
	getSubElementsQuery = "select sto.Object_ID, eno.Object_ID as endObject,count(*)     " & _
						" from ((t_connector c                                         " & _
						" inner join t_object sto on sto.Object_ID = c.Start_Object_ID)" & _
						" inner join t_object eno on eno.Object_ID = c.End_Object_ID)  " & _
						" where c.Connector_Type = 'Generalization'                    " & _
						" group by sto.Object_ID, eno.Object_ID                        " & _
						" having count(*) > 1                                          "
	set subElements = getElementsFromQuery(getSubElementsQuery)
	'Loop the connectors and delet the duplicate generalizations
	for each subElement in subElements
		set superClassIDs = CreateObject("System.Collections.ArrayList")
		dim generalization as EA.Connector
		dim i
		for i = subElement.Connectors.Count -1 to 0 step -1
			set generalization = subElement.Connectors.GetAt(i)
			if generalization.Type = "Generalization" then
				if superClassIDs.Contains(generalization.SupplierID) then
					subElement.Connectors.DeleteAt i,false 
					Session.Output "Deleting generalization between " & subElement.Name & " and superclassID " & generalization.SupplierID
				else
					superClassIDs.Add generalization.SupplierID
					Session.Output "Adding generalization to list between " & subElement.Name & " and superclassID " & generalization.SupplierID
				end if
			end if
		next
	next
end sub

main