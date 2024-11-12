'[path=\Projects\Project AP\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Remove Surplus Relations
' Author: Geert Bellekens
' Purpose: Remove the relations that are not required anymore
' Date: 2022-05-27
'
const outPutName = "Remove Surplus Relations"

function Main ()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	if selectedPackage is nothing then
		exit function
	end if
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting Remove Surplus Relations for package '" & selectedPackage.name & "'", 0
	'do the actual work
	removeSurplusRelations selectedPackage
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished Remove Surplus Relations for package '" & selectedPackage.name & "'", 0
		
end function

function removeSurplusRelations(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select c.Connector_ID from t_connector c                    " & vbNewLine & _
				" inner join t_object o on o.Object_ID = c.Start_Object_ID   " & vbNewLine & _
				" inner join t_package p on p.Package_ID = o.Package_ID      " & vbNewLine & _
				" inner join t_object oo on oo.Object_ID = c.End_Object_ID   " & vbNewLine & _
				" 						and o.Object_ID <> oo.Object_ID      " & vbNewLine & _
				" inner join t_package pp on pp.Package_ID = oo.Package_ID   " & vbNewLine & _
				" where c.Connector_Type = 'InformationFlow'                 " & vbNewLine & _
				" and not p.Name like 'SC%'                                  " & vbNewLine & _
				" and not pp.Name like 'SC%'                                 " & vbNewLine & _
				" and (                                                      " & vbNewLine & _
				" 	p.Package_ID in (" & packageTreeIDString & ")            " & vbNewLine & _
				" 	or pp.Package_ID in (" & packageTreeIDString & ")        " & vbNewLine & _
				" 	)                                                        " & vbNewLine & _
				" order by c.Start_Object_ID                                 "
	dim results
	set results = getConnectorsFromQuery(sqlGetData)
	Repository.WriteOutput outPutName, now() & " Found " & results.Count & " relations to remove", 0
	if results.Count = 0 then
		'no relations found
		exit function
	end if
	'make sure the user is sure:
	dim userIsSure
	userIsSure = Msgbox("Remove " & results.Count & " relations in package '" & package.Name & "'?", vbYesNo+vbQuestion, "Remove relations?")
	if userIsSure = vbYes then
		dim relation as EA.Connector
		dim i
		i = 0
		dim startObject as EA.Element
		set startObject = package.Element 'set to dummy element
		for each relation in results
			i = i + 1
			if relation.ClientID <> startObject.ElementID then
				set startObject = Repository.GetElementByID(relation.ClientID)
			end if
			'inform user
			Repository.WriteOutput outPutName, now() & " Removing relation " & i & " of " & results.Count & " from element '" & startObject.Name & "'", 0
			'actually delete relation
			deleteRelationFromStartObject startObject, relation
		next
	end if
end function

function deleteRelationFromStartObject(startObject, relation)
	dim i
	dim currentRelation as EA.Connector
	for i = startObject.Connectors.Count -1 to 0 step -1
		set currentRelation = startObject.Connectors.GetAt(i)
		if currentRelation.ConnectorID = relation.ConnectorID then
			startObject.Connectors.DeleteAt i, true 
			exit function
		end if 
	next
end function

main