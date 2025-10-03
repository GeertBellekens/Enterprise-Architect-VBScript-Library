'[path=\Projects\Project B\Conversion]
'[group=Conversion]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Transformation utils
' Author: Geert Bellekens
' Purpose: Reusable functions to aid in transformation scripts
' Date: 2025-10-02
'

function addAttributeDependencies(packageTreeIDString)
	'get all attributes that don't have a corresponding dependency
	dim sqlGetData
	sqlGetData = "select a.Object_ID, a.name, a.Classifier                 " & vbNewLine & _
				" , a.LowerBound, a.UpperBound                            " & vbNewLine & _
				" from t_attribute a                                      " & vbNewLine & _
				" inner join t_object o on o.Object_ID = a.Object_ID      " & vbNewLine & _
				" inner join t_object o2 on o2.Object_ID = a.Classifier   " & vbNewLine & _
				" where not exists                                        " & vbNewLine & _
				" 	(select c.Connector_ID                                " & vbNewLine & _
				" 	from t_connector c                                    " & vbNewLine & _
				" 	where c.Name = a.Name                                 " & vbNewLine & _
				" 	and c.Start_Object_ID = a.Object_ID                   " & vbNewLine & _
				" 	and c.Connector_Type = 'Dependency'                   " & vbNewLine & _
				" 	)                                                     " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString & ")       " & vbNewLine & _
				" order by a.Object_ID                                    "
	
	dim results
	set results = GetArrayListFromQuery(sqlGetdata)
	dim currentObjectID
	currentObjectID = 0
	dim row
	for each row in results
		dim objectID
		objectID = Clng(row(0))
		dim name
		name = row(1)
		dim classifierID
		classifierID = Clng(row(2))
		dim lowerBound
		lowerBound = row(3)
		dim upperBound
		upperBound = row(4)
		dim element as EA.Element
		if objectID <> currentObjectID then
			set element = Repository.GetElementByID(objectID)
			currentObjectID = objectID
			Repository.WriteOutput outPutName, now() & " Adding attribute dependencies for '" & element.Name &"'", 0
		end if
		dim dependency as EA.Connector
		set dependency = element.Connectors.AddNew(name, "Dependency")
		dependency.SupplierEnd.Cardinality = lowerBound & ".." & upperBound
		dependency.SupplierID = classifierID
		dependency.Update
	next
end function

function removeAttributeDependencies(packageTreeIDString)
	dim dependencyDictionary
	set dependencyDictionary = getDependencyDictionary(packageTreeIDString)
	dim currentObjectID
	currentObjectID = 0
	'Loop dictionary
	dim objectID
	for each objectID in dependencyDictionary.Items
		if objectID <> currentObjectID then 'only process each objectID once, when it is changed. This works because the dictionary is ordered by the object ids
			currentObjectID = objectID
			dim element as EA.Element
			set element = Repository.GetElementByID(objectID)
			deleteDependencies element, dependencyDictionary
		end if
	next
end function

function deleteDependencies(element, dependencyDictionary)
	'inform user
	Repository.WriteOutput outPutName, now() & " Removing attribute dependencies for '" & element.Name &"'", 0
	dim i
	for i = element.Connectors.Count - 1 to 0 step -1
		dim connector as EA.Connector
		set connector = element.Connectors.GetAt(i)
		if dependencyDictionary.Exists(connector.ConnectorID) then
			element.Connectors.DeleteAt i, false
		end if
	next
end function

function getDependencyDictionary(packageTreeIDString)
	dim dependencyDictionary
	set dependencyDictionary = CreateObject("Scripting.Dictionary")
	dim sqlGetData
	sqlGetData = "select distinct c.Connector_ID, o.Object_ID                " & vbNewLine & _
				" from t_connector c                                         " & vbNewLine & _
				" inner join t_object o on o.Object_ID = c.Start_Object_ID   " & vbNewLine & _
				" inner join t_attribute a on a.Object_ID = o.Object_ID      " & vbNewLine & _
				" 							and a.Name = c.Name              " & vbNewLine & _
				" where c.Connector_Type = 'Dependency'                      " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString & ")          " & vbNewLine & _
				" order by o.Object_ID, c.Connector_ID                       "
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	dim row
	for each row in results
		dim connectorID
		connectorID = Clng(row(0))
		dim objectID
		objectID = Clng(row(1))
		'add to dictionary
		dependencyDictionary.Add connectorID, objectID
	next
	'return
	set getDependencyDictionary = dependencyDictionary
end function