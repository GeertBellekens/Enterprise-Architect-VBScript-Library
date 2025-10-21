'[path=\General Scripts]
'[group=General Scripts]


!INC Wrappers.Include
!INC Local Scripts.EAConstants-VBScript

'
' Script Name: MergeDuplicatesMain
' Author: Geert Bellekens
' Purpose: Merge duplicated elements into one
' Date: 2022-05-13
'
'name of the output tab
const outPutName = "Merge Duplicates"

function mergeAllDuplicatesInPackage(package)
	'get a list of elements with their duplicate (original) element
	Repository.WriteOutput outPutName, now() & " Getting duplicate elements in package '" & package.Name & "'", 0
	dim duplicateDictionary
	set duplicateDictionary = getDuplicateDictionary(package)
	'don't do anything if no duplicates are found
	if not duplicateDictionary.Count > 0 then
		Repository.WriteOutput outPutName, now() & " No duplicates found in package '" & package.Name & "'", 0
		exit function
	end if
	'ask user if they are sure
	dim response
	response = msgbox("Merge " & duplicateDictionary.Count & " elements in package '" & package.Name & "'?" , vbYesNo+vbQuestion, "Merge Duplicates?")
	if not response = vbYes then
		Repository.WriteOutput outPutName, now() & " Merge cancelled by user", 0
		exit function
	end if
	'start merging
	dim element as EA.Element
	for each element in duplicateDictionary.Keys
		dim originalElement
		set originalElement = duplicateDictionary(element)
		'merge each one
		mergeElementWithOriginal element, originalElement
		'delete element
		Repository.WriteOutput outPutName, now() & " Deleting element '" & element.Name & "' with guid " & element.ElementGUID, 0
		deleteElement element
	next
end function

function getDuplicateDictionary(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select o.Object_ID as duplicateID, oo.Object_ID as originalID                                                                  " & vbNewLine & _
				" from t_object o                                                                                                               " & vbNewLine & _
				" inner join                                                                                                                    " & vbNewLine & _
				" 		(select *                                                                                                               " & vbNewLine & _
				" 		FROM (                                                                                                                  " & vbNewLine & _
				" 			SELECT                                                                                                              " & vbNewLine & _
				" 				oo.Object_ID,                                                                                                   " & vbNewLine & _
				" 				oo.Name,                                                                                                        " & vbNewLine & _
				" 				oo.Object_Type,                                                                                                 " & vbNewLine & _
				" 				oo.Stereotype,                                                                                                  " & vbNewLine & _
				" 				(SELECT COUNT(*)                                                                                                " & vbNewLine & _
				" 				 FROM t_diagramobjects dod                                                                                      " & vbNewLine & _
				" 				 WHERE dod.Object_ID = oo.Object_ID) AS diagramCount,                                                           " & vbNewLine & _
				" 				(SELECT COUNT(*)                                                                                                " & vbNewLine & _
				" 				 FROM t_connector c                                                                                             " & vbNewLine & _
				" 				 WHERE oo.Object_ID IN (c.Start_Object_ID, c.End_Object_ID)) AS relationCount,                                  " & vbNewLine & _
				" 				ROW_NUMBER() OVER (                                                                                             " & vbNewLine & _
				" 					PARTITION BY oo.Name, oo.Object_Type, oo.Stereotype                                                         " & vbNewLine & _
				" 					ORDER BY                                                                                                    " & vbNewLine & _
				" 						(SELECT COUNT(*) FROM t_diagramobjects dod WHERE dod.Object_ID = oo.Object_ID) DESC,                    " & vbNewLine & _
				" 						(SELECT COUNT(*) FROM t_connector c WHERE oo.Object_ID IN (c.Start_Object_ID, c.End_Object_ID)) DESC,   " & vbNewLine & _
				" 						oo.Object_ID ASC                                                                                        " & vbNewLine & _
				" 				) AS rn                                                                                                         " & vbNewLine & _
				" 			FROM t_object oo                                                                                                    " & vbNewLine & _
				" 			WHERE oo.Package_ID IN (" & packageTreeIDString  & ")                                                               " & vbNewLine & _
				" 		) ranked                                                                                                                " & vbNewLine & _
				" 		WHERE rn = 1                                                                                                            " & vbNewLine & _
				" 		) oo on oo.Name = o.Name                                                                                                " & vbNewLine & _
				" 			and oo.Object_ID <> o.Object_ID                                                                                     " & vbNewLine & _
				" 			and oo.Object_Type = o.Object_Type                                                                                  " & vbNewLine & _
				" 			and coalesce(oo.Stereotype, '') = coalesce(o.Stereotype, '')                                                        " & vbNewLine & _
				" where o.Package_ID IN (" & packageTreeIDString  & ")                                                                          "
	dim result
	set result = getArrayListFromQuery(sqlGetData)
	'create dictionary
	dim dictionary
	set dictionary = CreateObject("Scripting.Dictionary")
	dim row
	for each row in result
		dim elementID
		dim originalID
		elementID = Clng(row(0))
		originalID = Clng(row(1))
		dim element as EA.Element
		set element = Repository.GetElementByID(elementID)
		dim original as EA.Element
		set original = Repository.GetElementByID(originalID)
		if not dictionary.Exists(element) then
			dictionary.Add element, original
		end if
	next
	'return
	set getDuplicateDictionary = dictionary
end function


sub Merge 
	'exit if not on element
	if Repository.GetContextItemType() <> otElement then
		msgbox "Script only works on elements. Please select an element before executing this script"
		exit sub
	end if
	'get selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject()
	'Let the user select the original object
	msgbox "Please select the original object"
	dim originalElement as EA.Element
	set originalElement = getUserSelectedElement(selectedElement)
	'check if user selected 
	if originalElement is nothing then
	msgbox "User cancelled script"
	exit sub
	end if

	'check if user did not select the same element twice
	if originalElement.ElementID = selectedElement.ElementID then
		msgbox "Please select a different element"
		exit sub
	end if
	'check if the name and type of the element are the same
	if originalElement.Name <> selectedElement.Name _
	  or originalElement.Type <> selectedElement.Type then
		msgbox "Please select an element with the same name and type"
		exit sub
	end if
	'Ask user if he is sure
	dim response
	response = msgbox("Merge '" & selectedElement.FQName & "' to '" & originalElement.FQName & "'?" , vbYesNo+vbQuestion, "Merge Elements?")
	if response <> vbYes then
		'user did not confirm
		exit sub
	end if
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'actually merge elements
	mergeElementWithOriginal selectedElement, originalElement
	'refresh
	'Repository.RefreshModelView(0)
	response = msgbox("Delete '" & selectedElement.FQName & "'?" , vbYesNo+vbQuestion, "Delete Element?")
	if response = vbYes then
		'user confirmed
		deleteElement selectedElement
	end if
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished merge of '" & selectedElement.Name & "'", selectedElement.ElementID
end sub

function mergeElementWithOriginal(element, originalElement)
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting merge of '" & element.Name & "'", element.ElementID
	'start with diagram usages
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Diagrams for '" & element.Name & "'", element.ElementID
	fixDiagrams element, originalElement
	'then process relations
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Relations for '" & element.Name & "'", element.ElementID
	mergeRelations element, originalElement
	'then process instances
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Instances for '" & element.Name & "'", element.ElementID
	mergeInstances element, originalElement
	'process nested elements
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Nested Elements for '" & element.Name & "'", element.ElementID
	mergeNestedElements element, originalElement
	'process nested diagrams
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Nested Diagrams for '" & element.Name & "'", element.ElementID
	mergeNestedDiagrams element, originalElement
	'process attributes using this entity as type
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Attributes for '" & element.Name & "'", element.ElementID
	mergeUsingAttributes element, originalElement
	'process parameters using this entity as type
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Parameters for '" & element.Name & "'", element.ElementID
	mergeUsingParameters element, originalElement
	'process operations using this enity as return type
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Operations for '" & element.Name & "'", element.ElementID
	mergeUsingOperations element, originalElement
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Fixing Conveying Informationflows for '" & element.Name & "'", element.ElementID
	mergeConveyingInformationFlows element, originalElement
end function


function deleteElement(element)
	dim owner as EA.Element
	if element.ParentID > 0 then
		set owner = Repository.GetElementByID(element.ParentID)
	else
		set owner = Repository.GetPackageByID(element.PackageID)
	end if
	dim i
	for i = 0 to owner.Elements.Count -1
		dim temp
		set temp = owner.Elements(i)
		if temp.ElementID = element.ElementID then
			owner.Elements.DeleteAt i, false
			exit function
		end if
	next
end function

function mergeConveyingInformationFlows(duplicateElement, originalElement)
	'find operations that have the element as type of a parameter (or returntype)
	dim sqlUpdate
	sqlUpdate = "update x set x.Description = replace(x.Description, '" & duplicateElement.ElementGUID & "', '" & originalElement.ElementGUID & "')    " & vbNewLine & _
			" from t_xref x                                                                                                                          " & vbNewLine & _
			" where x.Behavior = 'conveyed'                                                                                                          " & vbNewLine & _
			" and x.Description like '%" & duplicateElement.ElementGUID & "%'                                                                        "
	
	'execute update statement
	Repository.Execute sqlUpdate
end function

function mergeUsingOperations(duplicateElement, originalElement)
	'find operations that have the element as type of a parameter (or returntype)
	dim sqlGetOperations
	sqlGetOperations = "select op.OperationID from t_operation op where op.Classifier = " & duplicateElement.ElementID
	dim operations
	set operations = getOperationsFromQuery(sqlGetOperations)
	'loop operations
	dim operation as EA.Method
	for each operation in operations
		operation.ClassifierID = originalElement.ElementID
		operation.ReturnType = originalElement.Name
		operation.Update
	next
end function

function mergeUsingParameters(duplicateElement, originalElement)
	'find operations that have the element as type of a parameter (or returntype)
	dim sqlGetOperations
	sqlGetOperations = "select distinct opr.OperationID from t_operationparams opr where opr.Classifier = " & duplicateElement.ElementID
	dim operations
	set operations = getOperationsFromQuery(sqlGetOperations)
	'loop operations
	dim operation as EA.Method
	for each operation in operations
		'loop parameters
		dim parameter as EA.Parameter
		for each parameter in operation.Parameters
			if parameter.ClassifierID = CStr(duplicateElement.ElementID) then 'need to convert to string because parameter.ClassifierID is a string and not a long
				parameter.ClassifierID = originalElement.ElementID
				parameter.Type = originalElement.Name
				parameter.Update
			end if
		next
	next
end function

function mergeUsingAttributes(duplicateElement, originalElement)
	'find using attributes
	dim sqlGetAttributes
	sqlGetAttributes = "select a.ID from t_attribute a where a.Classifier = " & duplicateElement.ElementID
	dim attributes 
	set attributes = getattributesFromQuery(sqlGetAttributes)
	'loop attributes
	dim attribute as EA.Attribute
	for each attribute in attributes
		attribute.ClassifierID = originalElement.ElementID
		attribute.Type = originalElement.Name
		attribute.Update
	next
end function

function mergeNestedDiagrams(duplicateElement, originalElement)
	dim nestedDiagram as EA.Diagram
	for each nestedDiagram in duplicateElement.Diagrams
		nestedDiagram.ParentID = originalElement.ElementID
		nestedDiagram.PackageID = originalElement.PackageID
		nestedDiagram.Update
	next
end function

function mergeNestedElements(duplicateElement, originalElement)
	dim nestedElement as EA.Element
	for each nestedElement in duplicateElement.Elements
		nestedElement.ParentID = originalElement.ElementID
		nestedElement.Update
	next
end function

function mergeInstances(duplicateElement, originalElement)
	'find all instances
	dim sqlFindInstances
	sqlFindInstances = "select  o.Object_ID from t_object o where o.Classifier = " & duplicateElement.ElementID
	dim instances
	set instances = getElementsFromQuery(sqlFindInstances)
	'loop instances
	dim instance as EA.Element
	for each instance in instances
		instance.ClassifierID = originalElement.ElementID
		instance.Update
	next
end function


'move all relations from and to the duplicate element to the original element
function mergeRelations(duplicateElement, originalElement)
	'move all relations from the dupliate element to the original element
	dim relation as EA.Connector
	for each relation in duplicateElement.Connectors
		'move the relation to the original element
		'set source
		if relation.ClientID = duplicateElement.ElementID then
			relation.ClientID = originalElement.ElementID
		end if
		'set target
		if relation.SupplierID = duplicateElement.ElementID then
			relation.SupplierID = originalElement.ElementID
		end if
		'check if such relation already exists
		dim mergedDuplicate as EA.Connector
		set mergedDuplicate = getMergedExistingRelation(relation, originalElement)
		if not mergedDuplicate is nothing then
			'save the possible changed to the merged duplicate
			mergedDuplicate.Update
		else
			'save the changes
			relation.Update
		end if
	next
end function

'check if a relation already exists
function getMergedExistingRelation(relation, originalElement)
	'initialize at nothing
	set getMergedExistingRelation = nothing
	dim orgRelation as EA.Connector
	for each orgRelation in originalElement.Connectors
		do 'do loop to be able to skip to next
			'check all parameters to skip to the next 
			if relation.Type <> orgRelation.Type _
				or relation.ClientID <> orgRelation.ClientID _
				or relation.SupplierID <> orgRelation.SupplierID _
				or relation.ConnectorID = orgRelation.ConnectorID then 'if it's the same then we skip as well
				exit do 'skip to next
			end if
			'set the name equal if empty
			if relation.Name <> orgRelation.Name then
				if len(orgRelation.Name) = 0 then
					orgRelation.Name = relation.Name
				end if
				if len(relation.Name) = 0 then
					relation.Name = orgRelation.Name
				end if
			end if
			'compare name
			if relation.Name <> orgRelation.Name then
				exit do 'skip to next
			end if
			'compare source end
			if not compareMergedConnectorEnd(relation.ClientEnd, orgRelation.ClientEnd) then
				exit do 'skip to next
			end if 
			'compare target end
			if not compareMergedConnectorEnd(relation.SupplierEnd, orgRelation.SupplierEnd) then
				exit do 'skip to next
			end if  
			'if we get here then we have a valid merged duplicate. Return connector and exit
			set getMergedExistingRelation = orgRelation
			exit function
		Loop While False
	next
end function

function compareMergedConnectorEnd (connectorEnd, orgConnectorEnd)
	'initialize false
	compareMergedConnectorEnd = false
	'merge cardinality
	if orgConnectorEnd.Cardinality <> connectorEnd.Cardinality then
		if len(orgConnectorEnd.Cardinality) = 0 then
			orgConnectorEnd.Cardinality = connectorEnd.Cardinality
		end if
		if len(connectorEnd.Cardinality) = 0 then
			connectorEnd.Cardinality = orgConnectorEnd.Cardinality
		end if
	end if
	'compare cardinality
	if getUnifiedMultiplicity(orgConnectorEnd.Cardinality) <> getUnifiedMultiplicity(connectorEnd.Cardinality) then
		exit function
	end if
	'compare aggregationKind
	if orgConnectorEnd.Aggregation <> connectorEnd.Aggregation then
		exit function
	end if
	'merge rolename
	if orgConnectorEnd.Role <> connectorEnd.Role then
		if len(orgConnectorEnd.Role) = 0 then
			orgConnectorEnd.Role = connectorEnd.Role
		end if
		if len(connectorEnd.Role) = 0 then
			connectorEnd.Role = orgConnectorEnd.Role
		end if
	end if
	'compare roleName
	if orgConnectorEnd.Role <> connectorEnd.Role then
		exit function
	end if
	'if we end up here they are the same
	compareMergedConnectorEnd = true
end function

function getUnifiedMultiplicity(multiplicity)
	getUnifiedMultiplicity = Replace(multiplicity, "0..*", "*")
	getUnifiedMultiplicity = Replace(getUnifiedMultiplicity, "1..1", "1")
end function


function getUserSelectedElement(duplicateElement)
	'get best match for original element
	dim candidateOriginal as EA.Element
	set  candidateOriginal =  getBestCandidateForOriginal(duplicateElement)
	'build construct picker string.
	dim constructpickerString
	constructpickerString = "IncludedTypes=" & duplicateElement.Type 
	if len(duplicateElement.Stereotype) > 0 then
		constructpickerString = constructpickerString & ";StereoType=" & duplicateElement.Stereotype
	end if
	if not candidateOriginal is nothing then
		constructpickerString = constructpickerString & ";Selection=" & candidateOriginal.ElementGUID 
	else
		constructpickerString = constructpickerString & ";Selection=" & duplicateElement.ElementGUID 
	end if
	'invoke the construct picker
	dim userSelectedElementID
	userSelectedElementID = Repository.InvokeConstructPicker(constructpickerString)
	if userSelectedElementID > 0 then
		set getUserSelectedElement = Repository.GetElementByID(userSelectedElementID)
	else
		set getUserSelectedElement = nothing
	end if
end function

function getBestCandidateForOriginal(duplicateElement)
	'initialize at null
	set getBestCandidateForOriginal = nothing
	dim sqlGetData
	sqlGetData = "select top(1) o.Object_ID from t_object o                             " & vbNewLine & _
					" left join                                                            " & vbNewLine & _
					"  (select do.Object_ID, count(*) as NumberOfUsesOnDiagram from       " & vbNewLine & _
					"   (select do.object_ID as Object_ID                              " & vbNewLine & _
					"   from t_diagramobjects do                                       " & vbNewLine & _
					"   union all                                                      " & vbNewLine & _
					"   select oo.Classifier                                           " & vbNewLine & _
					"   from t_diagramobjects do                                       " & vbNewLine & _
					"   inner join t_object oo on oo.Object_ID = do.Object_ID          " & vbNewLine & _
					"   ) do                                                           " & vbNewLine & _
					"   group by do.Object_ID                                          " & vbNewLine & _
					"  ) do2 on do2.Object_ID = o.Object_ID                               " & vbNewLine & _
					" where o.Name = '" & duplicateElement.Name & "'                       " & vbNewLine & _
					" and o.Object_ID <> " & duplicateElement.ElementID & "                " & vbNewLine & _
					" and o.Object_Type = '" & duplicateElement.Type & "'                  " & vbNewLine & _
					" and isnull(o.Stereotype, '') = '" & duplicateElement.Stereotype & "' " & vbNewLine & _
					" order by isnull(do2.NumberOfUsesOnDiagram, 0) desc                   "
	dim results
	set results = getElementsFromQuery(sqlGetData)
	if results.Count > 0 then
		set getBestCandidateForOriginal = results(0)
	end if
end function

'Quick and dirty via a database update
function fixDiagramsQuick(duplicateElement, originalElement)
	dim sqlUpdateDiagrams
	sqlUpdateDiagrams = "update t_diagramobjects set Object_ID = " & originalElement.ElementID & " where Object_ID = " & duplicateElement.ElementID
	Repository.Execute sqlUpdateDiagrams
end function

function fixDiagrams(duplicateElement, originalElement)
	'get diagrams where the duplicate element is shown
	dim sqlGetDiagrams
	sqlGetDiagrams = "select distinct do.Diagram_ID from t_diagramobjects do where do.Object_ID = " & duplicateElement.ElementID
	dim diagrams
	set diagrams = getDiagramsFromQuery(sqlGetDiagrams)
	'loop diagrams
	dim diagram as EA.Diagram
	for each diagram in diagrams
		'get diagramObject for the duplicate element
		dim diagramObject as EA.DiagramObject
		for each diagramObject in diagram.DiagramObjects
			if diagramObject.ElementID = duplicateElement.ElementID then
				diagramObject.ElementID = originalElement.ElementID
				diagramObject.Update
				'we could do an "exit for" here since in theory there should only be one diagramObject for a single element, but to be safe we don't.
			end if
		next
	next
end function

