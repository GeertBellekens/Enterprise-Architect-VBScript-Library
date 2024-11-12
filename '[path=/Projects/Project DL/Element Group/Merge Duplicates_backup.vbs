'[path='[path=\Projects\Project DL\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Utils.Util

'
' Script Name: Merge Duplicates
' Author: Geert Bellekens
' Purpose: Merge a duplicate element with it's original element
' Date: 2019-04-25
'

'name of the output tab
const outPutName = "Merge Duplicates"

sub main 
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
 'check if the type and stereotype of the element are the same
 if originalElement.Stereotype <> selectedElement.Stereotype _
   or originalElement.Type <> selectedElement.Type then
  msgbox "Please select an element with the same name and type"
  exit sub
 end if
 'Ask user if he is sure
 dim response
 response = msgbox("Merge '" & selectedElement.Name & "' to '" & originalElement.Name & "'?" , vbYesNo+vbQuestion, "Merge Elements?")
 if response <> vbYes then
  'user did not confirm
  exit sub
 end if
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Starting merge of '" & selectedElement.Name & "'", selectedElement.ElementID
 'start with diagram usages
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Diagrams for '" & selectedElement.Name & "'", selectedElement.ElementID
 fixDiagrams selectedElement, originalElement
 'then process relations
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Relations for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeRelations selectedElement, originalElement
 'then process instances
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Instances for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeInstances selectedElement, originalElement
 'process nested elements
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Nested Elements for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeNestedElements selectedElement, originalElement
 'process nested diagrams
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Nested Diagrams for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeNestedDiagrams selectedElement, originalElement
 'process attributes using this entity as type
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Attributes for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeUsingAttributes selectedElement, originalElement
 'process parameters using this entity as type
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Parameters for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeUsingParameters selectedElement, originalElement
 'process operations using this enity as return type
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Operations for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeUsingOperations selectedElement, originalElement
 'process tagged value references
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Fixing Tagged value references for '" & selectedElement.Name & "'", selectedElement.ElementID
 mergeTaggedValueReferences selectedElement, originalElement
 'refresh
 'Repository.RefreshModelView(0)
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Finished merge of '" & selectedElement.Name & "'", selectedElement.ElementID
end sub

function mergeTaggedValueReferences (duplicateElement, originalElement)
 'fix tagged value references directly in the database
 dim sqlUpdate
 'element tags
 sqlUpdate = "update tv set tv.Value = '" & originalElement.ElementGUID & "' from t_objectproperties tv where tv.Value = '" & duplicateElement.ElementGUID & "'"
 Repository.Execute sqlUpdate
 'attribute tags
 sqlUpdate = "update tv set tv.Value = '" & originalElement.ElementGUID & "' from t_attributetag tv where tv.Value = '" & duplicateElement.ElementGUID & "'"
 Repository.Execute sqlUpdate
 'connector tags
 sqlUpdate = "update tv set tv.Value = '" & originalElement.ElementGUID & "' from t_connectortag tv where tv.Value = '" & duplicateElement.ElementGUID & "'"
 Repository.Execute sqlUpdate
 'operation tags
 sqlUpdate = "update tv set tv.Value = '" & originalElement.ElementGUID & "' from t_operationtag tv where tv.Value = '" & duplicateElement.ElementGUID & "'"
 Repository.Execute sqlUpdate
 'other tags
 sqlUpdate = "update tv set tv.TagValue = '" & originalElement.ElementGUID & "' from t_taggedvalue tv where tv.TagValue = '" & duplicateElement.ElementGUID & "'"
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
 'build construct picker string.
 dim constructpickerString
 constructpickerString = "IncludedTypes=" & duplicateElement.Type 
 if len(duplicateElement.Stereotype) > 0 then
  constructpickerString = constructpickerString & ";StereoType=" & duplicateElement.Stereotype
 end if
 constructpickerString = constructpickerString & ";Selection=" & duplicateElement.ElementGUID 
 'invoke the construct picker
 dim userSelectedElementID
 userSelectedElementID = Repository.InvokeConstructPicker(constructpickerString)
 if userSelectedElementID > 0 then
  set getUserSelectedElement = Repository.GetElementByID(userSelectedElementID)
 else
  set getUserSelectedElement = nothing
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

main
