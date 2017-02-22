'[path=\Projects\Project K\DiagramGroup]
'[group=DiagramGroup]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Set composition source and target
' Author: Geert Bellekens
' Purpose: Make sure that the whole end is always the source, and the part end is always the target
' Date: 18/11/2015
'

'
' Diagram Script main function
'
sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()

	if not currentDiagram is nothing then
		'first save the diagram
		Repository.SaveDiagram currentDiagram.DiagramID
		' Get a reference to any selected connector/objects
		dim selectedConnector as EA.Connector
		set selectedConnector = currentDiagram.SelectedConnector

		if not selectedConnector is nothing then
			'correct the composition direction for a single composition
			correctCompositionDirection selectedConnector
			'reload diagram to show changes
			Repository.ReloadDiagram currentDiagram.DiagramID
		else
			'correct the composition direction for all compositions in the diagram
			dim diagramLink as EA.DiagramLink
			for each diagramLink in currentDiagram.DiagramLinks
				'get connector from diagram link
				dim connector as EA.Connector
				set connector = Repository.GetConnectorByID(diagramLink.ConnectorID)
				'set composition source and target
				correctCompositionDirection connector
			next
			'reload diagram to show changes
			Repository.ReloadDiagram currentDiagram.DiagramID
		end if
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if

end sub

OnDiagramScript

function correctCompositionDirection(relation)
	if relation.Type = "Association" or _
		relation.Type = "Aggregation" then
		'check aggregationKind
		if relation.SupplierEnd.Aggregation <> 0 _
			and relation.ClientEnd.Aggregation = 0 then
			'switch source and target
			'switch ID's
			dim tempID
			tempID = relation.ClientID
			relation.ClientID = relation.SupplierID
			relation.SupplierID = tempID
			'switch Ends
			switchRelationEnds relation
			'save relation
			relation.Update
		end if
	end if
end function

function switchRelationEnds (relation)
	dim tempVar
	tempvar = relation.ClientEnd.Aggregation
	relation.ClientEnd.Aggregation = relation.SupplierEnd.Aggregation
	relation.SupplierEnd.Aggregation       = tempvar
	tempvar = relation.ClientEnd.Alias
	relation.ClientEnd.Alias = relation.SupplierEnd.Alias
	relation.SupplierEnd.Alias             = tempvar
	tempvar = relation.ClientEnd.AllowDuplicates
	relation.ClientEnd.AllowDuplicates = relation.SupplierEnd.AllowDuplicates
	relation.SupplierEnd.AllowDuplicates   = tempvar
	tempvar = relation.ClientEnd.Cardinality
	relation.ClientEnd.Cardinality = relation.SupplierEnd.Cardinality
	relation.SupplierEnd.Cardinality       = tempvar
	tempvar = relation.ClientEnd.Constraint
	relation.ClientEnd.Constraint = relation.SupplierEnd.Constraint
	relation.SupplierEnd.Constraint        = tempvar
	tempvar = relation.ClientEnd.Containment
	relation.ClientEnd.Containment = relation.SupplierEnd.Containment
	relation.SupplierEnd.Containment       = tempvar
	tempvar = relation.ClientEnd.Derived
	relation.ClientEnd.Derived = relation.SupplierEnd.Derived
	relation.SupplierEnd.Derived           = tempvar
	tempvar = relation.ClientEnd.DerivedUnion
	relation.ClientEnd.DerivedUnion = relation.SupplierEnd.DerivedUnion
	relation.SupplierEnd.DerivedUnion      = tempvar
	tempvar = relation.ClientEnd.IsChangeable
	relation.ClientEnd.IsChangeable = relation.SupplierEnd.IsChangeable
	relation.SupplierEnd.IsChangeable      = tempvar
	tempvar = relation.ClientEnd.IsNavigable
	relation.ClientEnd.IsNavigable = relation.SupplierEnd.IsNavigable
	relation.SupplierEnd.IsNavigable       = tempvar
	tempvar = relation.ClientEnd.Navigable
	relation.ClientEnd.Navigable = relation.SupplierEnd.Navigable
	relation.SupplierEnd.Navigable         = tempvar
	tempvar = relation.ClientEnd.Ordering
	relation.ClientEnd.Ordering = relation.SupplierEnd.Ordering
	relation.SupplierEnd.Ordering          = tempvar
	tempvar = relation.ClientEnd.OwnedByClassifier
	relation.ClientEnd.OwnedByClassifier = relation.SupplierEnd.OwnedByClassifier
	relation.SupplierEnd.OwnedByClassifier = tempvar
	tempvar = relation.ClientEnd.Qualifier
	relation.ClientEnd.Qualifier = relation.SupplierEnd.Qualifier
	relation.SupplierEnd.Qualifier         = tempvar
	tempvar = relation.ClientEnd.Role
	relation.ClientEnd.Role = relation.SupplierEnd.Role
	relation.SupplierEnd.Role              = tempvar
	tempvar = relation.ClientEnd.RoleNote
	relation.ClientEnd.RoleNote = relation.SupplierEnd.RoleNote
	relation.SupplierEnd.RoleNote          = tempvar
	tempvar = relation.ClientEnd.RoleType
	relation.ClientEnd.RoleType = relation.SupplierEnd.RoleType
	relation.SupplierEnd.RoleType          = tempvar
	tempvar = relation.ClientEnd.Stereotype
	relation.ClientEnd.Stereotype = relation.SupplierEnd.Stereotype
	relation.SupplierEnd.Stereotype        = tempvar
	tempvar = relation.ClientEnd.StereotypeEx
	relation.ClientEnd.StereotypeEx = relation.SupplierEnd.StereotypeEx
	relation.SupplierEnd.StereotypeEx      = tempvar
'	tempvar = relation.ClientEnd.TaggedValues
'	relation.ClientEnd.TaggedValues = relation.SupplierEnd.TaggedValues
'	relation.SupplierEnd.TaggedValues      = tempvar
	tempvar = relation.ClientEnd.Visibility
	relation.ClientEnd.Visibility = relation.SupplierEnd.Visibility
	relation.SupplierEnd.Visibility        = tempvar
	relation.ClientEnd.Update
	relation.SupplierEnd.Update
end function