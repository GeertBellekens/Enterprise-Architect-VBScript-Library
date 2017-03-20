'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

' EA-Matic
' Script Name: SyncronizeOverrides
' Author: Geert Bellekens
' Purpose: Keeps the signature of the overridden operation in synch with that of the parent operation.
'          Every time an operation is changed we ask the user if he wants to synchronize the overrides
' Date: 09/02/2015
'

'remember the operation being edited
dim operationID 
'remember the list of overridden operations
dim overrides
'output tab
Repository.CreateOutputTab "EA-Matic"
'Repository.EnsureOutputVisible "EA-Matic"

'Event Called when a new element is selected in the context. We use this operation to keep the id of the selected operation and a list of its overrides
'Because now is the only moment we are able to find it's overrides. Once changed we cannot find the overrides anymore because then they already
'have a different signature
function EA_OnContextItemChanged(GUID, ot)
	'we only want to do something when the selected element is an operation
	if ot = otMethod then
		'get the model
		dim model 
		set model = getEAAddingFrameworkModel()
		'get the operation
		dim operation
		set operation = model.getOperationByGUID(GUID)
		'remember the operationID
		operationID = operation.id
		'remember the overrides
		set overrides = getOverrides(operation, model)
		Repository.WriteOutput "EA-Matic", overrides.Count & " overrides found for: " & operation.name,0
	end if
end function

'Event called when an element is changed. Unfortunately EA doesn't call it for an operation, only for the owner so we have to work with that.
function EA_OnNotifyContextItemModified(GUID, ot)
	'we only want to do something when the selected element is an operation
	if ot = otElement then		
		'get the operation
		'Here we use the EA API object directly as most set methods are not implemented in EA Addin Framework
		dim wrappedOperation
		set wrappedOperation = Repository.GetMethodByID(operationID)
		dim modifiedElement
		set modifiedElement = Repository.GetElementByGuid(GUID)
		if not wrappedOperation is Nothing and not modifiedElement is Nothing then
			'check to be sure we have the same operation
			if modifiedElement.ElementID = wrappedOperation.ParentID AND overrides.Count > 0 then
				dim synchronizeYes
				synchronizeYes = MsgBox("Found " & overrides.Count & " override(s) for operation "& modifiedElement.Name & "." & wrappedOperation.Name & vbNewLine & "Synchronize?" _
										,vbYesNo or vbQuestion or vbDefaultButton1, "Synchronize overrides?")
				if synchronizeYes = vbYes then
					synchronizeOverrides wrappedOperation
					'log to output
					
					Repository.WriteOutput "EA-Matic", "Operation: " & wrappedOperation.name &" synchronized" ,0
				end if
				'reset operationID to avoid doing it all again
				operationID = 0
			end if
		end if
	end if
end function

'Synchronizes the operation with it's overrides
function synchronizeOverrides(wrappedOperation)
	dim override
	for each override in overrides
		dim wrappedOverride 
		set wrappedOverride = override.WrappedOperation
		'synchronize the operation with the override
		synchronizeOperation wrappedOperation, wrappedOverride
		'tell EA something might have changed
		Repository.AdviseElementChange wrappedOverride.ParentID
	next
end function

'Synchronizes the operation with the given override
function synchronizeOperation(wrappedOperation, wrappedOverride)
	dim update 
	update = false
	'check name
	if wrappedOverride.Name <> wrappedOperation.Name then
		wrappedOverride.Name = wrappedOperation.Name
		update = true
	end if
	'check return type
	if wrappedOverride.ReturnType <> wrappedOperation.ReturnType then
		wrappedOverride.ReturnType = wrappedOperation.ReturnType
		update = true
	end if
	'check return classifier
	if wrappedOverride.ReturnType <> wrappedOperation.ReturnType then
		wrappedOverride.ReturnType = wrappedOperation.ReturnType
		update = true
	end if
	if update then
		wrappedOverride.Update
	end if
	'check parameters
	synchronizeParameters wrappedOperation, wrappedOverride
end function

'Synchronizes the parameters of the given operatin with that of the overrride
function synchronizeParameters(wrappedOperation, wrappedOverride)
	'first make sure they both have the same number of parameters
	if wrappedOverride.Parameters.Count < wrappedOperation.Parameters.Count then
		'add parameters as required
		dim i
		for i = 0 to wrappedOperation.Parameters.Count - wrappedOverride.Parameters.Count -1
			dim newParameter 
			set newParameter = wrappedOverride.Parameters.AddNew("parameter" & i,"")
			newParameter.Update
		next
		wrappedOverride.Parameters.Refresh
	elseif wrappedOverride.Parameters.Count > wrappedOperation.Parameters.Count then
		'remove parameters as required
		for i = wrappedOverride.Parameters.Count -1 to wrappedOperation.Parameters.Count step -1
			wrappedOverride.Parameters.DeleteAt i,false 
		next
		wrappedOverride.Parameters.Refresh
	end if
	'make parameters equal
	dim wrappedParameter 
	dim overriddenParameter
	dim j
	for j = 0 to wrappedOperation.Parameters.Count -1
		dim parameterUpdated
		parameterUpdated = false
		set wrappedParameter = wrappedOperation.Parameters.GetAt(j)
		set overriddenParameter = wrappedOverride.Parameters.GetAt(j)
		'name
		if overriddenParameter.Name <> wrappedParameter.Name then
			overriddenParameter.Name = wrappedParameter.Name 
			parameterUpdated = true
		end if
		'type
		if overriddenParameter.Type <> wrappedParameter.Type then
			overriddenParameter.Type = wrappedParameter.Type 
			parameterUpdated = true
		end if
		'default
		if overriddenParameter.Default <> wrappedParameter.Default then
			overriddenParameter.Default = wrappedParameter.Default 
			parameterUpdated = true
		end if
		'kind
		if overriddenParameter.Kind <> wrappedParameter.Kind then
			overriddenParameter.Kind = wrappedParameter.Kind 
			parameterUpdated = true
		end if
		'classifier
		if overriddenParameter.ClassifierID <> wrappedParameter.ClassifierID then
			overriddenParameter.ClassifierID = wrappedParameter.ClassifierID 
			parameterUpdated = true
		end if
		'update the parameter if it was changed
		if parameterUpdated then
			overriddenParameter.Update
		end if
	next
end function

'gets the overrides of the given operation by first getting all operations with the same signature and then checking if they are owned by a descendant
function getOverrides(operation, model)
	'first get all operations with the exact same signature
	dim overrideQuery
	overrideQuery = "select distinct op2.OperationID from (((t_operation op " & _
					"inner join t_operation op2 on op2.[Name] = op.name) "& _
					"left join t_operationparams opp on op.OperationID = opp.OperationID) "& _
					"left join t_operationparams opp2 on opp2.OperationID = op2.OperationID) "& _
					"where op.OperationID = "& operation.id &" "& _
					"and op2.ea_guid <> op.ea_guid "& _
					"and (op2.TYPE = op.Type OR (op2.TYPE is null AND op.Type is null)) "& _
					"and (op2.Classifier = op.Classifier OR (op2.Classifier is null AND op.Classifier is null)) "& _
					"and (opp.Name = opp2.Name OR (opp.Name is null AND opp2.Name is null)) "& _
					"and (opp.TYPE = opp2.TYPE OR (opp.TYPE is null AND opp2.Type is null)) "& _
					"and (opp.DEFAULT = opp2.DEFAULT OR (opp.DEFAULT is null AND opp2.DEFAULT is null)) "& _
					"and (opp.Kind = opp2.Kind OR (opp.Kind is null AND opp2.Kind is null)) "& _
					"and (opp.Classifier = opp2.Classifier OR (opp.Classifier is null AND opp2.Classifier is null)) "
	dim candidateOverrides
	set candidateOverrides = model.ToArrayList(model.getOperationsByQuery(overrideQuery))
	'then get the descendants of the owner 
	dim descendants
	dim descendant
	'first find all elements that either inherit from the owner or realize it
	dim owner
	set owner = model.toObject(operation.owner)
	set descendants = getDescendants(owner, model)
	'then filter the candidates to only those of the descendants
	'loop operations backwards
	dim i
	for i = candidateOverrides.Count -1 to 0 step -1
		dim found 
		found = false
		for each descendant in descendants
			if descendant.id = model.toObject(candidateOverrides(i).owner).id then
				'owner is a descendant, operation can stay
				found = true
				exit for
			end if
		next
		'remove operation from non descendants
		if not found then
			candidateOverrides.RemoveAt(i)
		end if
	next
	set getOverrides = candidateOverrides
end function

'gets all descendant of an element. That is all subclasses and classes that Realize the element.
'Works recursively to get them all.
function getDescendants(element, model)
	dim descendants
	dim getdescendantsQuery 
	getdescendantsQuery = "select c.Start_Object_ID as Object_ID from (t_object o " _
					& "inner join t_connector c on c.End_Object_ID = o.Object_ID) " _
					& "where "_
					& "(c.[Connector_Type] like 'Generali_ation' "_
					& "or c.[Connector_Type] like 'Reali_ation' )"_
					& "and o.Object_ID = " & element.id
	set descendants = model.toArrayList(model.getElementWrappersByQuery(getdescendantsQuery))
	'get the descendants descendants as well
	dim descendant
	dim descendantsChildren
	for each descendant in descendants
		if IsEmpty(descendantsChildren) then
			set descendantsChildren = getDescendants(descendant, model)
		else
			descendantsChildren.AddRange(getDescendants(descendant, model))
		end if
	next
	'add the descendantsChildren to the descendants
	if not IsEmpty(descendantsChildren) then
		if  descendantsChildren.Count > 0 then
			descendants.AddRange(descendantsChildren)
		end if
	end if
	set getDescendants = descendants
end function

'gets a new instance of the EAAddinFramework and initializes it with the EA.Repository
function getEAAddingFrameworkModel()
	'Initialize the EAAddinFramework model
    dim model 
    set model = CreateObject("TSF.UmlToolingFramework.Wrappers.EA.Model")
    model.initialize(Repository)
	set getEAAddingFrameworkModel = model
end function