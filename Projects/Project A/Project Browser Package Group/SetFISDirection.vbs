'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
dim outputTabName
outputTabName = "FISDirections"

sub main
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	Repository.WriteOutput outputTabName, now() & ": Starting Setting FIS direction",0
	'get selected package
	dim selectedPackage
	set selectedPackage = Repository.GetTreeSelectedPackage
	if not selectedPackage is nothing then
		'get tree of package
		dim packageTree
		set packageTree = getPackageTree(selectedPackage)
		'get id list
		dim packageIDList
		packageIDList =	makePackageIDString(packageTree)
		'make sql query
		dim sqlGetAllMessages
		sqlGetAllMessages = "select o.Object_ID from t_object o " & _
							" where o.Stereotype in ('FIS', 'Message') " & _
							" and o.Package_ID in (" & packageIDList & ")"
		dim allMessages
		set allMessages = Repository.GetElementSet(sqlGetAllMessages,2)
		dim message as EA.Element
		for each message in allMessages
			Repository.WriteOutput outputTabName, "Processing: " & message.Name,0
			dim directionTV as EA.TaggedValue
			set directionTV = getOrCreateTaggedValue (message, "Atrias::Direction")
			directionTV.Value = getMessageDirection(message)
			directionTV.Update
'			message.StereotypeEx = "Message"
'			message.Update
'			'add tagged value if not exists yet
'			getOrCreateTaggedValue message, "Atrias::Direction"
'			'remove tagged value Direction
'			removeTaggedValue message, "Direction"
		next
	end if
	Repository.WriteOutput outputTabName, now() & ": Finished setting FIS direction",0
end sub

function getMessageDirection (message)
	dim messageFlows 
	set messageFlows = findMessageFlows(message)
	dim messageFlow as EA.Connector
	dim direction
	dim previousDirection
	previousDirection = ""
	for each messageFlow in messageflows
		direction = getMessageFlowDirection (messageFlow)
		if previousDirection = "" then
			previousDirection = direction
		else
			if previousDirection <> direction and direction <> "" then
				direction = "InOut"
				exit for
			end if
		end if
	next
	getMessageDirection = direction
end function

function getMessageFlowDirection (messageFlow)
	'default empty
	getMessageFlowDirection = ""
	dim source as EA.Element
	set source = Repository.GetElementByID(messageFlow.ClientID)
	dim poolClassifier as EA.Element
	set poolClassifier = getPoolClassifier(source)
	if not poolClassifier is nothing then
		if poolClassifier.Name = "Central Market System" then
			getMessageFlowDirection = "Out"
		else
			getMessageFlowDirection = "In"
		end if
	else
		'try with target
		dim target as EA.Element
		set target = Repository.GetElementByID(messageFlow.SupplierID)
		if not poolClassifier is nothing then
			if poolClassifier.Name = "Central Market System" then
				getMessageFlowDirection = "In"
			else
				getMessageFlowDirection = "Out"
			end if
		end if
	end if	
end function

function getPoolClassifier(element)
	'dim element as EA.Element
	dim parent  as EA.Element
	if not element is nothing then
		if element.Type = "ActivityPartition"  and element.Stereotype = "Pool" then
			set getPoolClassifier = Repository.GetElementByID(element.ClassfierID)
		elseif element.type = "Package" or element.ParentID = 0 then 
			set getPoolClassifier = nothing
		else
			set parent = Repository.GetElementByID(element.ParentID)
			set getPoolClassifier = getPoolClassifier(parent)
		end if
	else
		set getPoolClassifier = nothing
	end if
end function

function findMessageFlows(message)
	dim sqlQuery
	sqlQuery = 	"select ctv.ElementID as Connector_ID from t_connectortag ctv " & _
				" where ctv.Property = 'messageRef' " & _
				" and ctv.VALUE = '" & message.ElementGUID & "'"
	set findMessageFlows = getConnectorsFromQuery(sqlQuery)
end function

function getOrCreateTaggedValue(element, taggedValueName)
	'add tagged value if not exists yet
	dim taggedValue as EA.TaggedValue
	dim taggedValueExists
	taggedValueExists = false
	for each taggedValue in element.TaggedValues
		if taggedValue.Name = taggedValueName then
			taggedValueExists = true
			exit for
		end if
	next
	'create tagged value is not existing yet
	if taggedValueExists = false then
		set taggedValue = element.TaggedValues.AddNew(taggedValueName,"")
		taggedValue.Update
	end if
	set getOrCreateTaggedValue = taggedValue
end function

function removeTaggedValue(element, taggedValueName)
'	dim taggedValue as EA.TaggedValue
'	dim i
'	'loop tagged values
'	for i = 0 to element.TaggedValues.Count -1
'		set taggedValue = element.TaggedValues.Getat(i)
'		if taggedValue.Name = taggedValueName then
'			element.TaggedValues.DeleteAt i, false
'			exit for
'		end if
'	next
	'use dirty SQL delete for performance
	dim sqlDelete
	sqlDelete =	"delete from  t_objectproperties " & _
		" where Object_ID = " & element.ElementID & _
		" and Property = 'Direction'"
	Repository.Execute sqlDelete
end function

main