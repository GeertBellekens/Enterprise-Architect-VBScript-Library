'[path=\Projects\Project E\EA-Matic]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

'EA-Matic
'
' Script Name: 
' Author: Geert Bellekens
' Purpose: Registers the created user and date when creating a new Informationflow connector
' Date: 2023-11-03
'


''the event called by EA
'function EA_OnPostNewConnector(Info)
'	'Msgbox "EA_OnPostNewConnector triggered"
'	'get the connector id from the Info
'	dim connectorID2
'	connectorID2 = Info.Get("ConnectorID")
'	'Msgbox "ConnectorID: " & connectorID2
'	dim connector
'	set connector = Repository.GetConnectorByID(connectorID2)
'	'Msgbox "Steretoype: " & connector.Stereotype 
'	'check if this is an Informationflow
'	if connector.Stereotype = "Elia_InformationFlow" then
'		Msgbox "Yes, found one!"
'		'set tagged values for CreatedUser and CreatedDate
'		setTagValue connector, "CreatedUser", "Geert Bellekens"
'		setTagValue connector, "CreatedDate", now()
'	end if
'
'end function

'function EA_OnPostNewConnector2(connectorID)
'	Msgbox "EA_OnPostNewConnector triggered"
'	'get the connector id from the Info
'	dim connectorID
'	connectorID = Info.Get("ConnectorID")
'	msgbox "ConnectorID: " & connectorID
'	dim connector
'	set connector = Repository.GetConnectorByID(connectorID)
'	Msgbox "Steretoype: " & connector.Stereotype 
'	'check if this is an Informationflow
'	if connector.Stereotype = "Elia_InformationFlow" then
'		Msgbox "Yes, found one!"
'		'set tagged values for CreatedUser and CreatedDate
'		setTagValue connector, "CreatedUser", "Geert Bellekens"
'		setTagValue connector, "CreatedDate", now()
'	end if
'
'end function


function setTagValue(owner, tagname, tagValue)
	dim taggedValue as EA.TaggedValue
	dim returnTag as EA.TaggedValue
	set returnTag = nothing
	'check if a tag with that name alrady exists
	for each taggedValue in owner.TaggedValues
		if lcase(taggedValue.Name) = lcase(tagName) then
			set returnTag = taggedValue
			exit for
		end if
	next
	'create new one if not found
	if returnTag is nothing then
		set returnTag = owner.TaggedValues.AddNew(tagname,"")
	end if
	'set value
	returnTag.Value = tagValue
	returnTag.Update
end function


function test 
	EA_OnPostNewConnector2 55480
end function
'test