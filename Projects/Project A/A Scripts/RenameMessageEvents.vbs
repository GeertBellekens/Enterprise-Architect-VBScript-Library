'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Atrias Scripts.LinkToCRMain
'
' Script Name: RenameMessageEvents
' Author: Geert Bellekens
' Purpose: Rename the Message Events to "Send" or "Receive" + the message name
' Date: 2016-11-23
'

const outPutName = "Rename Message Events"
const CRGUID = "{5638B62F-08E5-46bc-8C61-6DE16D3017BD}"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, "Starting Renaming Message Events " & now(), 0
	'start processing
	dim sqlGetEvents
	sqlGetEvents = "select o.Object_ID , case when o.Object_ID = c.Start_Object_ID then 'Send ' else 'Receive ' end + msg.Name AS NewName" & _
					" from t_object o " & _
					" inner join t_connector c on o.Object_ID in (c.Start_Object_ID, c.End_Object_ID) " & _
					" inner join t_connectortag ctv on ctv.ElementID = c.Connector_ID " & _
					" 							and ctv.Property = 'MessageRef' " & _
					" inner join t_object fis on fis.ea_guid = ctv.VALUE " & _
					" 						and fis.Stereotype = 'Message' " & _
					" inner join t_connector fis_msg on fis.Object_ID = fis_msg.End_Object_ID " & _
					" 							and fis_msg.Connector_Type in ('Realization','Realisation') " & _
					" inner join t_object msg on msg.Object_ID = fis_msg.Start_Object_ID " & _
					" 							and msg.Stereotype = 'Message' " & _
					" where o.Object_Type = 'Event' " & _
					" and o.name <> case when o.Object_ID = c.Start_Object_ID then 'Send ' else 'Receive ' end + msg.Name "
						
	dim queryResult 
	queryResult = Repository.SQLQuery(sqlGetEvents)
	dim eventResults
	eventResults = convertQueryResultToArray(queryResult)
	dim i
	Session.Output queryResult
	for i = 0 to Ubound(eventResults)
		'get the event ID
		dim eventID 
		eventID = eventResults(i,0)
		dim eventElement as EA.Element
		if eventID > 0 then
		set eventElement = Repository.GetElementByID(eventID)
			if not eventElement is nothing then
				'get the new name
				dim newName,oldname 
				newName = eventResults(i,1)
				oldName = eventElement.Name 
				if len(newName) > 0 then
					'log
					Repository.WriteOutput outPutName, "Renaming " & oldName & " to " & newName , 0
					eventElement.Name = newName
					eventElement.Update
					'add CR tag
					Repository.WriteOutput outPutName, "Adding CR tag to " & newName , 0
					addCrTag eventElement, oldName, newName 
				end if
			end if
		end if
	next
	'set timestamp
	Repository.WriteOutput outPutName, "Finished Renaming Message Events " & now(), 0
end sub

function addCrTag(eventElement,oldName,newName)
	dim crToUse, selectedItemType, userLogin, comments
	set crToUse = Repository.GetElementByGuid(CRGUID)
	if not crToUse is nothing then
		userLogin = getUserLogin
		selectedItemType = eventElement.ObjectType
		comments = "Name changed from '" & oldName & "' to '" & newName & "'"
		linkToCR eventElement, selectedItemType, CRToUse, userLogin, comments
	end if
end Function

main