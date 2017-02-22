'[path=\Projects\Project K\KING Scripts]
'[group=KING Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	'object tagged values
	dim objectTVsToDelete
	set objectTVsToDelete = CreateObject("System.Collections.ArrayList")
	addObjectTVsToDelete objectTVsToDelete 
	dim tvToDelete
	dim updateSQL
	for each tvToDelete in objectTVsToDelete
		updateSQL = "delete from t_objectproperties where [Property] = '" & tvToDelete & "'"
		Repository.Execute updateSQL
	next
	'attribute tagged values
	dim attributeTVsToDelete
	set attributeTVsToDelete = CreateObject("System.Collections.ArrayList")
	addAttributeTVsToDelete attributeTVsToDelete 
	for each tvToDelete in attributeTVsToDelete
		updateSQL = "delete from t_attributetag where [Property] = '" & tvToDelete & "'"
		Repository.Execute updateSQL
	next
	'connector tagged values
	dim connectorTVsToDelete
	set connectorTVsToDelete = CreateObject("System.Collections.ArrayList")
	addConnectorTVsToDelete connectorTVsToDelete 
	for each tvToDelete in connectorTVsToDelete
		updateSQL = "delete from t_connectortag where [Property] = '" & tvToDelete & "'"
		Repository.Execute updateSQL
	next
	msgbox "Finished"
end sub

sub addObjectTVsToDelete(objectTVsToDelete)
	objectTVsToDelete.Add "Aanduiding brondocument"
	objectTVsToDelete.Add "Aanduiding gebeurtenis"
	objectTVsToDelete.Add "Indicatie gebeurtenis"
end sub

sub addAttributeTVsToDelete(attributeTVsToDelete)
	attributeTVsToDelete.Add "Aanduiding brondocument"
	attributeTVsToDelete.Add "Aanduiding gebeurtenis"
	attributeTVsToDelete.Add "Indicatie gebeurtenis"
end sub

sub addConnectorTVsToDelete(connectorTVsToDelete)
	connectorTVsToDelete.Add "Aanduiding brondocument"
	connectorTVsToDelete.Add "Aanduiding gebeurtenis"
	connectorTVsToDelete.Add "Indicatie gebeurtenis"
	connectorTVsToDelete.Add "Naam terugrelatie"
end sub

main