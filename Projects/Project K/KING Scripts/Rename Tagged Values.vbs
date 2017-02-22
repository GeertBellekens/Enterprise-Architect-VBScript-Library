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
	dim objectTVNamePairs
	set objectTVNamePairs = CreateObject("System.Collections.ArrayList")
	addObjectTVNamePairs objectTVNamePairs 
	dim tvpair
	dim updateSQL
	for each tvpair in objectTVNamePairs
		updateSQL = "update t_objectproperties set [Property] = '" & tvpair(1) & "' where [Property] = '" & tvpair(0) & "'"
		Repository.Execute updateSQL
	next
	'attribute tagged values
	dim attributeTVNamePairs
	set attributeTVNamePairs = CreateObject("System.Collections.ArrayList")
	addAttributeTVNamePairs attributeTVNamePairs
	for each tvpair in attributeTVNamePairs
		updateSQL = "update t_attributetag set [Property] = '" & tvpair(1) & "' where [Property] = '" & tvpair(0) & "'"
		Repository.Execute updateSQL
	next
	'connector tagged values
	dim connectorTVNamePairs
	set connectorTVNamePairs = CreateObject("System.Collections.ArrayList")
	addConnectorTVNamePairs connectorTVNamePairs
	for each tvpair in connectorTVNamePairs
		updateSQL = "update t_connectortag set [Property] = '" & tvpair(1) & "' where [Property] = '" & tvpair(0) & "'"
		Session.Output updateSQL
		Repository.Execute updateSQL
	next
end sub

sub addObjectTVNamePairs(objectTVNamePairs)
	objectTVNamePairs.Add Array("Code objecttype","Code")
	objectTVNamePairs.Add Array("Datum opname objecttype","Datum opname")
	objectTVNamePairs.Add Array("Herkomst definitie objecttype","Herkomst definitie")
	objectTVNamePairs.Add Array("Herkomst objecttype","Herkomst")
	objectTVNamePairs.Add Array("Kwaliteitsbegrip objecttype","Kwaliteitsbegrip")
	objectTVNamePairs.Add Array("Populatie objecttype","Populatie")
	objectTVNamePairs.Add Array("Toelichting objecttype","Toelichting")
	objectTVNamePairs.Add Array("Herkomst attribuutsoort","Herkomst")
	objectTVNamePairs.Add Array("Code attribuutsoort","Code")
	objectTVNamePairs.Add Array("Herkomst definitie attribuutsoort","Herkomst definitie")
	objectTVNamePairs.Add Array("Datum opname attribuutsoort","Datum opname")
	objectTVNamePairs.Add Array("Toelichting attribuutsoort","Toelichting")
	objectTVNamePairs.Add Array("Regels attribuutsoort","Regels")
	objectTVNamePairs.Add Array("Datum opname referentielijst","Datum opname")
	objectTVNamePairs.Add Array("Herkomst definitie referentielijst","Herkomst definitie")
	objectTVNamePairs.Add Array("Herkomst referentielijst","Herkomst")
	objectTVNamePairs.Add Array("Toelichting referentielijst","Toelichting")
	objectTVNamePairs.Add Array("Code referentielijst","Code")
	objectTVNamePairs.Add Array("Datum opname union","Datum opname")
	objectTVNamePairs.Add Array("Herkomst union","Herkomst")
	objectTVNamePairs.Add Array("Aanduiding  strijdigheid/nietigheid","Aanduiding strijdigheid/nietigheid")
end sub

sub addAttributeTVNamePairs(attributeTVNamePairs)
	attributeTVNamePairs.Add Array("Herkomst attribuutsoort","Herkomst")
	attributeTVNamePairs.Add Array("Code attribuutsoort","Code")
	attributeTVNamePairs.Add Array("Herkomst definitie attribuutsoort","Herkomst definitie")
	attributeTVNamePairs.Add Array("Datum opname attribuutsoort","Datum opname")
	attributeTVNamePairs.Add Array("Toelichting attribuutsoort","Toelichting")
	attributeTVNamePairs.Add Array("Waardenverzameling","Patroon")
	attributeTVNamePairs.Add Array("Regels attribuutsoort","Regels")
	attributeTVNamePairs.Add Array("Herkomst referentiegegeven","Herkomst")
	attributeTVNamePairs.Add Array("Code referentiegegeven","Code")
	attributeTVNamePairs.Add Array("Herkomst definitie referentiegegeven","Herkomst definitie")
	attributeTVNamePairs.Add Array("Datum opname referentiegegeven","Datum opname")
	attributeTVNamePairs.Add Array("Toelichting referentiegegeven","Toelichting")
	attributeTVNamePairs.Add Array("Datum opname union element","Datum opname")
	attributeTVNamePairs.Add Array("Herkomst union element","Herkomst")
	attributeTVNamePairs.Add Array("Aanduiding  strijdigheid/nietigheid","Aanduiding strijdigheid/nietigheid")
end sub

sub addConnectorTVNamePairs(connectorTVNamePairs)
	connectorTVNamePairs.Add Array("Herkomst relatiesoort","Herkomst")
	connectorTVNamePairs.Add Array("Code relatiesoort","Code")
	connectorTVNamePairs.Add Array("Herkomst definitie relatiesoort","Herkomst definitie")
	connectorTVNamePairs.Add Array("Datum opname relatiesoort","Datum opname")
	connectorTVNamePairs.Add Array("Toelichting relatiesoort","Toelichting")
	connectorTVNamePairs.Add Array("Regels relatiesoort","Regels")
	connectorTVNamePairs.Add Array("Aanduiding  strijdigheid/nietigheid","Aanduiding strijdigheid/nietigheid")
end sub

main