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
	' TODO: Enter script code here!
	
	updateXref "Attribuutsoort"
	updateXref "Complex datatype"
	updateXref "Data element"
	updateXref "External"
	updateXref "Externe koppeling"
	updateXref "Gegevensgroep compositie"
	updateXref "Gegevensgroeptype"
	updateXref "Generalisatie"
	updateXref "Objecttype"
	updateXref "Referentie element"
	updateXref "Referentielijst"
	updateXref "Relatieklasse"
	updateXref "Relatiesoort"
	updateXref "Tekentechnisch"
	updateXref "Union"
	updateXref "Union element"
	updateXref "View"

end sub

function updateXref(stereotype)

	dim sqlupdate 
	sqlupdate = "update  t_xref  set description = '@STEREO;Name=" & stereotype & ";FQName=MIG::" & stereotype & ";@ENDSTEREO;'" & _
				" where [Description]  like '@STEREO;Name=" & stereotype & "*'"
	'Session.Output sqlupdate
	Repository.Execute sqlupdate
	
end function

main