'[path=\Projects\Project K\KING Scripts]
'[group=KING Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Rename Stereotypes
' Author: Geert Bellekens
' Purpose: Rename stereotypes on all types of elements 
' Environment: Tested on .eap file.
' Date: 13/10/2015
'
sub main
	renameElementStereotypes "Groepattribuutsoort", "Gegevensgroeptype"
	renameAttributeStereotypes "Referentiegegeven", "Referentie element"
'	renameConnectorStereotypes "FromStereo", "ToStereo"
'	renameOperationStereotypes "FromStereo", "ToStereo"
'	renameDiagramStereotypes "FromStereo", "ToStereo"
	Repository.RefreshModelView(0)
	msgbox "Finished renaming stereotypes"
end sub

sub renameElementStereotypes(fromStereo, toStereo)
	renameStereotypes "t_object", fromStereo, toStereo
end sub
sub renameAttributeStereotypes(fromStereo, toStereo)
	renameStereotypes "t_attribute", fromStereo, toStereo
end sub
sub renameConnectorStereotypes(fromStereo, toStereo)
	renameStereotypes "t_connector", fromStereo, toStereo
end sub
sub renameOperationStereotypes(fromStereo, toStereo)
	renameStereotypes "t_operation", fromStereo, toStereo
end sub
sub renameDiagramStereotypes(fromStereo, toStereo)
	renameStereotypes "t_diagram", fromStereo, toStereo
end sub

sub renameStereotypes (baseTable, fromStereo, toStereo)
	dim updateSQL
	'first the second part of of t_xref description
	updateSQL = "update (" & baseTable & " o inner join t_xref x on o.[ea_guid] = x.[Client]) "&_
			   " set x.Description = MID( x.Description, 1, INSTR(  x.Description, ':" & fromStereo & "') - 1) "&_
						 " + ':" & toStereo & "' "&_
						 " + MID(x.Description,INSTR(  x.Description, ':" & fromStereo & "') "&_ 
							  " + LEN(':" & fromStereo & "'), LEN(x.Description)  "&_
							  " - INSTR(  x.Description, ':" & fromStereo & "') "&_
							  " - LEN(':" & fromStereo & "')+ 1) "&_
			   " where o.Stereotype = '" & fromStereo & "' "&_
			   "  and x.Name = 'Stereotypes' "&_
			   " and INSTR(  x.Description, ':" & fromStereo & "') > 0  "			   
	Repository.Execute updateSQL
	'then the first part of t_xref description
	updateSQL = "update (" & baseTable & " o inner join t_xref x on o.[ea_guid] = x.[Client]) "&_
			   " set x.Description = MID( x.Description, 1, INSTR(  x.Description, '=" & fromStereo & "') - 1) "&_
						 " + '=" & toStereo & "' "&_
						 " + MID(x.Description,INSTR(  x.Description, '=" & fromStereo & "') "&_ 
							  " + LEN('=" & fromStereo & "'), LEN(x.Description)  "&_
							  " - INSTR(  x.Description, '=" & fromStereo & "') "&_
							  " - LEN('=" & fromStereo & "')+ 1) "&_
			   " where o.Stereotype = '" & fromStereo & "' "&_
			   "  and x.Name = 'Stereotypes' "&_			   
			   " and INSTR(  x.Description, '=" & fromStereo & "') > 0  "
	Repository.Execute updateSQL				
	'then the stereotype itself
	updateSQL = " update " & baseTable & " o "&_
			    " set o.[Stereotype] = '" & toStereo & "' "&_
	            " where o.Stereotype = '" & fromStereo & "' "
	Repository.Execute updateSQL
end sub

main