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
	'unlink the notes linked to the attributes on the stereotypes in order to preserve the information
	dim sqlUnlinkNotes
	sqlUnlinkNotes ="update ((t_object o "&_
					" inner join t_attribute a on a.[ID] like o.[PDATA2]) "&_
					" inner join t_object p on p.Object_ID = a.Object_ID) "&_
					" set o.[PDATA1] = null, o.[PDATA2] = null, o.[PDATA3] = null, o.[PDATA4] = null "&_
					" where o.[Object_Type] = 'Note' "&_
					" and p.[Stereotype] = 'stereotype' "
	Repository.Execute sqlUnlinkNotes
	'remove the notes from the attributes on the stereotypes
	dim sqlUpdate
	sqlUpdate = "update (t_attribute t "&_
				" inner join t_object o on o.Object_ID = t.Object_ID ) "&_
				" set t.[NOTES] = null   "&_
				" where o.stereotype = 'stereotype' "
	'Session.Output sqlUpdate
	Repository.Execute sqlUpdate
end sub

main