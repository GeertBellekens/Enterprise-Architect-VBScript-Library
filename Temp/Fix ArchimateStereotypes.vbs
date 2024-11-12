'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
 dim updateSQL
 updateSQL = "update o set o.Stereotype = replace(o.stereotype,'Archimate3::','')    " & vbNewLine & _
    " from t_object o                                                       " & vbNewLine & _
    " left join t_xref x on x.client = o.ea_guid                            " & vbNewLine & _
    "   and x.name = 'Stereotypes'                                      " & vbNewLine & _
    " where o.stereotype like 'Archimate3::%'                               "
    Repository.Execute updateSQL
end sub

main