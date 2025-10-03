'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: SetDefaultColorOnLDMClass
' Author: Geert Bellekens
' Purpose: Sets the default color defined on the parent package on the LDM Class
' Date: 2025-10-03
'
'EA-Matic


function EA_OnPostNewElement(Info)
	dim elementID
    elementID = Info.Get("ElementID")
	dim element as EA.Element
	set element = Repository.GetElementByID(elementID)
	if not element.Stereotype = "LDM_Class" then
		exit function
	end if
	dim defaultColor
	defaultColor = getDefaultColor(element)
	if not defaultColor = -1 then
		element.SetAppearance 1, 0, defaultColor
		element.Update
	end if
end function

function getDefaultColor(element)
	dim defaultColor
	defaultColor = -1
	dim sqlGetData
	sqlGetData = "select                                                                                                            " & vbNewLine & _
				" coalesce(                                                                                                        " & vbNewLine & _
				"         po.Backcolor, o1.Backcolor, o2.Backcolor, o3.Backcolor, o4.Backcolor,                                    " & vbNewLine & _
				"         o5.Backcolor, o6.Backcolor, o7.Backcolor, o8.Backcolor, o9.Backcolor, o10.Backcolor                      " & vbNewLine & _
				"     ) as Backcolor                                                                                               " & vbNewLine & _
				" from t_object o                                                                                                  " & vbNewLine & _
				" inner join t_package p on p.Package_ID = o.Package_ID                                                            " & vbNewLine & _
				" left join t_object po on po.ea_guid = p.ea_guid                                                                  " & vbNewLine & _
				"                      and po.Backcolor <> -1                                                                      " & vbNewLine & _
				" left join t_package p1 on p1.Package_ID = p.Parent_ID                                                            " & vbNewLine & _
				" left join t_object o1 on o1.ea_guid = p1.ea_guid                                                                 " & vbNewLine & _
				"                       and o1.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p2 on p2.Package_ID = p1.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o2 on o2.ea_guid = p2.ea_guid                                                                 " & vbNewLine & _
				"                       and o2.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p3 on p3.Package_ID = p2.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o3 on o3.ea_guid = p3.ea_guid                                                                 " & vbNewLine & _
				"                       and o3.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p4 on p4.Package_ID = p3.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o4 on o4.ea_guid = p4.ea_guid                                                                 " & vbNewLine & _
				"                       and o4.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p5 on p5.Package_ID = p4.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o5 on o5.ea_guid = p5.ea_guid                                                                 " & vbNewLine & _
				"                       and o5.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p6 on p6.Package_ID = p5.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o6 on o6.ea_guid = p6.ea_guid                                                                 " & vbNewLine & _
				"                       and o6.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p7 on p7.Package_ID = p6.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o7 on o7.ea_guid = p7.ea_guid                                                                 " & vbNewLine & _
				"                       and o7.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p8 on p8.Package_ID = p7.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o8 on o8.ea_guid = p8.ea_guid                                                                 " & vbNewLine & _
				"                       and o8.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p9 on p9.Package_ID = p8.Parent_ID                                                           " & vbNewLine & _
				" left join t_object o9 on o9.ea_guid = p9.ea_guid                                                                 " & vbNewLine & _
				"                       and o9.Backcolor <> -1                                                                     " & vbNewLine & _
				" left join t_package p10 on p10.Package_ID = p9.Parent_ID                                                         " & vbNewLine & _
				" left join t_object o10 on o10.ea_guid = p10.ea_guid                                                              " & vbNewLine & _
				"                        and o10.Backcolor <> -1                                                                   " & vbNewLine & _
				" where o.Stereotype = 'LDM_Class'                                                                                 " & vbNewLine & _
				" and isnull(o.Backcolor, -1) <> coalesce(po.Backcolor, o1.Backcolor, o2.Backcolor, o3.Backcolor, o4.Backcolor,    " & vbNewLine & _
				" 							o5.Backcolor, o6.Backcolor, o7.Backcolor, o8.Backcolor, o9.Backcolor, o10.Backcolor)   " & vbNewLine & _
				" and o.Object_ID = " & element.ElementID &"                                                                       "
	dim result
	result = getSingleValueFromQuery(sqlGetData)
	if len(result) > 0 then
		defaultColor = Clng(result)
	end if
	'return
	getDefaultColor = defaultColor
end function