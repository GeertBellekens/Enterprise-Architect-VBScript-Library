'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: set default colors
' Author: Geert Bellekens
' Purpose: Set the default colors for all elements according to the default color on their parent package
' Date: 2025-10-03
'
const outPutName = "Set default colors"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	if package is nothing then
		exit sub
	end if
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for package '"& package.Name & "'", 0
	'do the work
	setDefaultcolors package
	'let user know
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for package '"& package.Name & "'", 0
end sub


function setDefaultcolors(package)
	dim idsAndColors
	set idsAndColors = getIDsAndColors(package)
	dim row
	for each row in idsAndColors
		dim objectID
		objectID = Clng(row(0))
		dim backColor
		backColor = Clng(row(1))
		dim element as EA.Element
		set element = Repository.GetElementByID(objectID)
		Repository.WriteOutput outPutName, now() & " Setting default color for element '" & element.Name  & "'", 0
		element.SetAppearance 1, 0, backColor
		element.update
	next
end function

function getIDsAndColors(package)
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select                                                                                                            " & vbNewLine & _
				" o.Object_ID                                                                                                      " & vbNewLine & _
				" ,coalesce(                                                                                                       " & vbNewLine & _
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
				" and o.Package_ID in (" & packageTreeIDString & ")                                                                "
	dim result
	set result = getArrayListFromQuery(sqlGetData)
	'return
	set getIDsAndColors = result
end function

main