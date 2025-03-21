'[path=\Projects\Project B\Baloise Scripts]
'[group=Baloise Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Fix Broken Links
' Author: Geert Bellekens
' Purpose: Search for broken SourceAttribute tagged value links and fix them where possible
' Date: 2021-03-26
'

const outPutName = "Fix Broken Links"

function main()
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	if package is nothing then
		Msgbox "Please select a package in the project browser before executing this script"
		exit function
	end if
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting Fix Broken links for package '" & package.Name & "'", 0
	'do the actual work
	fixBrokenLinks package
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished Fix Broken links for package '" & package.Name & "'", 0
end function

function fixBrokenLinks(package)
	'get the attributes with broken links
	dim attributes
	set attributes = getAttributesWithBrokenLinks(package)
	dim attribute as EA.Attribute
	for each attribute in attributes
		'find correct source guid
		dim sqlGetData 
		sqlGetData = "select a2.ea_guid from t_attribute a                               " & vbNewLine & _
						" inner join t_object o on o.Object_ID = a.Object_ID                " & vbNewLine & _
						" inner join t_objectproperties tv on tv.Object_ID = o.Object_ID    " & vbNewLine & _
						" 								and tv.Property = 'SourceElement'   " & vbNewLine & _
						" inner join t_object o2 on o2.ea_guid = tv.Value                   " & vbNewLine & _
						" inner join t_attribute a2 on a2.Object_ID = o2.Object_ID          " & vbNewLine & _
						" 							and a2.Name = a.Name                    " & vbNewLine & _
						" where a.ea_guid = '" & attribute.AttributeGUID & "'               "
		dim results
		set results = getArrayListFromQuery(sqlGetData)
		dim fixed
		fixed = false
		if results.count > 0 then
			dim targetGUID
			targetGUID = results(0)(0)
			if len(targetGUID) > 0 then
				dim taggedValue as EA.TaggedValue
				for each taggedValue in attribute.TaggedValues
					if lcase(taggedValue.Name) = "sourceattribute" then
						taggedValue.Value = targetGUID
						taggedValue.Update
						fixed = true
						Repository.WriteOutput outPutName, now() & " Fixed attribute '" & attribute.Name & "' with GUID: " & attribute.AttributeGUID, 0
					end if
				next
			end if
		end if
		if not fixed then
			Repository.WriteOutput outPutName, now() & " ERROR: Could not fix attribute '" & attribute.Name & "' with GUID: " & attribute.AttributeGUID, 0
		end if
	next
end function

function getAttributesWithBrokenLinks(package)
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "SELECT a.ID                                                                      " & vbNewLine & _
				" FROM t_attribute a                                                              " & vbNewLine & _
				" inner join t_object o on o.Object_ID = a.Object_ID                              " & vbNewLine & _
				" inner join t_attributetag tv on tv.ElementID = a.ID                             " & vbNewLine & _
				" 								and tv.Property = 'sourceAttribute'               " & vbNewLine & _
				" left join t_xref as xref on xref.Client = a.ea_guid                             " & vbNewLine & _
				" 							and xref.name = 'Stereotypes'                         " & vbNewLine & _
				" WHERE 1=1                                                                       " & vbNewLine & _
				" and (xref.Description not like '%redefine%'  or xref.Client is null)            " & vbNewLine & _
				" and not exists (select oo.ID from t_attribute oo where oo.ea_guid = tv.Value)   " & vbNewLine & _
				" and (o.object_type = 'Enumeration' or o.Stereotype like 'XSD%')                 " & vbNewLine & _
				" AND o.Package_ID IN (" & packageTreeIDs & ")                                    "
	dim attributes
	set attributes = getAttributesFromQuery(sqlGetData)
	'return
	set getAttributesWithBrokenLinks = attributes
end function

main