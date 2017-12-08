'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	dim sql
	slq = " update nl set nl.value = o.name                                       "&_
			" from ((t_object o                                                     "&_
			" inner join t_objectproperties nl on nl.[Object_ID] = o.[Object_ID] )  "&_
			" inner join t_objectproperties fr on fr.[Object_ID] = o.[Object_ID] )  "&_
			" where o.stereotype = 'AtriasRequirement'	                            "&_
			" and nl.[Property] = 'Name NL'	                                        "&_
			" and fr.[Property] = 'Name FR'                                         "&_
			" and nl.VALUE is null                                                  "&_
			"                                                                       "&_
			" update fr set fr.value = o.ALIAS                                      "&_
			" from ((t_object o                                                     "&_
			" inner join t_objectproperties nl on nl.[Object_ID] = o.[Object_ID] )  "&_
			" inner join t_objectproperties fr on fr.[Object_ID] = o.[Object_ID] )  "&_
			" where o.stereotype = 'AtriasRequirement'                            "&_
			" and nl.[Property] = 'Name NL'	                                        "&_
			" and fr.[Property] = 'Name FR'                                         "&_
			" and fr.VALUE is null                                                  "&_
			"                                                                       "&_
			" update o set o.[Name] = o.pdata5                                      "&_
			" from ((t_object o                                                     "&_
			" inner join t_objectproperties nl on nl.[Object_ID] = o.[Object_ID] )  "&_
			" inner join t_objectproperties fr on fr.[Object_ID] = o.[Object_ID] )  "&_
			" where o.stereotype = 'AtriasRequirement'	                            "&_
			" and nl.[Property] = 'Name NL'	                                        "&_
			" and fr.[Property] = 'Name FR'                                         "&_
			" and o.[PDATA5] is not null                                            "
	Repository.Execute(sql)
	
end sub

sub test
	dim requirements as EA.Collection
	dim requirement as EA.Element
	dim sqlString
	sqlString = "select o.[Object_ID] from t_object o where o.stereotype = 'AtriasRequirement'"
	
	set requirements = Repository.GetElementSet(sqlString,2)
	for each requirement in Requirements
		
		requirement.StereotypeEx = "ATRIAS Requirements::Solution Requirement"
		requirement.Alias = ""
		if left(requirement.Name, 3) = "REQ" then
			requirement.Name = "SOL" & requirement.Name
		end if
		requirement.Update
		'copy the values of Name NL and Name FR to Title NL and Title FR
		dim taggedValue as EA.TaggedValue
		dim NameNLtag as EA.TaggedValue
		dim NameFRtag as EA.TaggedValue
		dim TitleNLtag as EA.TaggedValue
		dim TitleFRtag as EA.TaggedValue
		set NameNLTag = nothing
		set NameFRtag = nothing
		set TitleNLtag = nothing
		set TitleFRtag = nothing
		for each taggedValue in requirement.TaggedValues
			'session.output "aantal tagged values = " & requirement.TaggedValues.Count
			if taggedValue.Name = "Name NL" then
				set NameNLtag = taggedValue
			elseif taggedValue.Name = "Name FR" then
				set NameFRtag = taggedValue
			elseif taggedValue.Name = "Title NL" then
				set TitleNLtag = taggedValue
			elseif taggedValue.Name = "Title FR" then
				set TitleFRtag = taggedValue
			end if
		next
		if (not (NameNLtag is nothing))_
			and (not (NameFRtag is nothing))_
			and (not (TitleNLtag is nothing))_
			and (not (TitleFRtag is nothing)) then
			TitleNLtag.Value = NameNLtag.Value
			TitleNLtag.Update
			TitleFRtag.Value = NameFRtag.Value
			TitleFRtag.Update
		end if
	next
	'remove tags "Name NL"
	dim sqldelete
	sqldelete = " delete tv from t_objectproperties tv                                       "&_
					" where tv.[Property] = 'Name NL'                                            "&_
					" and exists (select tv2.[PropertyID] from t_objectProperties tv2 where      "&_
					"            tv2.[Property] = 'Title NL'                                     "&_
					"            and tv2.[Object_ID] = tv.Object_ID                              "&_
					"            and tv2.VALUE is not null)                                      "
	Repository.Execute sqldelete
	
	'remove tags "Name FR"
	sqldelete = " delete tv from t_objectproperties tv                                       "&_
					" where tv.[Property] = 'Name FR'                                            "&_
					" and exists (select tv2.[PropertyID] from t_objectProperties tv2 where      "&_
					"            tv2.[Property] = 'Title FR'                                     "&_
					"            and tv2.[Object_ID] = tv.Object_ID                              "&_
					"            and tv2.VALUE is not null)                                      "
	Repository.Execute sqldelete

	msgbox "Finished"
	'set requirement = Repository.GetElementByGuid("{36D0F895-E999-498a-9945-E9E036C9DAFF}")

end sub

sub correctImportance
	dim sqlNonEssential
	sqlNonEssential = " update tv set tv.VALUE = 'Non-essential' from t_objectproperties tv          "&_
				" inner join t_object o on tv.[Object_ID] = o.[Object_ID]                      "&_
				" where tv.[Property] = 'Importance'                                           "&_
				" and o.[Stereotype] = 'Solution Requirement'                                  "&_
				" and (tv.VALUE <> 'Must' or tv.VALUE is null)                                 "
	Repository.Execute sqlNonEssential
	
	dim sqlEssential
	sqlEssential = " update tv set tv.VALUE = 'Essential' from t_objectproperties tv              "&_
				" inner join t_object o on tv.[Object_ID] = o.[Object_ID]                      "&_
				" where tv.[Property] = 'Importance'                                           "&_
				" and o.[Stereotype] = 'Solution Requirement'                                  "&_
				" and tv.VALUE = 'Must'                                                        "
	Repository.Execute sqlEssential
	
end sub
correctImportance
'test
'main
'test
'main