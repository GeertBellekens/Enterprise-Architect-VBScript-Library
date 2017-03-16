'[path=\Projects\Project E\Element Group]
'[group=Element Group]

!INC Local Scripts.EAConstants-VBScript

sub main
	dim selectedSchema as EA.Element
	set selectedSchema = Repository.GetContextObject()
	if selectedSchema.ObjectType = otElement then
		if selectedSchema.Type = "Artifact" then
			msgbox "Please select the schema artifact to copy"
			dim schemaID
			dim schemaToCopy as EA.Element
			schemaID = Repository.InvokeConstructPicker("IncludedTypes=Artifact") 
			if schemaID > 0 then
				set schemaToCopy = Repository.GetElementByID(schemaID)
				dim response
				response = Msgbox("copy the schema of """ & schemaToCopy.Name & """ to """ & selectedSchema.Name & """ ?", vbYesNo+vbQuestion, "Copy Schema?")
				'check the response
				if response = vbYes then
					dim sqlUpdate
					sqlupdate = "update t_document set StrContent =  " & _
								"	(select replace(convert(varchar(max),d.StrContent),'<description name=""" & schemaToCopy.Name & """','<description name=""" & selectedSchema.Name & """') " & _
								"	from t_document d  " & _
								"	where d.ElementType = 'SC_MessageProfile' " & _
								"	and  d.ElementID = '" & schemaToCopy.ElementGUID & "') " & _
								" where ElementType = 'SC_MessageProfile' " & _
								" and  ElementID = '" & selectedSchema.ElementGUID & "' " 
					Repository.Execute sqlupdate
					'refresh element
					Repository.AdviseElementChange selectedSchema.ElementID
				end if
			end if
		end if
	end if
end sub

main