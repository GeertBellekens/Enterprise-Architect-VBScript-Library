'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Define tagged values	
' Author: Matthias Van der Elst	
' Purpose: Add tagged values to elements in specific packages	
' Date: 13/03/2017
'
Sub main
	Repository.ClearOutput "Script" 
	Dim element as EA.Element
	Dim i
	Dim WorkPackage as EA.Package
	Dim item as EA.Element
	'MessageType
		'-Message
		'-FIS
		'-IFIS
	'ApplicationLayer
		'-Service Layer
		'-Back-End Layer
	'ServiceType
		'-Mapper
		'-Mecoms Mapper
		'-IA Mapper
		
		
	'MessageType - Message
	
	'MessageType - FIS

	'MessageType - IFIS
	Set WorkPackage = Repository.GetPackageByGuid("{6373C565-E7B3-424b-AD55-486D2C39001A}")
	Session.Output WorkPackage.Name
		For Each element in WorkPackage.Elements
			If element.Stereotype = "Message" then
				Session.Output element.Name
				addTaggedValues element, "MessageType", "IFIS"
			end if
		Next 


	'ServiceType - Mapper
	Dim MapperGUIDs(6)
	MapperGUIDs(0) = "{14BFBA6C-8039-4773-88A5-0E3FF8857C79}" 'UMIG Mappers
	MapperGUIDs(1) = "{CA259A68-F695-4f02-9CA3-DE20386EE10A}" 'UMIG DGO Mappers
	MapperGUIDs(2) = "{F5D20C99-7A1E-40bc-8273-28077127D293}" 'UMIG TSO Mapper
	MapperGUIDs(3) = "{312511B7-4D92-4fe7-942D-F1885B1000BF}" 'UMIG TPDA Mappers
	MapperGUIDs(4) = "{4759CC01-F512-4758-AD8F-C38021D93E0C}" 'UMIG AO Mappers
	MapperGUIDs(5) = "{C6FA78B3-3B5D-4cf0-ABA4-784911FB7986}" 'UMIG PPP Mappers
	MapperGUIDs(6) = "{7CFCBCCC-0EAE-4d4f-8455-484D7EAC07F0}" 'UMIG ATOS Mappers
	
	
	For i = 0 To 6
		Set WorkPackage = Repository.GetPackageByGuid(MapperGUIDs(i))
		Session.Output WorkPackage.Name
		For Each element in WorkPackage.Elements
			If element.Stereotype = "ArchiMate_ApplicationService" then
				Session.Output element.Name
				addTaggedValues element, "ServiceType", "Mapper"
			end if
		Next 
	Next
	
	'ServiceType - Mecoms Mapper
	Set WorkPackage = Repository.GetPackageByGuid("{79D28497-184A-489a-AF7B-CE46AEDE732D}")
	Session.Output WorkPackage.Name
		For Each element in WorkPackage.Elements
			If element.Stereotype = "ArchiMate_ApplicationService" then
				Session.Output element.Name
				addTaggedValues element, "ServiceType", "Mecoms Mapper"
			end if
		Next 
	
	'ServiceType - IA Mapper
	Set WorkPackage = Repository.GetPackageByGuid("{C7F76287-8E1A-43e4-94A3-674C2FAC1F9A}")
	Session.Output WorkPackage.Name
		For Each element in WorkPackage.Elements
			If element.Stereotype = "ArchiMate_ApplicationService" then
				Session.Output element.Name
				addTaggedValues element, "ServiceType", "IA Mapper"
			end if
		Next 
		
	'ApplicationLayer - ServiceLayer
	Set WorkPackage = Repository.GetPackageByGuid("{B65F3118-8241-493a-8743-B72080EDBFE1}")
	Session.Output WorkPackage.Name
		For Each element in WorkPackage.Elements
			If element.Stereotype = "ArchiMate_ApplicationInterface" then
				Session.Output element.Name
				addTaggedValues element, "ApplicationLayer", "Service Layer"
			end if
		Next 

	'ApplicationLayer - Back-End Layer
	Set WorkPackage = Repository.GetPackageByGuid("{C1B43614-1D7D-4c44-ADEB-104FC89DA65A}")
	Session.Output WorkPackage.Name
		For Each element in WorkPackage.Elements
			If element.Stereotype = "ArchiMate_ApplicationInterface" then
				Session.Output element.Name
				addTaggedValues element, "ApplicationLayer", "Back-End Layer"
			end if
		Next 
	

	
End sub

   'addTaggedValues element, "ServiceType", "Mapper"
function addTaggedValues (item, name, value)
	dim TVExist
	dim tv
	TVExist = false

	'first check if it exists
	for each tv in item.TaggedValues
		if tv.Name = name then
			TVExist = true
		end if
	next
	'if not create the tagged values
	if not TVExist then
		set tv = item.TaggedValues.AddNew(name,"")
		tv.Value = value
		tv.Update
		item.Update
	else
		set tv = item.TaggedValues.GetByName(name)
		tv.Value = value
		tv.Update
		item.Update
	end if
	
end function



main