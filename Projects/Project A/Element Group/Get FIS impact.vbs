'[path=\Projects\Project A\Project Browser Element Group]
'[group=Project Browser Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.DocGenUtil

' Script Name: Get FIS impact
' Author: Matthias Van der Elst
' Purpose: Get impact for the selected FIS
' Date: 08/03/2017
'

'
' Project Browser Script main function
sub OnProjectBrowserElementScript()
	Repository.ClearOutput "Script"

	' Get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject()
	dim eStereoType 'Element stereotype
	eStereoType = selectedElement.Stereotype
	 if eStereoType = "Message" then
		getImpact(selectedElement)
	 else
		Session.Prompt "The selected element is not of the correct type", promptOK
	 end if
	
end sub

sub getImpact(selectedElement)
	dim sqlGetImpact 'Query
	dim eID 'Element ID (= Object_ID)
	dim element as EA.Element 'For looping the arraylist
	dim aasID 'Id of the impacted internal application service
	dim aiID 'Id of the impacted internal application interface
	dim adID 'Id of the impacted application component adapter (AtriasMIG6MessagePipeline)
	eID = selectedElement.ElementID
	
	' Get the impacted Application Services
	dim fisaas 
	set fisaas = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact =  "select aas.object_id " & _
					"from (((t_object fis " & _
					"inner join t_connector con " & _
					"on fis.object_id = con.start_object_id or fis.object_id = con.end_object_id) " & _
					"inner join t_object aas " & _
					"on aas.object_id = con.start_object_id or aas.object_id = con.end_object_id) " & _
					"inner join t_objectproperties op " & _
					"on aas.object_id = op.object_id) " & _
					"where fis.Object_ID = '" & eID & "' " & _
					"and aas.object_id <> fis.Object_ID " & _
					"and aas.stereotype = 'ArchiMate_ApplicationService' " & _
					"and op.property = 'ServiceType' " & _
					"and op.value = 'Mapper Service' "
	
	set fisaas = getElementsFromQuery(sqlGetImpact)
	for each element in fisaas
		aasID = element.ElementID
	next
	
	' Get the impacted Use Cases 
	' Directly by FIS
	dim exuccs
	set exuccs = CreateObject("System.Collections.ArrayList")
	dim fisuc
	set fisuc = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact =  "select uc.object_id " & _
					"from ((t_object fis " & _
					"inner join t_connector con " & _
					"on fis.object_id = con.start_object_id or fis.object_id = con.end_object_id) " & _
					"inner join t_object uc " & _
					"on uc.object_id = con.start_object_id or uc.object_id = con.end_object_id) " & _
					"where fis.Object_ID = '" & eID & "' " & _
					"and uc.object_id <> fis.Object_ID " & _
					"and uc.object_type = 'UseCase' "
					
	set fisuc = getElementsFromQuery(sqlGetImpact)
	for each element in fisuc
		exuccs.add(element)
	next
	
	' By Application Service
	dim aasuc
	set aasuc = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact =  "select uc.object_id " & _
					"from ((t_object aas " & _
					"inner join t_connector con " & _
					"on aas.object_id = con.start_object_id or aas.object_id = con.end_object_id) " & _
					"inner join t_object uc " & _
					"on uc.object_id = con.start_object_id or uc.object_id = con.end_object_id) " & _
					"where aas.Object_ID = '" & aasID & "' " & _
					"and uc.object_id <> aas.Object_ID " & _
					"and uc.object_type = 'UseCase' "
	
	
					
	set aasuc = getElementsFromQuery(sqlGetImpact)
	for each element in aasuc
		exuccs.add(element)
	next
	
	' Get the impacted executables   stereotype = 'BusinessProcess
	dim aasbp
	set aasbp = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact =  "select bp.object_id " & _
					"from ((t_object aas " & _
					"inner join t_connector con " & _
					"on aas.object_id = con.start_object_id or aas.object_id = con.end_object_id) " & _
					"inner join t_object bp " & _
					"on bp.object_id = con.start_object_id or bp.object_id = con.end_object_id) " & _
					"where aas.Object_ID = '" & aasID & "' " & _
					"and bp.object_id <> aas.Object_ID " & _
					"and bp.stereotype = 'BusinessProcess' "
			
	set aasbp = getElementsFromQuery(sqlGetImpact)
	' Save here incoming or outgoing
	'dim connectors
	'set connectors = CreateObject("System.Collections.ArrayList")
	'for each element  in aasbp
		
	'next
	

	' Get the impacted interfaces   stereotype = 'archimate_applicationinterface'
	dim aasai
	set aasai = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact =  "select ai.object_id " & _
					"from (((t_object aas " & _
					"inner join t_connector con " & _
					"on aas.object_id = con.start_object_id or aas.object_id = con.end_object_id) " & _
					"inner join t_object ai " & _
					"on ai.object_id = con.start_object_id or ai.object_id = con.end_object_id) " & _
					"inner join t_objectproperties op " & _
					"on ai.object_id = op.object_id) " & _
					"where aas.Object_ID = '" & aasID & "' " & _
					"and ai.object_id <> aas.Object_ID " & _
					"and ai.stereotype = 'archimate_applicationinterface' " & _
					"and op.property = 'ApplicationLayer' " & _
					"and op.value = 'CMS Service Layer' "
	
	set aasai = getElementsFromQuery(sqlGetImpact)
	for each element in aasai
		aiID = element.ElementID
	next
	
	' Get the impacted application component (AtriasMIG6MessagePipeline)
	dim aiad
	set aiad = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact =  "select ad.object_id " & _
					"from ((t_object ai " & _
					"inner join t_connector con " & _
					"on ai.object_id = con.start_object_id or ai.object_id = con.end_object_id) " & _
					"inner join t_object ad " & _
					"on ad.object_id = con.start_object_id or ad.object_id = con.end_object_id) " & _
					"where ai.Object_ID = '" & aiID & "' " & _
					"and ad.object_id <> ai.Object_ID " & _
					"and ad.stereotype = 'archimate_applicationcomponent' "
	
	set aiad = getElementsFromQuery(sqlGetImpact)
	for each element in aiad
		adID = element.ElementID
	next
	
	' Get the impacted application component (webMethods Enterprise Service Bus)
	dim adesb
	set adesb = CreateObject("System.Collections.ArrayList")
	
	sqlGetImpact =  "select esb.object_id " & _
					"from ((t_object ad " & _
					"inner join t_connector con " & _
					"on ad.object_id = con.start_object_id or ad.object_id = con.end_object_id) " & _
					"inner join t_object esb " & _
					"on esb.object_id = con.start_object_id or esb.object_id = con.end_object_id) " & _
					"where ad.Object_ID = '" & adID & "' " & _
					"and esb.object_id <> ad.Object_ID " & _
					"and esb.stereotype = 'archimate_applicationcomponent' "
	
	set adesb = getElementsFromQuery(sqlGetImpact)
	
	' Get the impacted services at the internal side, directly aas to aas, or via an executable, Mecoms Mappers
	dim backaas
	set backaas = CreateObject("System.Collections.ArrayList")
	' If incoming, only outgoing services
	' If outgoing, only incoming services
	' 1) via the executable, there can be multiple executables (aasbp)
	dim backaasout
	set backaasout = CreateObject("System.Collections.ArrayList")
	dim backaasin
	set backaasin = CreateObject("System.Collections.ArrayList")
	dim aas as EA.Element
	' Found executables
	' aasID = the id of the Mapper Service
	for each element in aasbp 
		dim bpaasout 'Outgoing
		set bpaasout = CreateObject("System.Collections.ArrayList")
		sqlGetImpact =  "select aas.object_id " & _
					"from (((t_object bp " & _
					"inner join t_connector con " & _
					"on bp.object_id = con.start_object_id) " & _
					"inner join t_object aas " & _
					"on aas.object_id = con.end_object_id) " & _
					"inner join t_objectproperties op " & _
					"on aas.object_id = op.object_id) " & _
					"where bp.Object_ID = '" & element.ElementID & "' " & _
					"and aas.object_id <> '" & aasID & "' " & _
					"and aas.stereotype = 'ArchiMate_ApplicationService' " & _
					"and op.property = 'ServiceType' " & _
					"and op.value = 'Mecoms Mapper Service' "

		
		set bpaasout = getElementsFromQuery(sqlGetImpact)
		for each aas in bpaasout
		backaasout.add(aas)
		backaas.add(aas)
		next
		
		dim bpaasin 'Incoming
		set bpaasin = CreateObject("System.Collections.ArrayList")
		sqlGetImpact =  "select aas.object_id " & _
					"from (((t_object bp " & _
					"inner join t_connector con " & _
					"on bp.object_id = con.end_object_id) " & _
					"inner join t_object aas " & _
					"on aas.object_id = con.start_object_id) " & _
					"inner join t_objectproperties op " & _
					"on aas.object_id = op.object_id) " & _
					"where bp.Object_ID = '" & element.ElementID & "' " & _
					"and aas.object_id <> '" & aasID & "' " & _
					"and aas.stereotype = 'ArchiMate_ApplicationService' " & _
					"and op.property = 'ServiceType' " & _
					"and op.value = 'Mecoms Mapper Service' "
		
		set bpaasin = getElementsFromQuery(sqlGetImpact)
		for each aas in bpaasin
		backaasin.add(aas)
		backaas.add(aas)
		next
		
	next				
		

	' 2) directly from external aas to internal aas
	dim aasaas 'Direct
	set aasaas = CreateObject("System.Collections.ArrayList")
	sqlGetImpact =  "select aas.object_id " & _
					"from (((t_object exaas " & _
					"inner join t_connector con " & _
					"on exaas.object_id = con.start_object_id or exaas.object_id = con.end_object_id) " & _
					"inner join t_object aas " & _
					"on aas.object_id = con.start_object_id or aas.object_id = con.end_object_id) " & _
					"inner join t_objectproperties op " & _
					"on aas.object_id = op.object_id) " & _
					"where exaas.Object_ID = '" & aasID & "' " & _
					"and aas.object_id <> exaas.Object_ID " & _
					"and aas.stereotype = 'ArchiMate_ApplicationService' " & _
					"and op.property = 'ServiceType' " & _
					"and op.value = 'Mecoms Mapper Service' "
	
	set aasaas = getElementsFromQuery(sqlGetImpact)
		for each aas in aasaas
		backaas.add(aas)
		next	

	' Get the internal impacted interfaces
	dim backint
	set backint = CreateObject("System.Collections.ArrayList")
	dim intaas as EA.Element
	' Found internal services
	for each intaas in backaas
		dim aasint
		set aasint = CreateObject("System.Collections.ArrayList")
		sqlGetImpact =  "select intai.object_id " & _
						"from (((t_object aas " & _
						"inner join t_connector con " & _
						"on aas.object_id = con.start_object_id or aas.object_id = con.end_object_id) " & _
						"inner join t_object intai " & _
						"on intai.object_id = con.start_object_id or intai.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on intai.object_id = op.object_id) " & _
						"where aas.Object_ID = '" & intaas.ElementID & "' " & _
						"and intai.stereotype = 'archimate_applicationinterface' " & _
						"and op.property = 'ApplicationLayer' " & _
						"and op.value = 'CMS Service Layer' "
		
		set aasint = getElementsFromQuery(sqlGetImpact)
		for each aas in aasint
		backint.add(aas)
		next	
	next
	
	
	' Get the impacted IFIS
	dim ifisses
	set ifisses = CreateObject("System.Collections.ArrayList")

	'Found internal services
	for each intaas in backaas
		dim intfis
		set intfis = CreateObject("System.Collections.ArrayList")
		sqlGetImpact =  "select ifis.object_id " & _
						"from (((t_object aas " & _
						"inner join t_connector con " & _
						"on aas.object_id = con.start_object_id or aas.object_id = con.end_object_id) " & _
						"inner join t_object ifis " & _
						"on ifis.object_id = con.start_object_id or ifis.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on ifis.object_id = op.object_id) " & _
						"where aas.Object_ID = '" & intaas.ElementID & "' " & _
						"and ifis.stereotype = 'Message' " & _
						"and op.property = 'MessageType' " & _
						"and op.value = 'IFIS' "
		
		set intfis = getElementsFromQuery(sqlGetImpact)
		for each element in intfis
		element.Tag = intaas.ElementID 'TAG = ID from the internal service to make to exclude them from the services found in next step
		ifisses.add(element)
		next
	next
	
	' Get the impacted IA services
	dim ifis as EA.Element
	dim mecasses
	set mecasses = CreateObject("System.Collections.ArrayList")
	
	for each ifis in ifisses
		dim mecas
		set mecas = CreateObject("System.Collections.ArrayList")
		sqlGetImpact =  "select mecas.object_id " & _
					"from (((t_object ifis " & _
					"inner join t_connector con " & _
					"on ifis.object_id = con.start_object_id or ifis.object_id = con.end_object_id) " & _
					"inner join t_object mecas " & _
					"on mecas.object_id = con.start_object_id or mecas.object_id = con.end_object_id) " & _
					"inner join t_objectproperties op " & _
					"on mecas.object_id = op.object_id) " & _
					"where ifis.Object_ID = '" & ifis.ElementID & "' " & _
					"and mecas.object_id <> '" & ifis.Tag & "' " & _ 
					"and mecas.stereotype = 'ArchiMate_ApplicationService' " & _ 
					"and op.property = 'ServiceType' " & _
					"and op.value = 'IA Service' "
		
		set mecas = getElementsFromQuery(sqlGetImpact)
		for each element in mecas
		mecasses.add(element)
		next
	next
	
	' Get the impacted MECOMS Use Cases
	dim mecaas as EA.Element
	dim mecuccs
	set mecuccs = CreateObject("System.Collections.ArrayList")

	for each mecaas in mecasses
		dim mecuc	
		set mecuc = CreateObject("System.Collections.ArrayList")
		
		sqlGetImpact =  "select mecuc.object_id " & _
						"from ((t_object mecas " & _
						"inner join t_connector con " & _
						"on mecas.object_id = con.start_object_id or mecas.object_id = con.end_object_id) " & _
						"inner join t_object mecuc " & _
						"on mecuc.object_id = con.start_object_id or mecuc.object_id = con.end_object_id) " & _
						"where mecas.Object_ID = '" & mecaas.ElementID & "' " & _
						"and mecuc.object_type = 'UseCase' "
		set mecuc = getElementsFromQuery(sqlGetImpact)
		
		for each element in mecuc
		mecuccs.add(element)
		next
	next
	
	' Get the impacted MECOMS Interfaces
	dim mecints
	set mecints = CreateObject("System.Collections.ArrayList")
	for each mecaas in mecasses
		dim mecint	
		set mecint = CreateObject("System.Collections.ArrayList")
		
		sqlGetImpact =  "select mecint.object_id " & _
						"from (((t_object mecas " & _
						"inner join t_connector con " & _
						"on mecas.object_id = con.start_object_id or mecas.object_id = con.end_object_id) " & _
						"inner join t_object mecint " & _
						"on mecint.object_id = con.start_object_id or mecint.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on mecint.object_id = op.object_id) " & _
						"where mecas.Object_ID = '" & mecaas.ElementID & "' " & _
						"and mecint.stereotype = 'archimate_applicationinterface' " & _
						"and op.property = 'ApplicationLayer' " & _
						"and op.value = 'CMS Back-End Layer' "

		set mecint = getElementsFromQuery(sqlGetImpact)
		
		for each element in mecint
		mecints.add(element)
		next
	next
	
	'CLEANING ARRAYLISTS
	
	
	'1. Selected FIS (external)
	Session.Output "1. Selected FIS: " & selectedElement.Name
	Session.Output "--------------------------------------------------------"
	
	'2. Related Application Services – Mapper (Service layer)
	Session.Output "2. Related Mapper (Service Layer):"
	set fisaas = removeDuplicates(fisaas)
	for each element in fisaas
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'3. Related Interfaces (Service Layer)
	Session.Output "3. Related Interfaces (Service Layer):"
	set aasai = removeDuplicates(aasai)
	for each element in aasai
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'4. Impacted Use Cases (Service Layer)
	Session.Output "4. Impacted Use Cases (Service Layer)"
	set exuccs = removeDuplicates(exuccs)
	for each element in exuccs
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'5. Impacted Executables (Service Layer)
	Session.Output "5. Impacted Executables (Service Layer):"
	set aasbp = removeDuplicates(aasbp)
	for each element in aasbp
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'6. Impacted Interfaces (Service Layer)
	Session.Output "6. Impacted Interfaces (Service Layer):"
	set backint = removeDuplicates(backint)
	for each element in backint
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.1 Impacted Application Services – (Mecom) Mapper (Service Layer) - OUT
	Session.Output "7.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable:"
	set backaasout = removeDuplicates(backaasout)
	for each element in backaasout
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.2 Impacted Application Services – (Mecom) Mapper (Service Layer) - IN
	Session.Output "7.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable:"
	set backaasin = removeDuplicates(backaasin)
	for each element in backaasin
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.3 Impacted Application Services – (Mecom) Mapper (Service Layer) - DIRECT (AS 2 AS)
	Session.Output "7.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper):"
	set aasaas = removeDuplicates(aasaas)
	for each element in aasaas
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'8. Impacted I/FIS (Service Layer)
	Session.Output "8. Impacted I/FIS (Service Layer):"
	set ifisses = removeDuplicates(ifisses)
	for each element in ifisses
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'9. Impacted Interface Agreement (IA) Application Services (Back-End Layer)
	Session.Output "9. Impacted Interface Agreement (IA) Application Services (Back-End Layer):"
	set mecasses = removeDuplicates(mecasses)
	for each element in mecasses
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'10. Impacted Use Cases - MECOMS (Back-End Layer)
	Session.Output "10. Impacted Use Cases - MECOMS (Back-End Layer):"
	set mecuccs = removeDuplicates(mecuccs)
	for each element in mecuccs
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'11. Impacted interface - MECOMS (Back- End Layer)
	Session.Output "11. Impacted interface - MECOMS (Back-End Layer):"
	set mecints = removeDuplicates(mecints)
	for each element in mecints
		Session.Output element.Name
	next
	
	
	dim result
	result = MsgBox ("Export results to document?", vbYesNo, "Impact Analysis")

	Select Case result
		Case vbYes
			'generateDoc selectedElement, fisaas, aasai, exuccs, aasbp, backint, backaas, ifisses, mecasses, mecuccs, mecints
		Case vbNo
			MsgBox("You chose No")
	End Select
	
	
	
end sub

sub generateDoc(selectedElement, fisaas, aasai, exuccs, aasbp, backint, backaas, ifisses, mecasses, mecuccs, mecints)
	'CREATION OF THE DOCUMENT
	'Add master document
	dim packageGUID
	dim documentName
	dim masterDocument as EA.Package
	dim element as EA.Element
	dim i
	i = 1
	'dev: "{21E715ED-25B2-4255-AF61-1EEAA8EE6305}"
	'productie: "{09D83051-0F18-4501-BAC0-611002EA928F}"
	packageGUID = "{09D83051-0F18-4501-BAC0-611002EA928F}"
	documentName = "Impact for " & selectedElement.Name
	'eTemplate = "IMP_Element" 'Is dit nodig?
	set masterDocument = addMasterdocument(packageGUID, documentName) 'without further details

	'1. Selected FIS
	addModelDocument masterDocument, "IMP_Element", selectedElement.Name, selectedElement.ElementGUID, i
	i = i + 1
	
	'2. Related Application Services – Mapper (Service layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE1", "Impacted Application Services (external)", "", i, "Impacted Application Services (external)"
	i = i + 1
	for each element in fisaas
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'3. Related Interfaces (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE4", "Impacted Interfaces (external):", "", i, ""
	i = i + 1
	for each element in aasai
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'4. Impacted Use Cases (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE2", "Impacted Use Cases (external):", "", i, ""
	i = i + 1
	for each element in exuccs
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'5. Impacted Executables (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE3", "Impacted Executables (external):", "", i, ""
	i = i + 1
	for each element in aasbp
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'6. Impacted Interfaces (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE5", "Impacted Interfaces (internal):", "", i, ""
	i = i + 1
	for each element in backint
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next

	'7. Impacted Application Services – (Mecom) Mapper (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE6", "Impacted Application Services (internal):", "", i, ""
	i = i + 1
	for each element in backaas
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'8. Impacted I/FIS (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE7", "Impacted IFIS (internal):", "", i, ""
	i = i + 1
	for each element in ifisses
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'9. Impacted Interface Agreement (IA) Application Services (Back-End Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE8", "Impacted MECOMS Services (internal):", "", i, ""
	i = i + 1
	for each element in mecasses
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'10. Impacted Use Cases - MECOMS (Back-End Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE9", "Impacted MECOMS Use Cases (internal):", "", i, ""
	i = i + 1
	for each element in mecuccs
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next

	'11. Impacted interface - MECOMS (Back- End Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE10", "Impacted MECOMS Interfaces (internal):", "", i, ""
	i = i + 1
	for each element in mecints
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	
	'reload the package to show the correct order
	'Repository.RefreshModelView(masterDocument.PackageID)
	
	MsgBox "Documentation is generated", vbOKOnly, "Impact Analysis"
	
	
	
end sub


'returns an ArrayList without duplicates
function removeDuplicates(arraylist)
	dim result
	set result = CreateObject("System.Collections.ArrayList")
	dim element as EA.Element
	for each element in arrayList
		dim b
		b = contains(result, element)
		if b = false then
			result.add(element)
		end if	
	next
	set removeDuplicates = result
end function


'returns boolean 
function contains(result, element)
	contains = false
	dim res as EA.Element
	for each res in result
		if res.ElementID = element.ElementID then
			contains = true
			exit for
		end if	
	next
end function


OnProjectBrowserElementScript