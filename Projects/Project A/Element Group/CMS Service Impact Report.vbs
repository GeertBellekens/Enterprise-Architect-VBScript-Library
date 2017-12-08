'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.DocGenUtil

' Script Name: CMS Service Impact Report
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
					"and op.value = 'Mapper' "
	
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
					"and op.value = 'Service Layer' "
	
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
					"and op.value = 'Mecoms Mapper' "

		
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
					"and op.value = 'Mecoms Mapper' "
		
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
					"and op.value = 'Mecoms Mapper' "
	
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
						"and op.value = 'Service Layer' "
		
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
					"and op.value = 'IA Mapper' "
		
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
						"and op.value = 'Back-End Layer' "

		set mecint = getElementsFromQuery(sqlGetImpact)
		
		for each element in mecint
		mecints.add(element)
		next
	next
	
	'CLEANING ARRAYLISTS
	

	
	
	'1. Selected FIS
	Session.Output "1. Selected FIS: " & selectedElement.Name
	Session.Output "--------------------------------------------------------"
	
	'2. Related Mapper (Service Layer)
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
	
	'7.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable
	Session.Output "7.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable:"
	set backaasout = removeDuplicates(backaasout)
	for each element in backaasout
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable
	Session.Output "7.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable:"
	set backaasin = removeDuplicates(backaasin)
	for each element in backaasin
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper)
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
	
	'11. Impacted Interfaces - MECOMS (Back-End Layer)
	Session.Output "11. Impacted Interfaces - MECOMS (Back-End Layer):"
	set mecints = removeDuplicates(mecints)
	for each element in mecints
		Session.Output element.Name
	next
	
	'show the results in the search window
	showResultInSearchWindow selectedElement,fisaas, aasai, exuccs, aasbp, backint, backaasout, backaasin, aasaas, ifisses, mecasses, mecuccs, mecints
	
	dim result
	result = MsgBox ("Export results to document?", vbYesNo, "Impact Analysis")

	Select Case result
		Case vbYes
			generateDoc selectedElement, fisaas, aasai, exuccs, aasbp, backint, backaasout, backaasin, aasaas, ifisses, mecasses, mecuccs, mecints
		Case vbNo
			MsgBox("You chose No")
	End Select
	
	
	
end sub

function showResultInSearchWindow(selectedElement,fisaas, aasai, exuccs, aasbp, backint, backaasout, backaasin, aasaas, ifisses, mecasses, mecuccs, mecints)
	'get the headers for the output
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.Add "CLASSGUID"
	headers.Add "CLASSTYPE"
	headers.Add "Name"
	headers.Add "Element Type"
	headers.Add "Stereotype"
	headers.Add "Level"
	headers.Add "Package_level1 "
	headers.Add "Package_level2"
	headers.Add "Package_level3"
	'create the output object
	dim searchOutput
	set searchOutput = new SearchResults
	searchOutput.Name = "CMS Service Impact"
	searchOutput.Fields = headers
	'put the contents in the output
	'1. Selected FIS
	dim FISList
	set FISList = CreateObject("System.Collections.ArrayList")
	FISList.Add selectedElement
	putContentInOutput searchOutput, FISList, "01. Selected FIS"
	'2. Related Mapper (Service Layer)
	putContentInOutput searchOutput, fisaas, "02. Related Mapper (Service Layer)"
	'3. Related Interfaces (Service Layer)
	putContentInOutput searchOutput, aasai, "03. Related Interfaces (Service Layer)"
	'4. Impacted Use Cases (Service Layer)
	putContentInOutput searchOutput, exuccs, "04. Impacted Use Cases (Service Layer)"
	'5. Impacted Executables (Service Layer)
	putContentInOutput searchOutput, aasbp, "05. Impacted Executables (Service Layer)"
	'6. Impacted Interfaces (Service Layer)
	putContentInOutput searchOutput, backint, "06. Impacted Interfaces (Service Layer)"
	'7.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable
	putContentInOutput searchOutput, backaasout, "07.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable"
	'7.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable
	putContentInOutput searchOutput, backaasin, "07.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable"
	'7.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper)
	putContentInOutput searchOutput, aasaas, "07.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper)"
	'8. Impacted I/FIS (Service Layer)
	putContentInOutput searchOutput, ifisses, "08. Impacted I/FIS (Service Layer)"
	'9. Impacted Interface Agreement (IA) Application Services (Back-End Layer)
	putContentInOutput searchOutput, mecasses, "09. Impacted Interface Agreement (IA) Application Services (Back-End Layer)"
	'10. Impacted Use Cases - MECOMS (Back-End Layer)
	putContentInOutput searchOutput, mecuccs, "10. Impacted Use Cases - MECOMS (Back-End Layer)"
	'11. Impacted Interfaces - MECOMS (Back-End Layer)
	putContentInOutput searchOutput, mecints, "11. Impacted Interfaces - MECOMS (Back-End Layer))"
	'show the output
	searchOutput.Show
end function

function putContentInOutput(searchOutput, elementList, Level)
	dim currentElement as EA.Element
	for each currentElement in elementList
		dim currentRow 
		set currentRow = CreateObject("System.Collections.ArrayList")
		currentRow.Add currentElement.ElementGUID
		currentRow.Add currentElement.Type
		currentRow.Add currentElement.Name
		currentRow.Add currentElement.Type
		currentRow.Add currentElement.Stereotype
		currentRow.Add level
		'get the parent packages
		'packageL1
		dim packageL1 as EA.Package
		set packageL1 = Repository.GetPackageByID(currentElement.PackageID)
		currentRow.Add packageL1.Name
		'packageL2
		dim packageL2 as EA.Package
		if packageL1.ParentID > 0 then
			set packageL2 = Repository.GetPackageByID(packageL1.ParentID)
			currentRow.Add packageL2.Name
			'packageL3
			dim packageL3 as EA.Package
			if packageL2.ParentID > 0 then
				set packageL3 = Repository.GetPackageByID(packageL2.ParentID)
				currentRow.Add packageL3.Name
			else
				currentRow.Add "" 'Package_level3
			end if
		else
			currentRow.Add "" 'Package_level2
			currentRow.Add "" 'Package_level3
		end if
		'add the row to the results
		searchOutput.Results.Add currentRow
	next
end function

sub generateDoc(selectedElement, fisaas, aasai, exuccs, aasbp, backint, backaasout, backaasin, aasaas, ifisses, mecasses, mecuccs, mecints)
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
	packageGUID = "{21E715ED-25B2-4255-AF61-1EEAA8EE6305}"
	documentName = "Impact for " & selectedElement.Name
	set masterDocument = addMasterdocument(packageGUID, documentName) 'without further details

	'1. Selected FIS
	addModelDocument masterDocument, "IMP_Element", selectedElement.Name, selectedElement.ElementGUID, i
	i = i + 1
	
	'2. Related Mapper (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE2", "Related Mapper (Service Layer)", "", i, ""
	i = i + 1
	if fisaas.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in fisaas
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'3. Related Interfaces (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE3", "Related Interfaces (Service Layer)", "", i, ""
	i = i + 1
	if aasai.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in aasai
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'4. Impacted Use Cases (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE4", "Impacted Use Cases (Service Layer)", "", i, ""
	i = i + 1
	if exuccs.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in exuccs
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'5. Impacted Executables (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE5", "Impacted Executables (Service Layer)", "", i, ""
	i = i + 1
	if aasbp.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in aasbp
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'6. Impacted Interfaces (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE6", "Impacted Interfaces (Service Layer)", "", i, ""
	i = i + 1
	if backint.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in backint
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'7. Impacted MECOMS Mappers (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE7", "Impacted MECOMS Mappers (Service Layer)", "", i, ""
	i = i + 1

	'7.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable
	addModelDocumentWithSearch masterDocument, "IMP_TITLE71", "Impacted MECOMS Mappers (Service Layer) - Transition from executable", "", i, ""
	i = i + 1
	if backaasout.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in backaasout
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	
	'7.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable
	addModelDocumentWithSearch masterDocument, "IMP_TITLE72", "Impacted MECOMS Mappers (Service Layer) - Transition to executable", "", i, ""
	i = i + 1
	if backaasin.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in backaasin
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'7.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE73", "Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper)", "", i, ""
	i = i + 1
	if aasaas.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in aasaas
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'8. Impacted I/FIS (Service Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE8", "Impacted I/FIS (Service Layer)", "", i, ""
	i = i + 1
	if ifisses.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in ifisses
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'9. Impacted Interface Agreement (IA) Application Services (Back-End Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE9", "Impacted Interface Agreement (IA) Application Services (Back-End Layer)", "", i, ""
	i = i + 1
	if mecasses.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in mecasses
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	'10. Impacted Use Cases - MECOMS (Back-End Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE10", "Impacted Use Cases - MECOMS (Back-End Layer)", "", i, ""
	i = i + 1
	if mecuccs.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in mecuccs
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next

	'11. Impacted Interfaces - MECOMS (Back-End Layer)
	addModelDocumentWithSearch masterDocument, "IMP_TITLE11", "Impacted Interfaces - MECOMS (Back-End Layer)", "", i, ""
	i = i + 1
	if mecints.count() = 0 then
		addModelDocumentWithSearch masterDocument, "IMP_EMPTY", "", "", i, ""
		i = i + 1
	end if
	for each element in mecints
		addModelDocument masterDocument, "IMP_EL", element.Name, element.ElementGUID, i
		i = i + 1
	next
	
	
	'reload the package to show the correct order
	Repository.RefreshModelView(masterDocument.PackageID)
	
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