'[path=\Projects\Project A\Element Group]
'[group=Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.DocGenUtil

' Script Name: Get Use Case impact
' Author: Matthias Van der Elst
' Purpose: Get impact for the selected Use Case
' Date: 16/03/2017
'

'
' Project Browser Script main function
sub OnProjectBrowserElementScript()
	Repository.ClearOutput "Script"
	' Get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject()
	dim eType 'Element type
	eType = selectedElement.Type
	
	 if eType = "UseCase" then
		getImpact(selectedElement)
	 else
		Session.Prompt "The selected element is not of the correct type", promptOK
	 end if
	
end sub


sub getImpact(selectedElement)
	dim sqlGetImpact 'Query
	dim eID 'Element ID (= Object_ID)
	dim element as EA.Element 'For looping the arraylist
	eID = selectedElement.ElementID
	
	'2. Related FIS (Service Layer)
	dim fisses
	set fisses = CreateObject("System.Collections.ArrayList")
	sqlGetImpact =  "select fis.object_id " & _
					"from ((t_object uc " & _
					"inner join t_connector con " & _
					"on uc.object_id = con.start_object_id or uc.object_id = con.end_object_id) " & _
					"inner join t_object fis " & _
					"on fis.object_id = con.start_object_id or fis.object_id = con.end_object_id) " & _
					"where uc.Object_ID = '" & eID & "' " & _
					"and fis.object_id <> uc.Object_ID " & _
					"and fis.stereotype = 'Message' " 
	set fisses = getElementsFromQuery(sqlGetImpact)
	set fisses = removeDuplicates(fisses)
	
	'3. Related Mapper (Service Layer)
	dim msses
	set msses = CreateObject("System.Collections.ArrayList")
	sqlGetImpact =  "select ms.object_id " & _
					"from (((t_object uc " & _
					"inner join t_connector con " & _
					"on uc.object_id = con.start_object_id or uc.object_id = con.end_object_id) " & _
					"inner join t_object ms " & _
					"on ms.object_id = con.start_object_id or ms.object_id = con.end_object_id) " & _
					"inner join t_objectproperties op " & _
					"on ms.object_id = op.object_id) " & _
					"where uc.Object_ID = '" & eID & "' " & _
					"and ms.object_id <> uc.Object_ID " & _
					"and ms.stereotype = 'archimate_applicationservice' " & _
					"and op.property = 'ServiceType' " & _
					"and op.value = 'Mapper Service' "
	
	set msses = getElementsFromQuery(sqlGetImpact)
	set msses = removeDuplicates(msses)

	'4. Related Interfaces (Service Layer)
	dim ris
	set ris = CreateObject("System.Collections.ArrayList")
	for each element in msses
		sqlGetImpact =  "select ri.object_id " & _
						"from (((t_object ms " & _
						"inner join t_connector con " & _
						"on ms.object_id = con.start_object_id or ms.object_id = con.end_object_id) " & _
						"inner join t_object ri " & _
						"on ri.object_id = con.start_object_id or ri.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on ri.object_id = op.object_id) " & _
						"where ms.Object_ID = '" & element.ElementID & "' " & _
						"and ri.object_id <> ms.Object_ID " & _
						"and ri.stereotype = 'archimate_applicationinterface' " & _
						"and op.property = 'ApplicationLayer' " & _
						"and op.value = 'CMS Service Layer' "
		ris.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set ris = removeDuplicates(ris)


	'5. Impacted Executables (Service Layer)
	dim excs
	set excs = CreateObject("System.Collections.ArrayList")
	for each element in msses
		sqlGetImpact =  "select exc.object_id " & _
						"from ((t_object ms " & _
						"inner join t_connector con " & _
						"on ms.object_id = con.start_object_id or ms.object_id = con.end_object_id) " & _
						"inner join t_object exc " & _
						"on exc.object_id = con.start_object_id or exc.object_id = con.end_object_id) " & _
						"where ms.Object_ID = '" & element.ElementID & "' " & _
						"and exc.object_id <> ms.Object_ID " & _
						"and exc.stereotype = 'BusinessProcess' "
		excs.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set excs = removeDuplicates(excs)
	

	'7.1/2 Impacted MECOMS Mappers (Service Layer) - Transition from/to executable
	dim mmso
	dim mmsi
	dim mmsd
	dim mms
	set mmso = CreateObject("System.Collections.ArrayList")
	set mmsi = CreateObject("System.Collections.ArrayList")
	set mmsd = CreateObject("System.Collections.ArrayList")
	set mms = CreateObject("System.Collections.ArrayList")
	
	for each element in excs
		sqlGetImpact =  "select mmo.object_id " & _
						"from (((t_object exc " & _
						"inner join t_connector con " & _
						"on exc.object_id = con.start_object_id) " & _
						"inner join t_object mmo " & _
						"on mmo.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on mmo.object_id = op.object_id) " & _
						"where exc.Object_ID = '" & element.ElementID & "' " & _
						"and mmo.stereotype = 'ArchiMate_ApplicationService' " & _
						"and op.property = 'ServiceType' " & _
						"and op.value = 'Mecoms Mapper Service' "
		mmso.AddRange(getElementsFromQuery(sqlGetImpact))
		
		sqlGetImpact =  "select mmi.object_id " & _
						"from (((t_object exc " & _
						"inner join t_connector con " & _
						"on exc.object_id = con.end_object_id) " & _
						"inner join t_object mmi " & _
						"on mmi.object_id = con.start_object_id) " & _
						"inner join t_objectproperties op " & _
						"on mmi.object_id = op.object_id) " & _
						"where exc.Object_ID = '" & element.ElementID & "' " & _
						"and mmi.stereotype = 'ArchiMate_ApplicationService' " & _
						"and op.property = 'ServiceType' " & _
						"and op.value = 'Mecoms Mapper Service' "
		mmsi.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set mmso = removeDuplicates(mmso)
	set mmsi = removeDuplicates(mmsi)
	mms.AddRange(mmso)
	mms.AddRange(mmsi)

	'7.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper)
	for each element in msses
		sqlGetImpact =  "select mmd.object_id " & _
						"from (((t_object ms " & _
						"inner join t_connector con " & _
						"on ms.object_id = con.start_object_id or ms.object_id = con.end_object_id) " & _
						"inner join t_object mmd " & _
						"on mmd.object_id = con.start_object_id or mmd.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on mmd.object_id = op.object_id) " & _
						"where ms.Object_ID = '" & element.ElementID & "' " & _
						"and mmd.stereotype = 'ArchiMate_ApplicationService' " & _
						"and op.property = 'ServiceType' " & _
						"and op.value = 'Mecoms Mapper Service' "
		mmsd.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set mmsd = removeDuplicates(mmsd)
	mms.AddRange(mmsd)
	
	'6. Impacted Interfaces (Service Layer)
	dim iis
	set iis = CreateObject("System.Collections.ArrayList")
	for each element in mms
		sqlGetImpact =  "select ii.object_id " & _
						"from (((t_object mm " & _
						"inner join t_connector con " & _
						"on mm.object_id = con.start_object_id or mm.object_id = con.end_object_id) " & _
						"inner join t_object ii " & _
						"on ii.object_id = con.start_object_id or ii.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on ii.object_id = op.object_id) " & _
						"where mm.Object_ID = '" & element.ElementID & "' " & _
						"and ii.stereotype = 'archimate_applicationinterface' " & _
						"and op.property = 'ApplicationLayer' " & _
						"and op.value = 'CMS Service Layer' "
		iis.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set iis = removeDuplicates(iis)
	
	'8. Impacted I/FIS (Service Layer)
	dim ifisses
	set ifisses = CreateObject("System.Collections.ArrayList")
	for each element in mms
		sqlGetImpact =  "select ifis.object_id " & _
						"from (((t_object mm " & _
						"inner join t_connector con " & _
						"on mm.object_id = con.start_object_id or mm.object_id = con.end_object_id) " & _
						"inner join t_object ifis " & _
						"on ifis.object_id = con.start_object_id or ifis.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on ifis.object_id = op.object_id) " & _
						"where mm.Object_ID = '" & element.ElementID & "' " & _
						"and ifis.stereotype = 'Message' " & _
						"and op.property = 'MessageType' " & _
						"and op.value = 'IFIS' "
		ifisses.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set ifisses = removeDuplicates(ifisses)

	'9. Impacted Interface Agreement (IA) Application Services (Back-End Layer)
	dim iass
	set iass = CreateObject("System.Collections.ArrayList")
	for each element in ifisses
		sqlGetImpact =  "select ias.object_id " & _
						"from (((t_object ifis " & _
						"inner join t_connector con " & _
						"on ifis.object_id = con.start_object_id or ifis.object_id = con.end_object_id) " & _
						"inner join t_object ias " & _
						"on ias.object_id = con.start_object_id or ias.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on ias.object_id = op.object_id) " & _
						"where ifis.Object_ID = '" & element.ElementID & "' " & _ 
						"and ias.stereotype = 'ArchiMate_ApplicationService' " & _ 
						"and op.property = 'ServiceType' " & _
						"and op.value = 'IA Service' "
		iass.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set iass = removeDuplicates(iass)

	'10. Impacted Use Cases - MECOMS (Back-End Layer)
	dim iucms
	set iucms = CreateObject("System.Collections.ArrayList")
	for each element in iass
		sqlGetImpact =  "select iucm.object_id " & _
						"from ((t_object ias " & _
						"inner join t_connector con " & _
						"on ias.object_id = con.start_object_id or ias.object_id = con.end_object_id) " & _
						"inner join t_object iucm " & _
						"on iucm.object_id = con.start_object_id or iucm.object_id = con.end_object_id) " & _
						"where ias.Object_ID = '" & element.ElementID & "' " & _
						"and iucm.object_type = 'UseCase' "
		iucms.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set iucms = removeDuplicates(iucms)

	'11. Impacted interface - MECOMS (Back- End Layer)
	dim iims
	set iims = CreateObject("System.Collections.ArrayList")
	for each element in iass
		sqlGetImpact =  "select iim.object_id " & _
						"from (((t_object ias " & _
						"inner join t_connector con " & _
						"on ias.object_id = con.start_object_id or ias.object_id = con.end_object_id) " & _
						"inner join t_object iim " & _
						"on iim.object_id = con.start_object_id or iim.object_id = con.end_object_id) " & _
						"inner join t_objectproperties op " & _
						"on iim.object_id = op.object_id) " & _
						"where ias.Object_ID = '" & element.ElementID & "' " & _
						"and iim.stereotype = 'archimate_applicationinterface' " & _
						"and op.property = 'ApplicationLayer' " & _
						"and op.value = 'CMS Back-End Layer' "
		iims.AddRange(getElementsFromQuery(sqlGetImpact))
	next
	set iims = removeDuplicates(iims)
	
	
	'1. The selected Use Case
	Session.Output "1. Selected Use Case: " & selectedElement.Name
	Session.Output "--------------------------------------------------------"
	
	'2. Related FIS (Service Layer)
	Session.Output "2. Related FIS (Service Layer): "
	for each element in fisses
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'3. Related Mapper (Service Layer)
	Session.Output "3. Related Mapper (Service Layer): "
	for each element in msses
		Session.Output element.Name
	next				
	Session.Output "--------------------------------------------------------"	
	
	'4. Related Interfaces (Service Layer)
	Session.Output "4. Related Interfaces (Service Layer): "
	for each element in ris
		Session.Output element.Name
	next				
	Session.Output "--------------------------------------------------------"	
	
	'5. Impacted Executables (Service Layer)
	Session.Output "5. Impacted Executables (Service Layer): "
	for each element in excs
		Session.Output element.Name
	next				
	Session.Output "--------------------------------------------------------"
	
	'6. Impacted Interfaces (Service Layer)
	Session.Output "6. Impacted Interfaces (Service Layer):"
	for each element in iis
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable
	Session.Output "7.1 Impacted MECOMS Mappers (Service Layer) - Transition from executable:"
	for each element in mmso
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable
	Session.Output "7.2 Impacted MECOMS Mappers (Service Layer) - Transition to executable:"
	for each element in mmsi
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'7.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper)
	Session.Output "7.3 Impacted MECOMS Mappers (Service Layer) - Direct (Mapper to MECOMS Mapper):"
	for each element in mmsd
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'8. Impacted I/FIS (Service Layer)
	Session.Output "8. Impacted I/FIS (Service Layer):"
	for each element in ifisses
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'9. Impacted Interface Agreement (IA) Application Services (Back-End Layer)
	Session.Output "9. Impacted Interface Agreement (IA) Application Services (Back-End Layer):"
	for each element in iass
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'10. Impacted Use Cases - MECOMS (Back-End Layer)
	Session.Output "10. Impacted Use Cases - MECOMS (Back-End Layer):"
	for each element in iucms
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
	
	'11. Impacted interface - MECOMS (Back- End Layer)
	Session.Output "11. Impacted interface - MECOMS (Back-End Layer):"
	for each element in iims
		Session.Output element.Name
	next
	Session.Output "--------------------------------------------------------"
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