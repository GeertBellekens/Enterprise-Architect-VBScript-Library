option explicit

	!INC Local Scripts.EAConstants-VBScript
	!INC Atrias Scripts.Util

	'
	' Script Name: 
	' Author: 
	' Purpose: 
	' Date: 
	'
	Dim XSDBaseTypes
	XSDBaseTypes = Array("string","boolean","decimal","float","double","duration","dateTime","time","date","gYearMonth","gYear","gMonthDay","gDay","gMonth","hexBinary","base64Binary","anyURI","QName","integer","long","int")

	sub main
		dim response
		response = Msgbox("This script will move all underlying elements into one package!" & vbnewLine & "This should only be done when making an XSD from the LDM." & vbnewLine & " Are you sure?", vbYesNo+vbExclamation, "Post XSD transformation")
		if response = vbYes then
			'Create new package for the whole of the schema
			dim package as EA.Package 
			set package = Repository.GetTreeSelectedPackage()
			dim schemaPackage as EA.Package
			set schemaPackage = package.Packages.AddNew(package.Name,"package")
			schemaPackage.Update
			schemaPackage.Element.Stereotype = "XSDSchema"
			schemaPackage.Update
			' move all elements from the subpackages to the newly create package
			mergeToSchemaPackage package, schemaPackage
			' fix the elements
			fixElements schemaPackage
			'fix the connectors
			fixConnectors schemaPackage
			'fix the attributes with a primitive type
			fixAttributePrimitives schemaPackage
			'reload
			Repository.RefreshModelView(package.PackageID)
			msgbox "Finished!"
		end if
	end sub

	function fixElements(schemaPackage)
		dim element as EA.Element
		for each element in schemaPackage.Elements
			if element.Stereotype = "XSDsimpleType" then
				if element.Attributes.Count = 0 then
					' fix xsdSimpleTypes
					fixXSDsimpleType element
				else
					element.Stereotype = "XSDComplexType"
					element.Update
				end if
			end if
		next	
	end function

	function fixXSDsimpleType(element)
		'find the element it was transformed from
		dim sourceElement as EA.Element
		dim sqlFindSource
		dim sourceElements
		sqlFindSource = "select o.[Object_ID] from  t_object o " & _
						"inner join t_xref x on x.[Supplier] = o.[ea_guid] " & _
						"where x.TYPE = 'Transformation' " & _
						"and x.[Client] =  '" & element.ElementGUID & "'"
		set sourceElements = getElementsFromQuery(sqlFindSource)
		if sourceElements.Count > 0 then
			set sourceElement = sourceElements(0)
			'copy the tagged values
			copyTaggedValues sourceElement, element
			'determine the "parent" type
			dim baseClass as EA.Element
			dim baseXSDType
			baseXSDType = "string" 'default
			for each baseClass in sourceElement.BaseClasses
				if Ubound(Filter(XSDBaseTypes, baseClass.Name )) > -1 then
					'found the base class
					baseXSDType = baseClass.Name
				end if
			next
			'set the base type
			element.Genlinks = "Parent=" & baseXSDType & ";"
			element.Update
		end if
	end function


	function mergeToSchemaPackage (package, schemaPackage)
		dim subPackage as EA.Package
		dim i
		for i = package.Packages.Count -1 to i = 1 step -1
			set subPackage = package.Packages.GetAt(i)
			'should only be done on XSDschema packages
			if subPackage.Element.Stereotype = "XSDschema" then
				dim element as EA.Element
				'move elements
				for each element in subPackage.Elements
					element.PackageID = schemaPackage.PackageID
					element.Update
				next
				'remove original package
				package.Packages.DeleteAt i,false
			end if
		next
	end function

	function fixConnectors(package)
		dim SQLgetConnectors
		SQLgetConnectors = "select distinct c.Connector_ID from " & _
							" ( " & _
							" select source.StartID, source.EndID from  " & _
							" ( " & _
							" select o.[Object_ID] AS StartID, con.[End_Object_ID] AS EndID " & _
							" from (t_object o " & _
							" inner join t_connector con on con.[Start_Object_ID] = o.[Object_ID]) " & _
							" where o.package_ID = "& package.PackageID & _
							" union all " & _
							" select o.[Object_ID] AS StartID, con.[Start_Object_ID] AS EndID " & _
							" from (t_object o " & _
							" inner join t_connector con on con.[End_Object_ID] = o.[Object_ID]) " & _
							" where o.package_ID = "& package.PackageID & _
							" ) source " & _
							" group by source.StartID, source.EndID " & _
							" having count(*) > 1 " & _
							" ) grouped, t_connector c " & _
							" where (c.[Start_Object_ID] = grouped.StartID  " & _
							"       and c.[End_Object_ID] = grouped.EndID) " & _
							"       or " & _
							"       (c.[Start_Object_ID] = grouped.EndID  " & _
							"       and c.[End_Object_ID] = grouped.StartID) "
		dim connectors
		set connectors = getConnectorsFromQuery(SQLgetConnectors)
		dim connector as EA.Connector
		dim changed
		changed = false
		for each connector in connectors
			if len(connector.ClientEnd.Role) < 1 then
				connector.ClientEnd.Role = replace(connector.Name, " ", "_")
				connector.ClientEnd.Update
				changed = true
			end if
			if len(connector.SupplierEnd.Role) < 1 then
				connector.SupplierEnd.Role = replace(connector.Name, " ", "_")
				connector.SupplierEnd.Update
				changed = true
			end if
			if changed then
				connector.Update
			end if
		next
	end function

	function fixAttributePrimitives(package)	
		
		dim sqlUpdate
		sqlUpdate = "update attr set attr.Classifier = 0  " & _
					" from t_attribute attr  " & _
					" inner join t_object o on attr.object_id = o.object_id " & _
					" where attr.[TYPE] in ('string','boolean','decimal','float','double','duration','dateTime','time','date','gYearMonth','gYear','gMonthDay','gDay','gMonth','hexBinary','base64Binary','anyURI','QName','integer','long','int') " & _
					" and attr.Classifier > 0 " & _
					" and o.[Package_ID] =  " & package.PackageID
		Repository.Execute sqlUpdate
	end function

	main