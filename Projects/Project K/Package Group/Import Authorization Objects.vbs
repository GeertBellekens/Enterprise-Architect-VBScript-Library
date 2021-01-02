'[path=\Projects\Project K\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: Import Authorization Objects
' Author: Geert Bellekens
' Purpose: Import bevoegdheidsobjecten from an xml file, see also https://help.sap.com/doc/saphelp_470/4.7/en-US/52/671285439b11d1896f0000e8322d00/content.htm?no_cache=true
' Date: 2020-03-27

const outPutName = "Import Authorization Objects"
const profilePackageGUID = "{F60260DF-B4AD-4f8e-A0C7-1DC3211D0BD0}"
const singleRolePackageGUID = "{297D8B49-1723-4cad-84F0-B96393E5E929}"
const compositeRolePackageGUID = "{CF3E3DD8-202B-4746-A053-69A4F2C34A4F}"

sub main
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting Import Authorization Objects", 0
		'get selected package
		dim package as EA.Package
		set package = Repository.GetTreeSelectedPackage()
		'exit if not selected
		if package is nothing then
			msgbox "Please select a package before running this script"
			exit sub
		end if
		'start import
		importXmlFile(package)
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Finished Import Authorization Objects", 0
end sub

function importXmlFile(package)
	dim importFile
	set importFile = new TextFile
	if importFile.UserSelect("","XML Files (*.xml)|*.xml") then
		dim xmlDOM 
		set xmlDOM = CreateObject("MSXML2.DOMDocument")
		If xmlDOM.LoadXML(importFile.Contents) Then
			'import authorization objects
			importAuthorizationObjects package, xmlDOM
			'import profiles
			'create profiles dictionary
			dim profilesDictionary
			set profilesDictionary = CreateObject("Scripting.Dictionary")
			importProfiles xmlDOM, package, profilesDictionary
			'link sub-profiles
			processSubProfiles profilesDictionary
			'import singleRoles
			importSingleRoles xmlDOM, package
			'import composite roles
			importCompositeRoles xmlDOM
			'reload package
			Repository.ReloadPackage package.PackageID
		else
			'error loading xml file
			Repository.WriteOutput outPutName, now() & " Error loading xmlFile " & importFile.FullPath, 0
		end if
	end if	
end function


function importCompositeRoles(xmlDOM)
	'get single role package
	dim compositeRolePackage as EA.Package
	set compositeRolePackage = Repository.GetPackageByGuid(compositeRolePackageGUID)
	'get roleNodes
	dim roleNodes
	set roleNodes = xmlDOM.SelectNodes("//compositeroles/compositerole")
	dim roleNode
	for each roleNode in roleNodes
		'get name
		dim nameNode
		set nameNode = roleNode.SelectSingleNode("./name")
		dim roleName
		roleName = nameNode.Text
		'inform user
		Repository.WriteOutput outPutName, now() & " Processing composite role '" & roleName & "'" , 0
		'get composite role element
		dim compositeRoleElement as EA.Element
		set compositeRoleElement = getCompositeRoleElement(roleName, compositeRolePackage)
		'process linked roles
		 processLinkedRoles compositeRoleElement, roleNode
	next
end function

function processLinkedRoles( compositeRoleElement, roleNode)
	'get single role package
	dim singleRolePackage as EA.Package
	set singleRolePackage = Repository.GetPackageByGuid(singleRolePackageGUID)
	'get single Role nodes
	dim singleRoleNodes
	set singleRoleNodes = roleNode.SelectNodes("./singleRole")
	'loop profileNodes
	dim singleRoleNode
	for each singleRoleNode in singleRoleNodes
		dim singleRoleElement as EA.Element
		set singleRoleElement = getSingleRoleElement(singleRoleNode.Text, singleRolePackage)
		'set use relation
		addUsageConnector compositeRoleElement, singleRoleElement
	next
end function

function getCompositeRoleElement(roleName, compositeRolePackage)
	'use query to get existing composite role
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                         " & vbNewLine & _
				" where o.Package_ID = " & compositeRolePackage.PackageID      & vbNewLine & _
				" and o.stereotype = 'SAP_compositeRole'                     " & vbNewLine & _
				" and o.Name = '" & roleName & "'                            "
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	dim element as EA.Element
	if elements.Count > 0 then
		set element = elements(0)
	else
		'add new profile element
		set element = compositeRolePackage.Elements.AddNew(roleName,"Class")
		element.Stereotype = "SAP_compositeRole"
		element.Update
	end if	
	'return
	set getCompositeRoleElement = element
end function

function importSingleRoles(xmlDOM, package)
	'get single role package
	dim singleRolePackage as EA.Package
	set singleRolePackage = Repository.GetPackageByGuid(singleRolePackageGUID)
	'get roleNodes
	dim roleNodes
	set roleNodes = xmlDOM.SelectNodes("//singleroles/singlerole")
	dim roleNode
	for each roleNode in roleNodes
		'get name
		dim nameNode
		set nameNode = roleNode.SelectSingleNode("./name")
		dim roleName
		roleName = nameNode.Text
		'inform user
		Repository.WriteOutput outPutName, now() & " Processing single role '" & roleName & "'" , 0
		'get single role element
		dim singleRoleElement as EA.Element
		set singleRoleElement = getSingleRoleElement(roleName, singleRolePackage)
		'process profiles
		processRoleProfiles singleRoleElement, roleNode
		'process authorizations
		processLinkedAuthorizations roleNode, singleRoleElement, package
	next		
end function

function processRoleProfiles( singleRoleElement, roleNode)
	'get profile package
	dim profilePackage as EA.Package
	set profilePackage = Repository.GetPackageByGuid(profilePackageGUID)
	'get profileNodes
	dim profileNodes
	set profileNodes = roleNode.SelectNodes("./profile")
	'loop profileNodes
	dim profileNode
	for each profileNode in profileNodes
		dim profileElement as EA.Element
		set profileElement = getProfileElement(profileNode.Text, profilePackage)
		'set use relation
		addUsageConnector singleRoleElement, profileElement
	next
end function

function getSingleRoleElement(roleName, singleRolePackage)
	'use query to get existing profile
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                      " & vbNewLine & _
				" where o.Package_ID = " & singleRolePackage.PackageID & "   " & vbNewLine & _
				" and o.stereotype = 'SAP_singleRole'                        " & vbNewLine & _
				" and o.Name = '" & roleName & "'                      "
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	dim element as EA.Element
	if elements.Count > 0 then
		set element = elements(0)
	else
		'add new profile element
		set element = singleRolePackage.Elements.AddNew(roleName,"Class")
		element.Stereotype = "SAP_singleRole"
		element.Update
	end if	
	'return
	set getSingleRoleElement = element
end function


function processSubProfiles(profilesDictionary)
	'get profile package
	dim profilePackage as EA.Package
	set profilePackage = Repository.GetPackageByGuid(profilePackageGUID)
	'loop profile
	dim profile as EA.Element
	for each profile in profilesDictionary.Keys
		dim subProfileName
		for each subProfileName in profilesDictionary(profile)
			'get the coresponding profile
			dim subProfile
			set subProfile = getProfileElement(subProfileName, profilePackage)
			'add usage relations
			addUsageConnector profile, subProfile
		next
	next
end function

function addUsageConnector(source, target)
	'link profile to subProfile
	dim relationExists
	relationExists = false
	'check if relation exists
	dim relation as EA.Connector
	for each relation in source.Connectors
		if relation.SupplierID = target.ElementID _
		  and relation.Type = "Usage" then
			relationExists = true
			exit for
		end if
	next
	'add relation if needed
	if not relationExists then
		'add use relation
		dim useRelation as EA.Connector
		set useRelation = source.Connectors.AddNew("","Usage")
		useRelation.SupplierID = target.ElementID
		useRelation.Update
	end if
end function 

function importAuthorizationObjects(package, xmlDOM)
	'get package nodes
	dim packageNodes
	set packageNodes = xmlDOM.SelectNodes("//authorizationclass")
	dim packageNode
	for each packageNode in packageNodes
		'get name
		dim nameNode
		set nameNode = packageNode.SelectSingleNode("./name")
		'inform user
		Repository.WriteOutput outPutName, now() & " Processing package '" & nameNode.Text & "'" , 0
		'get package (new or existing)
		dim boPackage as EA.Package
		set boPackage = getBoPackage(package, nameNode.Text)
		'get diagram
		dim newDiagram
		dim diagram as EA.Diagram
		set diagram = getDiagram(boPackage, newDiagram)
		'process bevoegheidsObjects
		processBevoegheidsObjects packageNode, boPackage, diagram
		'format diagram if new
		if newDiagram then
				dim diagramGUIDXml
				'The project interface needs GUID's in XML format, so we need to convert first.
				diagramGUIDXml = Repository.GetProjectInterface().GUIDtoXML(diagram.DiagramGUID)
				'Then call the layout operation
				Repository.GetProjectInterface().LayoutDiagramEx diagramGUIDXml, lsDiagramDefault, 4, 20 , 20, false
		end if
	next
end function

function getDiagram(boPackage, newDiagram)
	dim diagram as EA.Diagram
	set diagram = nothing
	'get first diagram
	dim tempDiagram
	for each tempDiagram in boPackage.Diagrams
		set diagram = tempDiagram
		newDiagram = false
		exit for
	next
	'create diagram if not exists
	if diagram is nothing then
		set diagram = boPackage.Diagrams.AddNew(boPackage.Name, "Logical")
		diagram.Update
		newDiagram = true
	end if
	'return
	set getDiagram = diagram
end function

function processBevoegheidsObjects(packageNode, boPackage, diagram)
	'get bevoegheidsobject nodes
	dim boNodes
	set boNodes = packageNode.SelectNodes("./authorizationobject")
	dim boNode
	dim boElements
	set boElements = CreateObject("Scripting.Dictionary")
	for each boNode in boNodes
		'get name
		dim nameNode
		set nameNode = boNode.SelectSingleNode("./name")
		'inform user
		Repository.WriteOutput outPutName, now() & " Processing authorizationobject '" & nameNode.Text & "'" , 0
		'get bevoegdheidsObject element (new or existing)
		dim boElement as EA.Element
		set boElement = getBoElement(boPackage, nameNode.Text)
		'set description
		dim descriptionNode
		set descriptionNode = boNode.SelectSingleNode("./description")
		if not descriptionNode is nothing then
			boElement.Notes = descriptionNode.Text
			boElement.Update
		end if
		'add boElement to list
		boElements.Add boElement.ElementID, boElement
		'add boElement to diagram
		addElementToDiagram boElement, diagram, 200, 200
		'process properties
		processProperties boNode, boElement
		'process authorizations
		processRoles boNode, boElement, diagram
	next
	'remove all other elements
	dim subElement as EA.Element
	dim i
	'refresh
	boPackage.Elements.Refresh
	for i = 0 to boPackage.Elements.Count -1
		set subElement = boPackage.Elements.Getat(i)
		if not boElements.Exists(subElement.ElementID) _
		  and subElement.Stereotype = "SAP_authorizationobject" then
			boPackage.Elements.DeleteAt i, false
		end if
	next
end function

function processRoles(boNode, boElement, diagram)
	'get role nodes
	dim roleNodes
	set roleNodes =  boNode.SelectNodes("./authorizations/authorization")
	'role elements
	dim roleElements
	set roleElements = CreateObject("Scripting.Dictionary")
	dim roleNode
	for each roleNode in roleNodes
		'get name
		dim nameNode
		set nameNode = roleNode.SelectSingleNode("./name")
		'get role element
		dim roleElement as EA.Element
		set roleElement = getRoleElement(boElement, nameNode.text)
		'add to dictionary
		if not roleElements.Exists(roleElement.ElementID) then
			roleElements.Add roleElement.ElementID, roleElement
		end if
		'add to diagram
		addElementToDiagram roleElement, diagram, 400, 400
		'set runstate
		processRoleProperties roleElement, roleNode
	next
	'remove all other elements
	dim subElement as EA.Element
	dim i
	'refresh
	boElement.Elements.Refresh
	for i = 0 to boElement.Elements.Count -1
		set subElement = boElement.Elements.Getat(i)
		if not roleElements.Exists(subElement.ElementID) then
			boElement.Elements.DeleteAt i, false
		end if
	next
end function

function importProfiles(xmlDOM, package, profilesDictonary)
	'get profile package
	dim profilePackage as EA.Package
	set profilePackage = Repository.GetPackageByGuid(profilePackageGUID)
	'get profileNodes
	dim profileNodes
	set profileNodes = xmlDOM.SelectNodes("//profiles/profile")
	'loop profileNodes
	dim profileNode
	for each profileNode in profileNodes
		'get name
		dim nameNode
		set nameNode = profileNode.SelectSingleNode("./name")
		dim profileName
		profileName = nameNode.Text
		'inform user
		Repository.WriteOutput outPutName, now() & " Processing profile '" & profileName & "'" , 0
		'get profile
		dim profileElement as EA.Element
		set profileElement = getProfileElement(profileName, profilePackage)
		'process authorizations
		processLinkedAuthorizations profileNode, profileElement, package
		'get all subProfiles
		dim subProfiles
		set subProfiles = getSubProfileNames(profileNode, profileNames)
		'add the profileElement to the dictionary
		profilesDictonary.Add profileElement, subProfiles
	next
end function

function getSubProfileNames(profileNode, profileNames)
	'create ArrayList
	dim subProfiles
	set subProfiles = CreateObject("System.Collections.ArrayList")
	'loop profile nodes
	dim subprofileNodes
	set subprofileNodes = profileNode.SelectNodes("./profile")
	dim subProfileNode
	for each subProfileNode in subprofileNodes
		subProfiles.Add subProfileNode.Text
	next
	'return
	set getSubProfileNames = subProfiles
end function

function processLinkedAuthorizations(node, element, package)
	'loop authorization nodes
		dim authorizationNodes
		set authorizationNodes = node.SelectNodes("./authorization")
		dim authorizationNode
		for each authorizationNode in authorizationNodes
			'add relation to authorizations
			dim authorizationElement
			set authorizationElement = getAuthorizationElement(authorizationNode.Text, package)
			'report error if authorizationElement not found
			if authorizationElement is nothing then
				'Report error
				Repository.WriteOutput outPutName, now() & " ERROR: Authorization '" & authorizationNode.Text & "' for " & element.Stereotype & " '" & element.Name & "' not found!" , 0
			else
				'add usage relation
				addUsageConnector element, authorizationElement
			end if
		next
end function

function getProfileElement(profileName, profilePackage)
	'use query to get existing profile
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                      " & vbNewLine & _
				" where o.Package_ID = " & profilePackage.PackageID & "   " & vbNewLine & _
				" and o.stereotype = 'SAP_profile'                        " & vbNewLine & _
				" and o.Name = '" & profileName & "'                      "
	dim profileElements
	set profileElements = getElementsFromQuery(sqlGetData)
	dim profileElement as EA.Element
	if profileElements.Count > 0 then
		set profileElement = profileElements(0)
	else
		'add new profile element
		set profileElement = profilePackage.Elements.AddNew(profileName,"Class")
		profileElement.Stereotype = "SAP_profile"
		profileElement.Update
	end if	
	'return
	set getProfileElement = profileElement
end function

function processRoleProperties(roleElement, roleNode)
	dim runstateString
	'get propertyNodes
	dim propertyNodes
	set propertyNodes = roleNode.SelectNodes("./properties/property")
	'loop propertyNodes
	dim propertyNode
	for each propertyNode in propertyNodes
		'get name
		dim nameNode
		set nameNode = propertyNode.SelectSingleNode("./name")
		'start of the runstatestring
		runstateString = runstateString & "@VAR;Variable=" & nameNode.Text
		'get the values
		dim valueString
		valueString = ""
		dim valueNodes
		set valueNodes = propertyNode.SelectNodes("./value")
		dim valueNode
		for each valueNode in valueNodes
			if len(valueString) > 0 then
				valueString = valueString & ","
			end if
			valueString = valueString & valueNode.Text
		next
		'add the valueString to the runstatestring
		runstateString = runstateString & ";Value=" & valueString & ";Op==;@ENDVAR;"
	next
	'set runstate on element
	if len(runstateString) > 0 then
		roleElement.RunState = runstateString
		roleElement.Update
	end if
end function

function getAuthorizationElement(name, package)
	'get the package tree id's
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o      " & vbNewLine & _
				" where o.Name = '" & name & "'           " & vbNewLine & _
				" and o.Stereotype = 'SAP_authorization'  " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDs & ")"
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	dim element as EA.Element
	if elements.Count > 0 then
		set element = elements(0)
	else
		set element = nothing
	end if
	'return
	set getAuthorizationElement = element
end function


function getRoleElement(boElement, roleName)
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o       " & vbNewLine & _
				" where o.Stereotype = 'SAP_authorization' " & vbNewLine & _
				" and o.Classifier = " & boElement.ElementID & vbNewLine & _
				" and o.ParentID = " & boElement.ElementID   & vbNewLine & _
				" and o.Name = '" & roleName & "'"
	dim roleElements
	set roleElements = getElementsFromQuery(sqlGetData)
	dim roleElement as EA.Element
	if roleElements.Count > 0 then
		set roleElement = roleElements(0)
	else
		'create new one
		set roleElement = boElement.Elements.AddNew(roleName,"Object")
		roleElement.ClassifierID = boElement.ElementID
		roleElement.Stereotype = "SAP_authorization"
		roleElement.Update
	end if
	'return
	set getRoleElement = roleElement
end function

function processProperties(boNode, boElement)
	'get propertyNodes node
	dim propertyNodes
	set propertyNodes = boNode.SelectNodes("./properties/property")
	dim boAttributes
	set boAttributes = CreateObject("Scripting.Dictionary")
	'loop propertynodes
	dim propertyNode
	for each propertyNode in propertyNodes
		'get name
		dim nameNode
		set nameNode = propertyNode.SelectSingleNode("./name")
		'get attribute
		dim boAttribute as EA.Attribute
		set boAttribute = getBoAttribute(boElement, nameNode.Text)
		'set description
		dim descriptionNode
		set descriptionNode = propertyNode.SelectSingleNode("./description")
		if not descriptionNode is nothing then
			boAttribute.Notes = descriptionNode.Text
			Session.Output "boAttribute.Name: " &  boAttribute.Name
			boAttribute.Update
		end if
		'add to the dictionary
		if not boAttributes.Exists(boAttribute.AttributeID) then
			boAttributes.Add boAttribute.AttributeID, boAttribute
		end if
	next
	'remove other attributes
	'refresh to make sure we have all attributes
	boElement.Attributes.Refresh
	dim i
	dim attribute as EA.Attribute
	for i = 0 to boElement.Attributes.Count -1
		set attribute = boElement.Attributes.Getat(i)
		if not boAttributes.Exists(attribute.AttributeID) then
			boElement.Attributes.DeleteAt i, false
		end if
	next
	 
end function



function getBoAttribute(boElement, propertyName)
	dim boAttribute as EA.Attribute
	set boAttribute = nothing
	dim attr as EA.Attribute
	'get existing attribute
	for each attr in boElement.Attributes
		if attr.Name = propertyName then
			set boAttribute = attr
			exit for
		end if
	next
	'create new attribute
	if boAttribute is nothing then
		set boAttribute = boElement.Attributes.AddNew(propertyName, "")
		boAttribute.Update
	end if
	'return
	set getBoAttribute = boAttribute
end function

function getBoElement(boPackage, boName)
	'build query to get package elements
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o        " & vbNewLine & _
				" where o.Stereotype = 'SAP_authorizationobject' " & vbNewLine & _
				" and o.name = '" & boName & "'               " & vbNewLine & _
				" and o.Package_ID =  " & boPackage.PackageID
	'loop elements
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	dim boElement as EA.Element
	if elements.Count > 0 then
		'get existing element
		set boElement = elements(0)
	else
		'create new element
		set boElement = boPackage.Elements.AddNew(boName, "Class")
		boElement.Stereotype = "SAP_authorizationobject"
		boElement.Update
	end if
	'return
	set getBoElement = boElement
end function

function getBoPackage(package, packageName)
	'get the package tree id's
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	'build query to get package elements
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_package p 				" & vbNewLine & _
				 " inner join t_object o on o.ea_guid = p.ea_guid   " & vbNewLine & _
				 " where p.Name = '" & packageName & "'			    " & vbNewLine & _
				 " and o.Stereotype = 'SAP_authorisationclass'		" & vbNewLine & _
				 " and (p.Package_ID in (" & packageTreeIDs & ")    " & vbNewLine & _
				 " or p.Parent_ID in (" & packageTreeIDs & "))"
	'loop elements
	dim elements
	set elements = getElementsFromQuery(sqlGetData)
	dim packageElement as EA.Element
	dim boPackage
	'initialize at nothing
	set boPackage  = nothing
	for each packageElement in elements
		set boPackage = Repository.GetPackageByGuid(packageElement.ElementGUID)
		'stop after first package
		exit for
	next
	'check if we found an existing package. If not then create a new package
	if boPackage is nothing then
		set boPackage = package.Packages.AddNew(packageName,"")
		boPackage.update
		boPackage.Element.Stereotype = "SAP_authorisationclass"
		boPackage.Element.Update
	end if
	'return
	set getBoPackage = boPackage
end function

main