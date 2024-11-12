'[path=\Projects\Project BF]
'[group=Belfius]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Application Portfolio
' Author: Geert Bellekens
' Purpose: import applications
' Date: 2023-11-07

const outPutName = "Import Application PortFolio"
const obsoletePackageGUID = "{0C293A84-542F-4702-824C-B460954D0A82}"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName, 0
	'Actual work
	importApplications
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName, 0
end sub

function importApplications()
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage
	'report progress
	Repository.WriteOutput outPutName, now() & " Selected Package '" & package.Name & "'", 0
	dim importFile
	set importFile = new TextFile
	if importFile.UserSelect("","XML Files (*.xml)|*.xml") then
		dim xmlDOM 
		set xmlDOM = CreateObject("MSXML2.DOMDocument")
		If xmlDOM.LoadXML(importFile.Contents) Then
			Repository.WriteOutput outPutName, now() & " Reading XML file succeeded", 0
			processApplicationsXML package, xmlDom
		else
			'error loading xml file
			Repository.WriteOutput outPutName, now() & " Error loading xmlFile " & importFile.FullPath, 0
		end if
	end if	
end function

function processApplicationsXML(package, xmlDom)
	dim applicationNodes
	set applicationNodes = xmlDOM.SelectNodes("//application")
	dim applicationNode
	dim applications
	set applications = CreateObject("System.Collections.ArrayList")
	for each applicationNode in applicationNodes
		applications.Add importApplication(package, applicationNode)
	next
	addApplicationsToDiagram package, applications
end function

function addApplicationsToDiagram(package, applications)
	'create diagram
	dim diagram as EA.Diagram
	set diagram = package.Diagrams.AddNew("Applications Import", "Logical")
	diagram.Update
	'add applications to diagram
	dim application as EA.Element
	for each application in applications
		dim diagramObject as EA.Element
		set diagramObject = diagram.DiagramObjects.AddNew("","")
		diagramObject.ElementID = application.ElementID
		diagramObject.Update
	next
	'layout diagram
	dim diagramGUIDXml
	'The project interface needs GUID's in XML format, so we need to convert first.
	diagramGUIDXml = Repository.GetProjectInterface().GUIDtoXML(diagram.DiagramGUID)
	'Then call the layout operation
	Repository.GetProjectInterface().LayoutDiagramEx diagramGUIDXml, lsDiagramDefault, 4, 20 , 20, false
	'open the diagram
	Repository.OpenDiagram(diagram.DiagramID)
end function

function importApplication(package, applicationNode)
	dim applicationName
	applicationName = applicationNode.getAttribute("name")
	dim applicationID
	applicationID = applicationNode.getAttribute("id")
	dim obsolete
	obsolete = applicationNode.getAttribute("obsolete")
	
	Repository.WriteOutput outPutName, now() & " Processing application '" & applicationName & "'" , 0
	dim application as EA.Element
	set application = getExistingApplicationSQL(package, applicationID)

	dim notesNode
	set notesNode = applicationNode.SelectSingleNode("notes")
	dim notes
	notes = ""
	if not notesNode is nothing then
		notes = notesNode.Text
	end if
	if application is nothing then
		set application = package.Elements.AddNew(applicationName, "ArchiMate3::ArchiMate_ApplicationComponent")
		Repository.WriteOutput outPutName, now() & " created application '" & application.Name & "'" , 0
		application.Update
		dim idTag as EA.TaggedValue
		set idTag  = application.TaggedValues.AddNew("applicationID", "")
		idTag.Value = applicationID
		idTag.Update
	end if
	application.Name = applicationName
	application.Notes = notes
	application.Update
	if lcase(obsolete) = "true" then
		'move the application to the obsolete package
		dim obsoletePackage as EA.Package
		set obsoletePackage = Repository.GetPackageByGuid(obsoletePackageGUID)
		application.PackageID = obsoletePackage.PackageID
		application.Update
	end if
	'return
	set importApplication = application
end function

function getExistingApplication(package, applicationID)
	dim application as EA.Element
	set application = nothing
	dim element as EA.Element
	for each element in package.Elements
		if element.Alias = applicationID then
			set application = element
		end if
	next
	'recurse down
	if application is nothing then
		dim subPackage as EA.Package
		for each subPackage  in package.Packages
			set application = getExistingApplication(subPackage, applicationID)
			if not application is nothing then
				exit for
			end if
		next
	end if
	'return
	set getExistingApplication = application
end function

function getExistingApplicationSQL(package, applicationID)
	dim application
	set application = nothing
	dim packageTreeIDString
	packageTreeIDString = getPackageTreeIDString(package)
	dim sqlGetData
	sqlGetData = "select o.Object_ID from t_object o                                " & vbNewLine & _
				" inner join t_objectproperties tv on tv.Object_ID = o.Object_ID   " & vbNewLine & _
				" 						and tv.Property = 'applicationID'          " & vbNewLine & _
				" where tv.Value = '"& applicationID & "'                          " & vbNewLine & _
				" and o.Package_ID in (" & packageTreeIDString & ")                "
	dim applications
	set applications = getElementsFromQuery(sqlGetData)
	if applications.Count > 0 then
		set application = applications(0)
	end if
	'return
	set getExistingApplicationSQL = application
end function

main