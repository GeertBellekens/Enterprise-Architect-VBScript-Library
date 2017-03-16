'[path=\Projects\Project M\Diagram Group]
'[group=Diagram Group]
'[group_type=DIAGRAM]

option explicit

!INC Local Scripts.EAConstants-VBScript


'
' Script Name: Set Default Diagram Template
' Author: Geert Bellekens
' Purpose: Copy the elements from the default diagram in the template package
' Date: 2016-06-10
'

'
' Diagram Script main function
'
sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	'set currentDiagram = Repository.GetDiagramByGuid("{BE3217CD-E80B-4848-ACF5-135A1D61BCC2}")
	set currentDiagram = Repository.GetCurrentDiagram()

	if not currentDiagram is nothing then
		'get template package
		dim templatePackage as EA.Package
		set templatePackage = getTemplatePackage
		'get the corresponding diagram
		dim templateDiagram as EA.Diagram
		set templateDiagram = getCorrespondingDiagram(currentDiagram, templatePackage)
		if not templateDiagram is nothing then
			copyDiagramTemplate currentDiagram, templateDiagram
			Repository.ReloadDiagram currentDiagram.DiagramID
		else
			msgbox "No template diagram found for this type of diagram: " & vbNewLine & currentDiagram.Type & " - " & currentDiagram.MetaType _
					, vbExclamation, "No template diagram found"
		end if 
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if

end sub

'copy all elements from the template diagram to the current diagram
'depending on the type of element we make a duplicate or use a link to the same element (a bit like a smart copy)
function copyDiagramTemplate(currentDiagram, templateDiagram)
	'loop all diagramObjects
	dim ownerPackage as EA.Package
	set ownerPackage = Repository.GetPackageByID(currentDiagram.PackageID)
	dim diagramObject as EA.DiagramObject
	for each diagramObject in templateDiagram.DiagramObjects
		dim element as EA.Element
		set element = Repository.GetElementByID(diagramObject.ElementID)
		select case element.Type
		case "Text", "Note", "Boundary"
			set element = duplicate(element, ownerPackage)
		end select
		dim newDiagramObject as EA.DiagramObject
		'add the element to the  diagram
		set newDiagramObject = currentDiagram.DiagramObjects.AddNew("","")
		'copy properties of the diagramobject
		newDiagramObject.ElementID = element.ElementID
		newDiagramObject.left = diagramObject.left
		newDiagramObject.right = diagramObject.right
		newDiagramObject.top = diagramObject.top
		newDiagramObject.bottom = diagramObject.bottom
		newDiagramObject.Style = diagramObject.Style
		newDiagramObject.Sequence = diagramObject.Sequence
		'save diagramObject
		newDiagramObject.Update
	next
end function

function duplicate(original, ownerPackage)
'	dim original as EA.Element
'	dim ownerPackage as EA.Package
	dim newElement as EA.Element
	set newElement = ownerPackage.Elements.AddNew(original.Name, original.Type)
	'copy all features
	newElement.Notes= original.Notes
	newElement.StereotypeEx = original.StereotypeEx
	newElement.StyleEx = original.StyleEx
	newElement.Subtype = original.Subtype
	'save new element
	newElement.Update
	'return element
	set duplicate = newElement
end function

function getCorrespondingDiagram(currentDiagram, templatePackage)
	'initialize at nothing
	set getCorrespondingDiagram = nothing
	'loop diagrams
	dim diagram as EA.Diagram
	for each diagram in templatePackage.Diagrams
		'look for the first diagram with the same type and stereotype
		if diagram.Type = currentDiagram.Type AND _
			diagram.Stereotype = currentDiagram.Stereotype and _
			diagram.MetaType = currentDiagram.MetaType then
			set getCorrespondingDiagram = diagram
			exit for
		end if
	next
end function

function getTemplatePackage()
	'initialize at nothing
	set getTemplatePackage = nothing
	dim sqlGetPackageObject 
	sqlGetPackageObject = "select o.Object_ID from ((t_package p " & _
							" inner join usys_system syst on (syst.Property = 'TemplatePkg' " & _
							"  								and syst.Value = p.Package_ID)) " & _
							" inner join t_object o on o.ea_guid = p.ea_guid) "
    dim packageObjectCollection
	set packageObjectCollection = Repository.GetElementSet(sqlGetPackageObject, 2)					
	dim packageObject as EA.Element
	dim templatePackage as EA.Package
	for each packageObject in packageObjectCollection
		set templatePackage = Repository.GetPackageByGuid(packageObject.ElementGUID)
		set getTemplatePackage = templatePackage
		exit for
	next
end function

OnDiagramScript