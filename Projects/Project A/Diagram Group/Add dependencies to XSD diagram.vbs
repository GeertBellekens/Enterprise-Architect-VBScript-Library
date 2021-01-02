'[path=\Projects\Project A\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.DocGenUtil

'
' This code has been included from the default Diagram Script template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.
'
' Script Name: Add dependencies to XSD diagram
' Author: Matthias Van der Elst
' Purpose: 
' Date: 05/05/2017
'
'
' Diagram Script main function
'
sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()

	if not currentDiagram is nothing then
		' Get a reference to any selected connector/objects
		dim selectedConnector as EA.Connector
		dim selectedObjects as EA.Collection
		set selectedConnector = currentDiagram.SelectedConnector
		set selectedObjects = currentDiagram.SelectedObjects

		if not selectedConnector is nothing then
			' A connector is selected
		elseif selectedObjects.Count > 0 then
			' One or more diagram objects are selected
		else
			' Nothing is selected
			AddDependencies(currentDiagram)
		end if
	else
		Session.Output "This script requires a diagram to be visible", promptOK
	end if

end sub

sub AddDependencies(currentDiagram)
	dim diagramID, sqlGetObjects
	diagramID = currentDiagram.DiagramID
	dim dos
	set dos = CreateObject("System.Collections.ArrayList")
	
	sqlGetObjects = "select o.object_id " & _
					"from t_diagramobjects do " & _
					"inner join t_object o " & _
					"on do.Object_ID = o.Object_ID " & _
					"where do.Diagram_ID = '" & diagramID & "' "
					
	set dos = getElementsFromQuery(sqlGetObjects)
	
	dim element as EA.Element
	for each element in dos
		dim classifiers
		set classifiers = CreateObject("System.Collections.ArrayList")
		dim sqlGetClassifiers
		sqlGetClassifiers =	"select o2.Object_ID " & _
							"from ((t_object o " & _
							"inner join t_attribute a " & _
							"on o.Object_ID = a.Object_ID) " & _
							"inner join t_object o2 " & _
							"on  a.Classifier = o2.Object_ID) " & _
							"where a.Classifier is not null " & _
							"and  o.Object_ID = '" & element.ElementID & "' "
		
		set classifiers = getElementsFromQuery(sqlGetClassifiers)
		
		dim classifier as EA.Element
		for each classifier in classifiers
			' Kijken of het type van het attribuut ook aanwezig is op het diagram
			if contains(classifier, currentDiagram.DiagramObjects) then
				' Eerst controleren of de dependency nog niet bestaat
				dim dependencies
				set dependencies = CreateObject("System.Collections.ArrayList")
				dim sqlGetDependencies
				sqlGetDependencies = 	"select con.Connector_ID " & _ 
										"from t_connector con " & _
										"where con.Start_Object_ID = '" & element.ElementID & "' and con.End_Object_ID = '" & Classifier.ElementID & "' "
				
				set dependencies = getConnectorsFromQuery(sqlGetDependencies)
				if dependencies.count() = 0 then
					' Indien zo, een dependency maken van de owner van het attribuut naar het type van dat attribuut
					dim dependency as EA.Connector
					set dependency = element.Connectors.AddNew("", "Dependency")
					dependency.SupplierID = classifier.ElementID
			
					' Bij target end de multipliciteit overnemen van dat attribuut. Default is dit 1..1, indien ingevuld die gebruiken.
					dim attributes
					set attributes = CreateObject("System.Collections.ArrayList")
					dim sqlGetAttributes
					sqlGetAttributes = 	"select att.ID " & _
										"from t_attribute att " & _
										"where att.Classifier = '" & Classifier.ElementID & "' "
					
					set attributes =  getAttributesByQuery(sqlGetAttributes)
					dim attribute as EA.Attribute
					for each attribute in attributes
						dependency.SupplierEnd.Cardinality = attribute.LowerBound & ".." & attribute.UpperBound
					next
					dependency.Update()
				
				end if
			end if
		next
		
	next
	Session.Output "Dependencies added"
end sub


function contains(classifier, result)
	contains = false
	dim res as EA.Element
	for each res in result
		if res.ElementID = classifier.ElementID then
			contains = true
			exit for
		end if	
	next
end function

OnDiagramScript