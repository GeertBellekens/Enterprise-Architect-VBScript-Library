'[path=\Projects\Project A\Search Group]
'[group=Search Group]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Atrias Scripts.DocGenUtil

' Script Name: Quality Control - Business Process Layer
' Author: Matthias Van der Elst
' Purpose: Generates a quality report for every domain in the business process layer
' Date: 2017-04-14

sub main()
	Repository.ClearOutput "Script"
	Session.Output "Quality Control - Business Process Layer"
	Session.Output "--------------------------------------------"
	dim bpGUID
	dim bp as EA.Element
	bpGUID = "{148E3EA1-DA86-4ba1-AF1A-B9A4482B7A74}"
	set bp = Repository.GetElementByGuid(bpGUID) 'Get the business process
	
	dim bpd as EA.Diagram
	set bpd = bp.CompositeDiagram() 'Get the business process diagram
	Session.Output "DiagramID: " & bpd.DiagramID
	getDiagramObjects(bpd.DiagramID) 'Get the objects on that business process diagram
end sub

sub getDiagramObjects(DiagramID)
	dim sqlGetDOs
	dim element as EA.Element
	dim DOs
	set DOs = CreateObject("System.Collections.ArrayList")
	dim StartEvents, EndEvents, Activities, IntermediateEvents, Gateways, Lanes, Pools, Messages, BAM_Specifications 'Stereotypes
	set StartEvents = CreateObject("System.Collections.ArrayList")
	set EndEvents = CreateObject("System.Collections.ArrayList")
	set Activities = CreateObject("System.Collections.ArrayList")
	set IntermediateEvents = CreateObject("System.Collections.ArrayList")
	set Gateways = CreateObject("System.Collections.ArrayList")
	set Lanes = CreateObject("System.Collections.ArrayList")
	set Pools = CreateObject("System.Collections.ArrayList")
	set Messages = CreateObject("System.Collections.ArrayList")
	set BAM_Specifications = CreateObject("System.Collections.ArrayList")
	
	sqlGetDOs = "select o.object_id " & _
				"from t_diagramobjects do " & _
				"inner join t_object o " & _		
				"on do.object_id = o.object_id " & _
				"where do.diagram_id = '" & DiagramID & "' " 
	
	set DOs = getElementsFromQuery(sqlGetDOs)
				
	for each element in DOs
	
		Select Case element.Stereotype
			Case "IntermediateEvent"
				IntermediateEvents.add(element)	 
			Case "Gateway"
				Gateways.add(element)
			Case "Lane"
				Lanes.add(element)
			Case "Pool"
				Pools.add(element)
			Case "Message"
				Messages.add(element)
			Case "BAM_Specification"
				BAM_Specifications.add(element)
			Case "Activity"
				Activities.add(element)
			Case "StartEvent"
				StartEvents.add(element)
			Case "EndEvent"
				EndEvents.add(element)
		End Select
	next
	
	Session.Output "Pools: " & pools.count()
	Session.Output "Lanes: " & lanes.count()
	Session.Output "StartEvents: " & startevents.count()
	Session.Output "IntermediateEvents: " & intermediateevents.count()
	Session.Output "EndEvents: " & endevents.count()
	Session.Output "Activities: " & activities.count()
	Session.Output "Messages: " & messages.count()
	Session.Output "Gateways: " & gateways.count()
	Session.Output "BAM_Specifications: " & BAM_specifications.count()
	
	'CallActivityRef --> linkt naar activiteit, sub activiteit in geval van subproces
	

end sub

main