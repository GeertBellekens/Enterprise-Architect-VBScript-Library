'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
' ---------------- Disabled because it doesn't work. Requirements to be reviewed with reporting modelling team -----------------
'option explicit
'
'!INC Local Scripts.EAConstants-VBScript
'
'' EA-Matic
'' Script Name: BIG-RepActions
'' Author: Matthias Van der Elst
'' Purpose: Automate some actions when creating a BI & Reporting FD
'' Date: 10/05/2017
'
'
'function EA_OnPostNewElement(Info)
'    dim elementID
'    elementID = Info.Get("ElementID")
'	dim Element 
'	set Element = Repository.GetElementByID(elementID)
'	
'	if Element.Stereotype = "REP_BI-Group" then
'		'Create a class diagram under the new created BI-Group
'		dim BIDiagram 
'		set BIDiagram = Element.Diagrams.GetByName(Element.Name)
'		dim classDiagram
'		set classDiagram = Element.Diagrams.AddNew("LDM-BI Entities Overview", "Class")
'		classDiagram.Orientation = "L" 'Still not working
'		classDiagram.Update()
'		Element.Diagrams.Refresh()
'		
'		'Copy the new created BI-Group to the BI-Group diagram
'		dim BIGroupObject 
'		set BIGroupObject = BIDiagram.DiagramObjects.AddNew("l=10;r=200;t=-20;b=-250", "")
'		BIGroupObject.ElementID = Element.ElementID
'		BIGroupObject.Update()
'	
'	elseif Element.Stereotype = "REP_Report" then	
'		'Copy the new created Report to the Report diagram
'		dim REPDiagram 
'		set REPDiagram = Element.Diagrams.GetByName(Element.Name)
'		dim REPObject 
'		set REPObject = REPDiagram.DiagramObjects.AddNew("l=10;r=500;t=-20;b=-200", "")
'		REPObject.ElementID = Element.ElementID
'		REPObject.Update()
'		
'		
'	end if
'	
'	
'
'	
'
'	
'end function