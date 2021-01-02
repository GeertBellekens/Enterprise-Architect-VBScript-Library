'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

sub main
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	'create diagram
	dim diagram as EA.Diagram
	set diagram = selectedPackage.Diagrams.AddNew("bugdemo","Logical")
	diagram.Update()
	'create classes
	dim class1 as EA.Element
	dim class2 as EA.Element
	set class1 = selectedPackage.Elements.AddNew("Class1","Class")
	set class2 = selectedPackage.Elements.AddNew("Class2","Class")
	'create associations
	dim goodAssociation as EA.Connector
	dim badAssociation as EA.Connector
	set goodAssociation = class1.Connectors.AddNew("goodAssociation", "Association")
	set badAssociation = class1.Connectors.AddNew("goodAssociation", "Association")
	'set the other side
	goodAssociation.SupplierID = class2.ElementID
	badAssociation.SupplierID = class2.ElementID
	'manipulate association ends good order
	goodAssociation.ClientEnd.Role = "partRole"
	'composite end last
	goodAssociation.ClientEnd.Role = "compositeRole"
	goodAssociation.ClientEnd.Aggregation = 2 'composite
	'save the association
	goodAssociation.Update
	
	'manipulate association ends reverse order
	'composite end first
	badAssociation.ClientEnd.Role = "compositeRole"
	badAssociation.ClientEnd.Aggregation = 2 'composite
	'part end last
	badAssociation.ClientEnd.Role = "partRole"
	'save the association	
	badAssociation.Update
	
	'add elements to diagram
	dim class1Do as EA.DiagramObject
	dim class2Do as EA.DiagramObject
	set class1Do = diagram.DiagramObjects.AddNew("l=10;r=70;t=10;b=50;","")
	set class2Do = diagram.DiagramObjects.AddNew("l=100;r=170;t=10;b=50;","")
	class1Do.ElementID = class1.ElementID
	class1Do.Update
	class2Do.ElementID = class2.ElementID
	class2Do.Update
	
	'layout diagram (which will show the diagram as well)
	dim diagramGUIDXml
	'The project interface needs GUID's in XML format, so we need to convert first.
	diagramGUIDXml = Repository.GetProjectInterface().GUIDtoXML(diagram.DiagramGUID)
	'Then call the layout operation
	Repository.GetProjectInterface().LayoutDiagramEx diagramGUIDXml, lsDiagramDefault, 4, 20 , 20, false
	'diagram.Update
end sub

main