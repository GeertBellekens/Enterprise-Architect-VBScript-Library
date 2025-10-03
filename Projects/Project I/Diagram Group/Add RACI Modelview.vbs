'[path=\Projects\Project I\Diagram Group]
'[group=Diagram Group]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Add RACI Modelview
' Author: Geert Bellekens
' Purpose: Add a modelview for the RACI of this business process
' Date: 2023-04-27
'

const outPutName = "Add RACI ModelView"
'
' Diagram Script main function
'
sub OnDiagramScript()

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	
	'exit if no current diagram found
	if currentDiagram is nothing then
		exit sub
	end if
	
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'inform user
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName , 0
	
	'first save the diagram
	Repository.SaveDiagram currentDiagram.DiagramID
	
	'add the RACI
	addRACIModelView currentDiagram
	
	'reload diagram to show changes
	Repository.ReloadDiagram currentDiagram.DiagramID
	
	'inform user
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0
	
end sub

function addRACIModelView(currentDiagram)
	'only works on a diagram that is owned by an element
	if not currentDiagram.ParentID > 0 then
		Repository.WriteOutput outPutName, now() & " ERROR: This script only works on diagrams that are owned by an element" , 0
		exit function
	end if
	'get business process
	dim bp as EA.Element
	set bp = Repository.GetElementByID(currentDiagram.ParentID)
	if not bp is nothing then
		'add modelView element
		dim modelView as EA.Element
		set modelView = bp.Elements.AddNew("RACI", "EAUML::ModelView")
		modelView.Update
		'fill the tagged value with the SQL
		dim tv as EA.TaggedValue
		for each tv in modelView.TaggedValues
			if tv.Name = "ViewProperties" then
				tv.Value = "<memo>"
				tv.Notes = getViewSQL(bp)
				tv.Update
			end if
		next
		'put modelview element on the diagram
		dim diagramObject as EA.DiagramObject
		set diagramObject = currentDiagram.DiagramObjects.AddNew( "l=200;r=1000;t=1200;b=2000;", "" )
		diagramObject.ElementID = modelView.ElementID
		diagramObject.Update
	end if
end function

function getViewSQL(bp)
	dim viewSQL
	viewSQL = "<modelview>" & vbNewLine & _
	"	<source customSQL=""select bpa.ea_guid as CLASSGUID, bpa.Object_Type as CLASSTYPE, bpa.name as Activity, &#xA;r.RoleName&#xA;, CASE WHEN r.RASCI = 'R' THEN r.RASCI ELSE NULL END as R&#xA;, CASE WHEN r.RASCI = 'A' THEN r.RASCI ELSE NULL END as A&#xA;, CASE WHEN r.RASCI = 'C' THEN r.RASCI ELSE NULL END as C&#xA;, CASE WHEN r.RASCI = 'I' THEN r.RASCI ELSE NULL END as I&#xA;from t_object bp &#xA;left join t_object pl on pl.ParentID = bp.Object_ID&#xA;      and pl.Stereotype = 'Pool'&#xA;left join t_object ln on ln.ParentID in (pl.Object_ID, bp.Object_ID)&#xA;      and ln.Stereotype = 'Lane'&#xA;inner join t_object bpa on bpa.ParentID in (ln.Object_ID, pl.Object_ID, bp.Object_ID)&#xA;      and bpa.Stereotype = 'Activity'&#xA;left join t_objectproperties tv on tv.Object_ID = bpa.Object_ID&#xA;     and tv.Property = 'calledActivityRef'&#xA;left join t_object ca on ca.ea_guid = tv.Value&#xA;left join &#xA;  (select c.Start_Object_ID, ro.Name as RoleName, tv.VALUE as RASCI from t_connector c&#xA;  inner join t_object ro on ro.Object_ID = c.End_Object_ID&#xA;        and ro.Stereotype = 'Archimate_BusinessRole'&#xA;  left join t_connectortag tv on tv.ElementID = c.Connector_ID&#xA;      and tv.Property = 'RACI'&#xA;  where c.Stereotype = 'trace') r on r.Start_Object_ID = bpa.Object_ID&#xA;inner join t_diagram d on d.ParentID = bp.Object_ID&#xA;inner join t_diagramobjects do on do.Diagram_ID = d.Diagram_ID &#xA;      and do.Object_ID = bpa.Object_ID&#xA;where bp.ea_guid = 'GUID_TO_REPLACE'&#xA;order by do.RectLeft, do.RectTop desc, bpa.Name, CASE WHEN r.RASCI = 'R' THEN 1&#xA;      WHEN r.RASCI = 'A' THEN 2&#xA;      WHEN r.RASCI = 'S' THEN 3&#xA;      WHEN r.RASCI = 'C' THEN 4&#xA;      WHEN r.RASCI = 'I' THEN 5 END""/>" & vbNewLine & _
	"</modelview>"
	'replace GUID_TO_REPLACE by the actual guid of the Business Process
	viewSQL = replace(viewSQL, "GUID_TO_REPLACE", bp.ElementGUID)
	'return
	getViewSQL = viewSQL
end function
'
'function test
'	dim diagram as EA.Diagram
'	set diagram = Repository.GetDiagramByGuid("{BA415D6C-A41D-4d33-B96A-00332C6E9C1F}")
'	addRACIModelView(diagram)
'end function
'
'test
OnDiagramScript
