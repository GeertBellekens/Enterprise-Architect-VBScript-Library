'[path=\Projects\Project M\Template fragments]
'[group=Template fragments]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: WPActivitiesTranslated
' Author: Geert Bellekens
' Purpose: Returns the data needed for the template fragment WP Activities. 
' It returns name and translated description for each activity on the diagram in the given package
' Date: 
'

function MyRtfData (packageID, tagname)
	
	dim xmlDOM 
	set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
	'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
	xmlDOM.validateOnParse = false
	xmlDOM.async = false
	 
	dim node 
	set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
    xmlDOM.appendChild node
'
	dim xmlRoot 
	set xmlRoot = xmlDOM.createElement( "EADATA" )
	xmlDOM.appendChild xmlRoot

	dim xmlDataSet
	set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
	xmlRoot.appendChild xmlDataSet
	 
	dim xmlData 
	set xmlData = xmlDOM.createElement( "Data" )
	xmlDataSet.appendChild xmlData
	
	'create the rows for each Activity on the diagram
	createRows xmlDOM, xmlData, packageID, tagname
	
	MyRtfData = xmlDOM.xml
end function

function createRows(xmlDOM, xmlData,packageID, tagname)
	'first get the activities we need
	dim activities
	dim sqlGetActivities
	sqlGetActivities = 	"select act.Object_ID from ((t_object act                         " & _
						" inner join t_diagramobjects do on do.Object_ID = act.Object_ID) " & _
						" inner join t_diagram d on d.Diagram_ID = do.Diagram_ID)         " & _
						" where d.Package_ID = " & packageID & "                          " & _
						" and act.Object_Type = 'Activity'                                " & _
						" and act.Stereotype = 'Activity'                                 " & _
						" order by do.RectLeft, do.RectTop                                "
	
	set activities = getElementsFromQuery(sqlGetActivities)
	'create row for each activity
	dim activity as EA.Element
	for each activity in activities
		dim xmlRow
		set xmlRow = xmlDOM.createElement( "Row" )
		xmlData.appendChild xmlRow
		
		'name
		dim xmlActivityName
		set xmlActivityName = xmlDOM.createElement( "ActivityName" )
		if tagname = "EN" then
			xmlActivityName.text = activity.Name
		else
			xmlActivityName.text = activity.Alias
		end if
		xmlRow.appendChild xmlActivityName
		
		'description
		dim formattedAttr 
		set formattedAttr = xmlDOM.createAttribute("formatted")
		formattedAttr.nodeValue="1"
			
		dim xmlDescription
		set xmlDescription = xmlDOM.createElement( "Description" )	

		xmlDescription.text = getTagContent(activity.Notes, tagname)
		xmlDescription.setAttributeNode(formattedAttr)
		xmlRow.appendChild xmlDescription
	next
	
	
end function

'msgbox MyPackageRtfData(3357,"")
function test
	dim outputString
	dim fileSystemObject
	dim outputFile
	
	outputString =  MyRtfData(598, "EN")
	
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	set outputFile = fileSystemObject.CreateTextFile( "c:\\temp\\NLFRtest.xml", true )
	outputFile.Write outputString
	outputFile.Close
end function 

test