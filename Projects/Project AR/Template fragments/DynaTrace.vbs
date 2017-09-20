'[path=\Projects\Project AR\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
function MyRtfData (ContainerID)

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
	
	dim xmlRow
	set xmlRow = xmlDOM.createElement( "Row" )
	xmlData.appendChild xmlRow
	
	'get the dynatrace value
	dim dynaTraceValue
	dynaTraceValue = getDynaTraceValue(containerID)
	
	'add dynatrace value to xml
	dim xmlDynaTrace
	set xmlDynaTrace = xmlDOM.createElement( "DynaTrace" )	
	xmlDynaTrace.text = dynaTraceValue
	xmlRow.appendChild xmlDynaTrace
	
	'return the xml
	MyRtfData = xmlDOM.xml
end function

function getDynaTraceValue(containerID)
	dim dynaTraceValue
	'default = "Nee"
	dynaTraceValue = "Nee"
	dim container as EA.Element
	set container = Repository.GetElementByID(containerID)
	if not container is nothing then
		'get the Dynatrace elements
		dim sqlGetDynaTrace
		sqlGetDynaTrace = 	"select dt.[Object_ID]                                                          " & _
							" from (((((([t_object] cnt                                                     " & _
							" inner join [t_diagramobjects] cntdo on cntdo.[Object_ID] = cnt.[Object_ID])   " & _
							" inner join t_connector c on c.[Start_Object_ID] = cnt.[Object_ID])            " & _
							" inner join [t_object] srv on (c.[End_Object_ID] = srv.[Object_ID]             " & _
							"                              and srv.[Stereotype] = 'ArchiMate_Node'))        " & _
							" inner join [t_diagramobjects] srvdo on srvdo.[Object_ID] = srv.[Object_ID])   " & _
							" inner join [t_diagramobjects] dtdo on dtdo.[Diagram_ID] = srvdo.[Diagram_ID]) " & _
							" inner join [t_object] dt on (dt.[Object_ID] = dtdo.[Object_ID]                " & _
							"                             and dt.[Object_Type] = 'Action'))                 " & _
							" inner join [t_object] dto on (dt.[Classifier] =  dto.[Object_ID]              " & _
							"                              and dto.[Name] = 'Dynatrace')                    " & _
							" where cnt.[Object_ID] = "& containerID &"                                     " & _
							" and srvdo.[Diagram_ID] = cntdo.[Diagram_ID]                                   " & _
							" and dtdo.[RectTop] <= srvdo.[RectTop]                                         " & _
							" and dtdo.[RectLeft] >= srvdo.[RectLeft]                                       " & _
							" and dtdo.[RectRight] <= srvdo.[RectRight]                                     " & _
							" and dtdo.[RectBottom] >= srvdo.[RectBottom]                                   "
		dim dynatraceElements 
		set dynatraceElements = getElementsFromQuery(sqlGetDynaTrace)
		if dynatraceElements.Count > 0 then
			dynaTraceValue = "Ja"
		end if
	end if
	'return the value
	getDynaTraceValue = dynaTraceValue
end function


'msgbox MyRtfData(4864)