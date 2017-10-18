'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: APPRQ x UC x UI overview
' Author: Geert Bellekens
' Purpose: Creates an Excel file containing
' Date: 2017-10-17
'
sub main
	'get the output
	dim body
	body = getOutPut()
	'create headers
	dim headers(0,7)
	headers(0,0) = "RequirementID"
	headers(0,1) = "Title"
	headers(0,2) = "Description"
	headers(0,3) = "Domain"
	headers(0,4) = "Acceptation Criteria"
	headers(0,5) = "Use Case"
	headers(0,6) = "Functional Design"
	headers(0,7) = "User Interface"
	'combine headers and content
	dim content
	content = mergeArrays(headers, body)
	'write the output to excel
	writeToExcel(content)
end sub

function writeToExcel(content)
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'create the tab
	excelOutput.createTab "APPRQ x UC x UI", content, true, "TableStyleMedium4"
	'save the excel file
	excelOutput.save
end function

function getOutPut()
	dim sqlGetContent
	sqlGetContent = " select apprq.RequirementID, apprq.Title, apprq.Description, apprq.Domain, apprq.AcceptationCriteria, " & _
					" uc.name as UseCase, d.Name as FunctionalDesign, null as UserInterface                                " & _
					" from                                                                                                 " & _
					" (select apprq.Object_ID, apprq.Name as RequirementID, tvTitle.Value as Title,                        " & _
					" apprq.Note as Description, tvDom.Value as Domain, tvAcc.Notes as AcceptationCriteria                 " & _
					" from t_object apprq                                                                                  " & _
					" left join t_objectproperties tvTitle on apprq.Object_ID = tvTitle.Object_ID                          " & _
					" 										and tvTitle.Property = 'Title'                                 " & _
					" left join t_objectproperties tvAcc on apprq.Object_ID = tvAcc.Object_ID                              " & _
					" 							and tvAcc.Property = 'Acceptance Criteria'                                 " & _
					" left join t_objectproperties tvDom on apprq.Object_ID = tvDom.Object_ID                              " & _
					" 							and tvDom.Property = 'Domain'  							                   " & _
					" where apprq.Stereotype = 'application requirement') apprq                                            " & _
					" inner join t_connector c on c.End_Object_ID = apprq.Object_ID                                        " & _
					" 							and c.Connector_Type = 'abstraction'                                       " & _
					" 							and c.Stereotype = 'trace'                                                 " & _
					" inner join t_object uc on uc.Object_ID = c.Start_Object_ID                                           " & _
					" 							and uc.Object_Type = 'UseCase'                                             " & _
					" inner join t_diagramobjects dob on dob.Object_ID = uc.Object_ID                                      " & _
					" inner join t_diagramObjects dobb on dobb.[Diagram_ID] = dob.[Diagram_ID]                             " & _
					" inner join t_object boundary on boundary.[Object_ID] = dobb.[Object_ID]                              " & _
					"                           and boundary.[Object_Type] = 'Boundary'                                    " & _
					" inner join t_diagram d on dob.Diagram_ID = d.Diagram_ID                                              " & _
					" where                                                                                                " & _
					" 1=1                                                                                                  " & _
					" and dob.[RectLeft] >= dobb.[RectLeft]                                                                " & _
					" and dob.[RectLeft] <= dobb.[RectRight]                                                               " & _
					" and dob.[RectTop]  <= dobb.[RectTop]                                                                 " & _
					" and dob.[RectTop]  >= dobb.[RectBottom]                                                              " & _
					" union all                                                                                            " & _
					" select apprq.RequirementID, apprq.Title, apprq.Description, apprq.Domain, apprq.AcceptationCriteria, " & _
					" null as UseCase, null as FunctionalDesign, apli.Name as UserInterface                                " & _
					" from                                                                                                 " & _
					" (select apprq.Object_ID, apprq.Name as RequirementID, tvTitle.Value as Title,                        " & _
					" apprq.Note as Description, tvDom.Value as Domain, tvAcc.Notes as AcceptationCriteria                 " & _
					" from t_object apprq                                                                                  " & _
					" left join t_objectproperties tvTitle on apprq.Object_ID = tvTitle.Object_ID                          " & _
					" 										and tvTitle.Property = 'Title'                                 " & _
					" left join t_objectproperties tvAcc on apprq.Object_ID = tvAcc.Object_ID                              " & _
					" 							and tvAcc.Property = 'Acceptance Criteria'                                 " & _
					" left join t_objectproperties tvDom on apprq.Object_ID = tvDom.Object_ID                              " & _
					" 							and tvDom.Property = 'Domain'  							                   " & _
					" where apprq.Stereotype = 'application requirement') apprq                                            " & _
					" inner join t_connector aplic on aplic.End_Object_ID = apprq.Object_ID                                " & _
					" inner join t_object apli on aplic.Start_Object_ID = apli.Object_ID                                   " & _
					" 							and apli.Object_Type = 'Interface'                                         " & _
					"                                                                                                      " & _
					" union all                                                                                            " & _
					" select apprq.RequirementID, apprq.Title, apprq.Description, apprq.Domain, apprq.AcceptationCriteria, " & _
					" uc.name as UseCase, d.Name as FunctionalDesign, apli.Name as UserInterface                           " & _
					" from                                                                                                 " & _
					" (select apprq.Object_ID, apprq.Name as RequirementID, tvTitle.Value as Title,                        " & _
					" apprq.Note as Description, tvDom.Value as Domain, tvAcc.Notes as AcceptationCriteria                 " & _
					" from t_object apprq                                                                                  " & _
					" left join t_objectproperties tvTitle on apprq.Object_ID = tvTitle.Object_ID                          " & _
					" 										and tvTitle.Property = 'Title'                                 " & _
					" left join t_objectproperties tvAcc on apprq.Object_ID = tvAcc.Object_ID                              " & _
					" 							and tvAcc.Property = 'Acceptance Criteria'                                 " & _
					" left join t_objectproperties tvDom on apprq.Object_ID = tvDom.Object_ID                              " & _
					" 							and tvDom.Property = 'Domain'  							                   " & _
					" where apprq.Stereotype = 'application requirement') apprq                                            " & _
					" inner join t_connector c on c.End_Object_ID = apprq.Object_ID                                        " & _
					" 							and c.Connector_Type = 'abstraction'                                       " & _
					" 							and c.Stereotype = 'trace'                                                 " & _
					" inner join t_object uc on uc.Object_ID = c.Start_Object_ID                                           " & _
					" 							and uc.Object_Type = 'UseCase'                                             " & _
					" inner join t_diagramobjects dob on dob.Object_ID = uc.Object_ID                                      " & _
					" inner join t_diagramObjects dobb on dobb.[Diagram_ID] = dob.[Diagram_ID]                             " & _
					" inner join t_object boundary on boundary.[Object_ID] = dobb.[Object_ID]                              " & _
					"                           and boundary.[Object_Type] = 'Boundary'                                    " & _
					" inner join t_diagram d on dob.Diagram_ID = d.Diagram_ID                                              " & _
					" inner join t_object act on act.ParentID = uc.Object_ID                                               " & _
					" 						   and act.Object_Type = 'Activity'                                            " & _
					" inner join t_object apli on apli.Object_Type = 'Interface'                                           " & _
					" 						and apli.Object_ID in                                                          " & _
					" 						(                                                                              " & _
					" 						select acui.End_Object_ID from t_connector acui                                " & _
					" 						inner join t_object ac on ac.ParentID = act.Object_ID                          " & _
					" 									and ac.Object_Type = 'Action'                                      " & _
					" 									and acui.Start_Object_ID = ac.Object_ID                            " & _
					" 						where acui.Connector_Type = 'abstraction'                                      " & _
					" 						and acui.Stereotype = 'trace'                                                  " & _
					" 						)                                                                              " & _
					" where                                                                                                " & _
					" 1=1                                                                                                  " & _
					" and dob.[RectLeft] >= dobb.[RectLeft]                                                                " & _
					" and dob.[RectLeft] <= dobb.[RectRight]                                                               " & _
					" and dob.[RectTop]  <= dobb.[RectTop]                                                                 " & _
					" and dob.[RectTop]  >= dobb.[RectBottom]                                                              " & _
					" order by RequirementID,UseCase,FunctionalDesign, UserInterface                                       "
	dim outputArray
	outputArray = getArrayFromQuery(sqlGetContent)
	getOutPut = formatOutput(outputArray)
end function


function formatOutput(outputArray)
	'The third column is in formatted text. We convert it to plain text
	dim i
	for i = 0 to Ubound(outputArray)
		outputArray(i,2) = Repository.GetFormatFromField("TXT",outputArray(i,2))
	next
	formatOutput = outputArray
end function

main