'[path=\Projects\Project E\Documents]
'[group=Documents]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
'
' Script Name: Capability Stories document
' Author: Geert Bellekens
' Purpose: Generate the capability stories for the selected package
' Date: 2024-05-17
'


dim capabilityStoriesDocument
set capabilityStoriesDocument = new Document
capabilityStoriesDocument.Name = "Capability Stories"
capabilityStoriesDocument.Description = "A Word document that the Capability stories with details of the Information flows"
capabilityStoriesDocument.ValidQuery = "select top 1 d.Diagram_ID                                             " & vbNewLine & _
							" from t_diagram d                                                     " & vbNewLine & _
							" inner join t_diagramlinks dl on dl.DiagramID = d.Diagram_ID          " & vbNewLine & _
							" inner join t_connector c on c.Connector_ID = dl.ConnectorID          " & vbNewLine & _
							" 						and c.Connector_Type = 'InformationFlow'       " & vbNewLine & _
							" inner join t_object cap1 on cap1.Object_ID = c.Start_Object_ID       " & vbNewLine & _
							" 						and cap1.Stereotype = 'ArchiMate_Capability'   " & vbNewLine & _
							" inner join t_object cap2 on cap2.Object_ID = c.Start_Object_ID       " & vbNewLine & _
							" 						and cap2.Stereotype = 'ArchiMate_Capability'   " & vbNewLine & _
							" where dl.Hidden = 0                                                  " & vbNewLine & _
							" and d.Package_ID =  #PACKAGEID#                                      "

sub GenerateCapabilityStoriesDocument
	'get selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	if package is nothing then
		exit sub
	end if
	dim project as EA.Project
	set project  = Repository.GetProjectInterface()
	dim fileName
	fileName = getUserSelectedDocumentFileName()
	project.RunReport package.PackageGUID, "CAP_Story", fileName
	CreateObject("WScript.Shell").Run("""" & fileName & """")
end sub