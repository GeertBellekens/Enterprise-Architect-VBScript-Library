'[path=\Projects\Experiment\ProjectBrowserGroup]
'[group=ProjectBrowserGroup]

option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name:
' Author:
' Purpose:
' Date:
'

'
' Project Browser Script main function
'
sub OnProjectBrowserScript()
 
dim theElement AS EA.Element
set theElement = Repository.GetTreeSelectedObject()
dim docName
docName = "C:\\temp\\doc2.docx"
msgbox("Adding document " & vbcrlf & docName & vbcrlf & "to element " & theElement.name)
theElement.ImportInternalDocumentArtifact(docname)
 
end sub

OnProjectBrowserScript