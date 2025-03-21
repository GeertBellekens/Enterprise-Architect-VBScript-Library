'[path=\Framework\Tools\UML Profile]
'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate UML profiles
' Author: Geert Bellekens
' Purpose: Generate the JSON UML profile and MDG technology
' Date: 2025-01-24
'
const outPutName = "Generate JSON Modelling MDG"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Starting " & outPutName , 0
 'do the actual work
 'UML profile
 Repository.WriteOutput outPutName, now() & " Generating UML profile", 0
 Repository.SavePackageAsUMLProfile "{AB7625DC-1E9E-4f3e-A847-E7FD1431EF5C}", ""
 'diagram profile
 Repository.WriteOutput outPutName, now() & " Generating diagram profile", 0
 Repository.SavePackageAsUMLProfile "{867299E8-B050-4e56-BB0C-96CB0F5F2383}", ""
 'SAP toolbox profile
 Repository.WriteOutput outPutName, now() & " Generating toolbox profiles", 0
 Repository.SaveDiagramAsUMLProfile "{F44ABF85-64AB-4552-97B8-B37BDC515523}", "" 
 'MDG file
 Repository.WriteOutput outPutName, now() & " Generating MDG file", 0
 Repository.GenerateMDGTechnology "C:\Users\geert\Documents\GitHub\Enterprise-Architect-Toolpack\EAJSON\Files\JSON MDG.mts"

 'set timestamp
 Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0 
end sub



main