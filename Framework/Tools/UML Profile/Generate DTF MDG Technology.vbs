'[path=\Framework\Tools\UML Profile]
'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript


'
' Script Name: Generate UML profiles
' Author: Geert Bellekens
' Purpose: Generate the DTF UML profile and MDG technology
' Date: 2025-05-15
'
const outPutName = "Generate DTF MDG"

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
 Repository.SavePackageAsUMLProfile "{B4705A35-C75B-01E3-A3CF-32CBE1B4159B}", ""
 'diagram profile
 Repository.WriteOutput outPutName, now() & " Generating diagram profile", 0
 Repository.SavePackageAsUMLProfile "{BBDAFDD4-DADA-AA2B-9144-B81871B12E6E}", ""
 'SAP toolbox profile
 Repository.WriteOutput outPutName, now() & " Generating toolbox profiles", 0
 Repository.SaveDiagramAsUMLProfile "{9A4B2673-0E1C-8B32-B471-E77A39514EE8}", "" 
 'MDG file
 Repository.WriteOutput outPutName, now() & " Generating MDG file", 0
 Repository.GenerateMDGTechnology "G:\My Drive\Klanten Bellekens IT\Imagine 4D\MDG Technology\MDG Files\DTF MDG Technology.mts"

 'set timestamp
 Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0 
end sub



main