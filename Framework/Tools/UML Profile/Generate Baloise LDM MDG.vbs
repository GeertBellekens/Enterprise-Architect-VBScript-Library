'[path=\Framework\Tools\UML Profile]
'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate UML profiles
' Author: Geert Bellekens
' Purpose: Generate the Baloise XML MDG profiles and then the MDG
' Date: 2021-05-30
'
const outPutName = "Generate Baloise XML MDG"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName, 0
	'do the actual work
	'UML profile
	Repository.WriteOutput outPutName, now() & " Generating UML profile", 0
	Repository.SavePackageAsUMLProfile "{1B25FCD1-CCE7-4857-AEEC-AC0CD54B92AB}", ""
	'diagram profile
	Repository.WriteOutput outPutName, now() & " Generating diagram profile", 0
	Repository.SavePackageAsUMLProfile "{2DD7CA2B-54E3-4505-B45D-00B807246129}", ""
	'toolbox profile
	Repository.WriteOutput outPutName, now() & " Generating toolbox profile", 0
	Repository.SaveDiagramAsUMLProfile "{C59F6412-ED52-4ab6-A912-3ED4FF39C79E}", "" 
	'MDG file
	Repository.WriteOutput outPutName, now() & " Generating MDG file", 0
	Repository.GenerateMDGTechnology "G:\My Drive\Klanten Bellekens IT\Baloise\Baloise LDM MDG\MDG Files\Baloise LDM MDG.mts"
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0	
end sub



main