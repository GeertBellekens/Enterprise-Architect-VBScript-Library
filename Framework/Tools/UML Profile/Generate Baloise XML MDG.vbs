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
	Repository.WriteOutput outPutName, now() & " Starting Generate KUL Modelling MDG", 0
	'do the actual work
	'UML profile
'	Repository.WriteOutput outPutName, now() & " Generating SAP UML profile", 0
'	Repository.SavePackageAsUMLProfile "{3568ECD9-954A-4c00-A553-9BE406375B27}", ""
	'diagram profile
	Repository.WriteOutput outPutName, now() & " Generating diagram profile", 0
	Repository.SavePackageAsUMLProfile "{B54B2477-6763-4456-8B2B-57298327B83A}", ""
	'toolbox profile
	Repository.WriteOutput outPutName, now() & " Generating toolbox profiles", 0
	Repository.SaveDiagramAsUMLProfile "{E0935CD8-6608-445a-9F44-AFB2D6441CBB}", "" 'XML toolbox
	'MDG file
	Repository.WriteOutput outPutName, now() & " Generating MDG file", 0
	Repository.GenerateMDGTechnology "G:\My Drive\Klanten Bellekens IT\Baloise\Baloise XML MDG\Baloise XML MDG.mts"
	'import technology (doesn't seem to work)
'	dim mdgfile
'	set mdgfile = new TextFile
'	mdgFile.FullPath = "G:\My Drive\Klanten Bellekens IT\KUL\MDG\KUL Modelling MDG.xml"
'	mdgfile.LoadContents
'	dim mdgstring
'	mdgstring = mdgFile.Contents
'	dim success
'	success = Repository.ImportTechnology(mdgFile.contents)
'	if not success then
'		Repository.WriteOutput outPutName, now() & "ERROR: importing KUL Modelling MDG" , 0	
'	end if
	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished Generate KUL Modelling MDG" , 0	
end sub



main