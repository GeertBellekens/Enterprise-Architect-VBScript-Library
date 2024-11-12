'[path=\Framework\Tools\UML Profile]
'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate UML profiles
' Author: Geert Bellekens
' Purpose: Generate the SAP UML profiles and then the KUL Modelling MDG
' Date: 2020-05-29
'
const outPutName = "Generate EDSN Modelling MDG"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Starting Generate EDSN Modelling MDG", 0
 'do the actual work
 'UML profile
 Repository.WriteOutput outPutName, now() & " Generating EDSN UML profile", 0
 Repository.SavePackageAsUMLProfile "{3442BECF-27B1-4c6d-83B6-B65EC0DB184C}", ""
 'diagram profile
 Repository.WriteOutput outPutName, now() & " Generating EDSN diagram profile", 0
 Repository.SavePackageAsUMLProfile "{01050A5F-5FC5-4f6c-B20F-240B403856AB}", ""
 'SAP toolbox profile
 Repository.WriteOutput outPutName, now() & " Generating EDSN toolbox profiles", 0
 Repository.SaveDiagramAsUMLProfile "{A1A867E6-222E-4ba9-9F36-55D1EEE48986}", "" 'Applications
 'MDG file
 Repository.WriteOutput outPutName, now() & " Generating MDG file", 0
 Repository.GenerateMDGTechnology "G:\My Drive\Klanten Bellekens IT\EDSN\MDG\MDG Files\EDSN MDG Technology.mts"
 'import technology (doesn't seem to work)
' dim mdgfile
' set mdgfile = new TextFile
' mdgFile.FullPath = "G:\My Drive\Klanten Bellekens IT\KUL\MDG\KUL Modelling MDG.xml"
' mdgfile.LoadContents
' dim mdgstring
' mdgstring = mdgFile.Contents
' dim success
' success = Repository.ImportTechnology(mdgFile.contents)
' if not success then
'  Repository.WriteOutput outPutName, now() & "ERROR: importing KUL Modelling MDG" , 0 
' end if
 'set timestamp
 Repository.WriteOutput outPutName, now() & " Finished Generate EDSN Modelling MDG" , 0 
end sub



main