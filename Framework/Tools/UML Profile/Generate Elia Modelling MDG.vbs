'[path=\Framework\Tools\UML Profile]
'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: Generate Elia Modelling MDG
' Author: Geert Bellekens
' Purpose: Generate the Elia Modelling MDG profiles and MDG file
' Date: 2023-01-20
'
const outPutName = "Generate Elia Modelling MDG"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'report progress
	Repository.WriteOutput outPutName, now() & " Starting "& outPutName, 0
	'do the actual work
	'UML Profile
	Repository.WriteOutput outPutName, now() & " Generating UML profile", 0
	Repository.SavePackageAsUMLProfile "{2FFFB254-7166-463f-ACA7-03E56674C4FD}", ""
	'Diagram profile
	Repository.WriteOutput outPutName, now() & " Generating diagram profile", 0
	Repository.SavePackageAsUMLProfile "{D2892CF3-D2B0-40e1-B345-D3CA867EAD4F}", ""
	'Toolbox profile
	Repository.WriteOutput outPutName, now() & " Generating toolbox profiles", 0
	Repository.SaveDiagramAsUMLProfile "{72485AD5-A3CB-4f38-AA08-09A0E6EFEF7E}", "" 'capability
	Repository.SaveDiagramAsUMLProfile "{46C24E20-A777-49b5-B068-DC263D1A1D8D}", "" 'LBIM
	Repository.SaveDiagramAsUMLProfile "{C0B746D5-43CE-4641-9542-EEB7206C0206}", "" 'CBIM
	Repository.SaveDiagramAsUMLProfile "{ADE9273A-85E5-4451-9066-7C0B1210B0B6}", "" 'Application Architecture
	Repository.SaveDiagramAsUMLProfile "{1D0D1850-46BF-4e05-88CD-2EE6492E40AD}", "" 'APM
	Repository.SaveDiagramAsUMLProfile "{BEA35F6F-59AB-4207-B3CA-FBAD1272431E}", "" 'Business Architecture

	'MDG file
	Repository.WriteOutput outPutName, now() & " Generating MDG file", 0
	'determine version
	dim mtsFilePath
	mtsFilePath = "U:\VC\MDG\Elia Modelling MDG\MDG files\Elia Modelling.mts"
	'set version number
	setVersionNumberInMTS mtsFilePath
	'generate mdg file
	Repository.GenerateMDGTechnology mtsFilePath

	'export xmi
	Repository.WriteOutput outPutName, now() & " Exporting XMI file", 0
	dim project as EA.Project
	set project = Repository.GetProjectInterface()
	dim mdgPackageGUID
	mdgPackageGUID = project.GUIDtoXML("{6FFA2F90-C36C-4809-BEA1-FA00E6917925}")
	project.ExportPackageXMI mdgPackageGUID, xmiEADefault, 1, -1, 0, 0, "U:\VC\MDG\Elia Modelling MDG\XMI export\Elia Modelling MDG.xmi"

	'export shapescripts
	Repository.WriteOutput outPutName, now() & " Exporting shapescripts", 0
	exportShapescriptsFromMDGFile "U:\VC\MDG\Elia Modelling MDG\MDG files\Elia Modelling.xml", "U:\VC\MDG\Elia Modelling MDG\Shapescripts"

	'report progress
	Repository.WriteOutput outPutName, now() & " Finshed "& outPutName, 0
end sub

function setVersionNumberInMTS(mtsFilePath)
	Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
	xDoc.Load(mtsFilePath)
	dim technologyNode
	set technologyNode = xDoc.SelectSingleNode( "//Technology")
	dim mdgPackage as EA.Package
	set mdgPackage = Repository.GetPackageByGuid("{6FFA2F90-C36C-4809-BEA1-FA00E6917925}")
	'set version to that of the MDG package
	technologyNode.setAttribute "version", mdgPackage.Version
	'save
	xDoc.Save mtsFilePath
end function



main