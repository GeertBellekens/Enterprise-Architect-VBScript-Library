'[path=\Framework\Tools\UML Profile]
'[group=UML Profile]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Generate ODCS MDG Technology
' Author: Geert Bellekens
' Purpose: Generate the ODCS UML profile and MDG technology
' Date: 2025-10-14
'
const outPutName = "Generate ODCS MDG"

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
	Repository.SavePackageAsUMLProfile "{A2127FD1-9E16-701D-BD39-FF80B8C1D8DE}", ""
	'diagram profile
	Repository.WriteOutput outPutName, now() & " Generating diagram profile", 0
	Repository.SavePackageAsUMLProfile "{9B5727A2-DF5C-DDFC-A260-FFA707967B50}", ""
	'SAP toolbox profile
	Repository.WriteOutput outPutName, now() & " Generating toolbox profiles", 0
	Repository.SaveDiagramAsUMLProfile "{A1CF8B62-FC73-9C58-8B25-B3D0BB1BF25A}", "" 
	'MDG file
	Repository.WriteOutput outPutName, now() & " Generating MDG file", 0
	Repository.GenerateMDGTechnology "C:\Users\geert\Documents\GitHub\Enterprise-Architect-Toolpack\EADataContract\MDG technology\MDG Files\ODCS MDG technology selection.mts"
 
 	'export xmi
	Repository.WriteOutput outPutName, now() & " Exporting XMI file", 0
	dim project as EA.Project
	set project = Repository.GetProjectInterface()
	dim mdgPackageGUID
	mdgPackageGUID = project.GUIDtoXML("{A936878D-C168-4ADA-BEBF-F5DB91F44A96}")
	project.ExportPackageXMI mdgPackageGUID, xmiEADefault, 1, -1, 0, 0, "C:\Users\geert\Documents\GitHub\Enterprise-Architect-Toolpack\EADataContract\MDG technology\XMI Exports\ODCS MDG Technology.xmi"

	'export shapescripts
	Repository.WriteOutput outPutName, now() & " Exporting shapescripts", 0
	exportShapescriptsFromMDGFile "C:\Users\geert\Documents\GitHub\Enterprise-Architect-Toolpack\EADataContract\MDG technology\MDG Files\ODCS MDG Technology.xml", "C:\Users\geert\Documents\GitHub\Enterprise-Architect-Toolpack\EADataContract\MDG technology\Shapescripts"

	'set timestamp
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName , 0 
end sub



main