'[path=\Projects\Project K\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include


'
' Script Name: Import database specification
' Author: Geert Bellekens
' Purpose: Import database specification
' Date: 2019-04-26

const outPutName = "Import Database Specification"

sub main
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting Import Database Specification", 0
		'get selected package
		dim selectedPackage as EA.Package
		set selectedPackage = Repository.GetTreeSelectedPackage()
		'exit if not selected
		if selectedPackage is nothing then
			msgbox "Please select a package before running this script"
			exit sub
		end if
		'start import
		importDatabaseSpec(selectedPackage)
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Finished Import Database Specification", 0
end sub

function importDatabaseSpec(package)
	dim importFile
	set importFile = new TextFile
	if importFile.UserSelect("","XML Files (*.xml)|*.xml") then
		dim xmlDOM 
		set xmlDOM = CreateObject("MSXML2.DOMDocument")
		If xmlDOM.LoadXML(importFile.Contents) Then
			importDatabaseFromXmlDoc package, xmlDOM
		else
			'error loading xml file
			Repository.WriteOutput outPutName, now() & " Error loading xmlFile " & importFile.FullPath, 0
		end if
	end if	
end function

function importDatabaseFromXmlDoc(package, xmlDOM)
	'get table nodes
	dim tableNodes
	set tableNodes = xmlDOM.SelectNodes("//Table")
	dim tableNode
	for each tableNode in tableNodes
		'get name
		dim nameNode
		set nameNode = tableNode.SelectSingleNode("./Name")
		'create table
		dim table as EA.Element
		set table = createTable(package, nameNode.Text)
		'get columns
		dim columnNodes
		set columnNodes = tableNode.SelectNodes("./Column")
		dim columnNode
		for each columnNode in columnNodes
			'process columnNode
			processColumnNode table, columnNode
		next
	next
end function

function processColumnNode(table, columnNode)
	'get name
	dim nameNode
	set nameNode = columnNode.SelectSingleNode("./Name")
	'get type
	dim typeNode
	set typeNode = columnNode.SelectSingleNode("./Type")
	'get notnull
	dim notNullNode
	set notNullNode = columnNode.SelectSingleNode("./NotNull")
	'create column
	dim column as EA.Attribute
	dim tempColumn as EA.Attribute
	set column = nothing
	'check if exists
	for each tempColumn in table.Attributes
		if tempColumn.Name = nameNode.Text _
		  and lcase(tempColumn.Stereotype) = "column" then
			set column = tempColumn
			exit for
		end if
	next
	'create if not exists
	if column is nothing then
		set column = table.Attributes.AddNew(nameNode.Text, typeNode.Text)
		column.StereotypeEx = "EAUML::Column"
	end if
	if lcase(notNullNode.Text) = "true" then
		column.AllowDuplicates = true
	end if
	column.Type = typeNode.Text
	column.Update
end function

function createTable(package, name)
	dim tempTable as EA.Element
	dim table
	set table = nothing 'initialize
	for each tempTable in package.elements
		if tempTable.Name = name _
			and lcase(tempTable.Type) = "class" _
			and lcase(tempTable.Stereotype) = "table" then
			set table = tempTable
			exit for
		end if
	next
	if table is nothing then
		set table = package.Elements.AddNew(name, "EAUML::Table")
	end if
	table.Gentype = "Oracle" 'set database type
	table.update
	'return
	set createTable = table
end function

main