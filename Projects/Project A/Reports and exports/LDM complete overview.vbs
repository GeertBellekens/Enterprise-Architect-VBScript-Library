'[path=\Projects\Project A\Reports and exports]
'[group=Reports and exports]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: LDM Complete overview
' Author: Geert Bellekens
' Purpose: Create an excel output containing all details of the selected LDM model
' Date: 2019-09-11
'

const outPutName = "LDM complete overview"
const excelTemplate = "G:\Projects\80 Enterprise Architect\Output\Excel export templates\CMS - FD - XD - Logical Data Model template.xltx"


sub main
	'reset output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim package as EA.Element
	set package = Repository.GetTreeSelectedPackage
	if not package is nothing then
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting LDM Complete overview for '"& package.Name &"'", 0
		'actually export data
		generateLDMComplete(package)
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Finished LDM Complete overview for '"& package.Name &"'", 0
	end if
end sub

function generateLDMComplete(package)
	'inform user
	Repository.WriteOutput outPutName, now() & " Getting package tree", 0
	'get the package tree ID list
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	'inform user
	Repository.WriteOutput outPutName, now() & " Getting Classes and Attributes", 0
	'get classes and attributes info
	dim classesAndAttributes
	classesAndAttributes = getLDMClassesAndAttributes(packageTreeIDs)
	'inform user
	Repository.WriteOutput outPutName, now() & " Getting Relationships", 0
	'get Relationships info
	dim relationships
	relationships = getLDMRelationships(packageTreeIDs)
	'inform user
	Repository.WriteOutput outPutName, now() & " Getting Datatypes", 0
	'get Datatypes info 
	dim datatypeInfo
	datatypeInfo = getLDMDatatypes(packageTreeIDs)
	'inform user
	Repository.WriteOutput outPutName, now() & " Getting Enumerations", 0
	'get enumerations info
	dim enumerations
	enumerations = getLDMEnumerations(packageTreeIDs)
	'inform user
	Repository.WriteOutput outPutName, now() & " Generating Excel file", 0
	'write output to excel
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'open template
	excelOutput.NewFile excelTemplate
	'add the output to a sheet in excel
	dim sheet
	'create classes and attributes sheet
	set sheet = excelOutput.createTab ("Classes-Attributes", classesAndAttributes, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, getClassesAndAttributesHeaders().Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'create Relationships sheet
	set sheet = excelOutput.createTab("Relationships", relationships, true, "TableStyleMedium4")
	'set headers to atrias red
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, getRelationshipsHeaders().Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'create enumerations sheet
	set sheet = excelOutput.createTab("Enumerations", enumerations, true, "TableStyleMedium4")
	'set headers to atrias red
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, getEnumerationsHeaders().Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'create datatypes sheet
	set sheet = excelOutput.createTab("Datatypes", datatypeInfo, true, "TableStyleMedium4")
	'set headers to atrias red
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, getDatatypeHeaders().Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'save the excel file
	excelOutput.save
end function

function getLDMEnumerations(packageTreeIDs)
	dim sqlGetData
	sqlGetData = getLDMEnumerationsQuery(packageTreeIDs)
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	'format description
	dim row
	for each row in results
		row(3) = Repository.GetFormatFromField("TXT", row(3))
	next
	dim headers
	set headers = getEnumerationsHeaders()
	'add the headers to the results
	results.Insert 0, headers
	'return array of results
	getLDMEnumerations = makeArrayFromArrayLists(results)
end function

function getLDMEnumerationsQuery(packageTreeIDs)
	dim sqlGetData
	sqlGetData = "SELECT c.name AS Enumeration, '''' + a.name AS LiteralValue, '''' + a.style AS Code   " & vbNewLine & _
				", a.Notes as Description                                                               " & vbNewLine & _
				" FROM ((t_attribute  a                                                                 " & vbNewLine & _
				" INNER JOIN t_object c ON a.object_id = c.object_id)                                   " & vbNewLine & _
				" INNER JOIN t_package p ON c.package_id = p.package_id)                                " & vbNewLine & _
				" WHERE c.Object_Type = 'Enumeration'                                                   " & vbNewLine & _
				" AND p.Package_ID IN (#Branch#)                                                        " & vbNewLine & _
				" ORDER BY 1,3,2                                                                        "
	'set the package tree id's
	sqlGetData = Replace(sqlGetData,"#Branch#",packageTreeIDs)
	'debug
'	dim textFile
'	set textFile = new TextFile
'	textFile.Contents = sqlGetData
'	textFile.FullPath = "H:\temp\debug.txt"
'	textFile.Save
	'return
	getLDMEnumerationsQuery = sqlGetData
end function

function getEnumerationsHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Enumeration")
	headers.add("Literal value")
	headers.add("Code")
	headers.add("Description")
	set getEnumerationsHeaders = headers
end function


function getLDMDatatypes(packageTreeIDs)
	dim sqlGetData
	sqlGetData = getLDMDatatypesQuery(packageTreeIDs)
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	dim headers
	set headers = getDatatypeHeaders()
	'add the headers to the results
	results.Insert 0, headers
	'return array of results
	getLDMDatatypes = makeArrayFromArrayLists(results)
end function

function getLDMDatatypesQuery(packageTreeIDs)
	dim sqlGetData
	sqlGetData = "SELECT dt.Name AS Datatype, rb.Name AS Inheritsfrom,                                                    " & vbNewLine & _
				" fractionDigits.Value AS fractionDigits, l.Value AS length, maxExclusive.Value AS maxExclusive,         " & vbNewLine & _
				" maxInclusive.Value AS maxInclusive, ml.Value AS maxLength,                                             " & vbNewLine & _
				" minExclusive.Value AS minExclusive, minInclusive.Value AS minInclusive, minLength.Value AS minLength,  " & vbNewLine & _
				" pattern.Value AS pattern, totalDigits.Value AS totalDigits, whiteSpace.Value AS whiteSpace             " & vbNewLine & _
				" FROM t_object dt                                                                                       " & vbNewLine & _
				" LEFT JOIN t_connector g ON g.Start_Object_ID = dt.Object_ID                                            " & vbNewLine & _
				" 							AND g.connector_type = 'Generalization'                                      " & vbNewLine & _
				" LEFT JOIN t_object rb ON g.End_Object_ID = rb.Object_ID                                                " & vbNewLine & _
				" LEFT JOIN t_objectproperties fractionDigits ON fractionDigits.[Object_ID] = dt.[Object_ID]             " & vbNewLine & _
				" 											AND fractionDigits.[Property] = 'fractionDigits'             " & vbNewLine & _
				" LEFT JOIN t_objectproperties l ON l.[Object_ID] = dt.[Object_ID]                                       " & vbNewLine & _
				" 								AND l.[Property] = 'length'                                              " & vbNewLine & _
				" LEFT JOIN t_objectproperties maxExclusive ON maxExclusive.[Object_ID] = dt.[Object_ID]                 " & vbNewLine & _
				" 											AND maxExclusive.[Property] = 'maxExclusive'                 " & vbNewLine & _
				" LEFT JOIN t_objectproperties maxInclusive ON maxInclusive.[Object_ID] = dt.[Object_ID]                 " & vbNewLine & _
				" 											AND maxInclusive.[Property] = 'maxInclusive'                 " & vbNewLine & _
				" LEFT JOIN t_objectproperties ml ON ml.[Object_ID] = dt.[Object_ID]                                     " & vbNewLine & _
				" 											AND ml.[Property] = 'maxLength'                              " & vbNewLine & _
				" LEFT JOIN t_objectproperties minExclusive ON minExclusive.[Object_ID] = dt.[Object_ID]                 " & vbNewLine & _
				" 											AND minExclusive.[Property] = 'minExclusive'                 " & vbNewLine & _
				" LEFT JOIN t_objectproperties minInclusive ON minInclusive.[Object_ID] = dt.[Object_ID]                 " & vbNewLine & _
				" 											AND minInclusive.[Property] = 'minInclusive'                 " & vbNewLine & _
				" LEFT JOIN t_objectproperties minLength ON minLength.[Object_ID] = dt.[Object_ID]                       " & vbNewLine & _
				" 											AND minLength.[Property] = 'minLength'                       " & vbNewLine & _
				" LEFT JOIN t_objectproperties pattern ON pattern.[Object_ID] = dt.[Object_ID]                           " & vbNewLine & _
				" 											AND pattern.[Property] = 'pattern'                           " & vbNewLine & _
				" LEFT JOIN t_objectproperties totalDigits ON totalDigits.[Object_ID] = dt.[Object_ID]                   " & vbNewLine & _
				" 											AND totalDigits.[Property] = 'totalDigits'                   " & vbNewLine & _
				" LEFT JOIN t_objectproperties whiteSpace ON whiteSpace.[Object_ID] = dt.[Object_ID]                     " & vbNewLine & _
				" 											AND whiteSpace.[Property] = 'whiteSpace'                     " & vbNewLine & _
				" INNER JOIN t_package p ON dt.package_id = p.package_id                                                 " & vbNewLine & _
				" WHERE dt.Object_Type IN ('Datatype')                                                                   " & vbNewLine & _
				" AND p.Package_ID IN (#Branch#)                                                                         " & vbNewLine & _
				" order by p.Name, dt.Name                                                                               "
	'set the package tree id's
	sqlGetData = Replace(sqlGetData,"#Branch#",packageTreeIDs)
	'debug
'	dim textFile
'	set textFile = new TextFile
'	textFile.Contents = sqlGetData
'	textFile.FullPath = "H:\temp\debug.txt"
'	textFile.Save
	'return
	getLDMDatatypesQuery = sqlGetData
end function

function getDatatypeHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Datatype")
	headers.Add("Inherits from")
	headers.Add("fractionDigits")
	headers.Add("length")
	headers.add("maxExclusive")
	headers.Add("maxInclusive")
	headers.Add("maxLength")
	headers.Add("minExclusive")
	headers.Add("minInclusive")
	headers.Add("minLength")
	headers.Add("pattern")
	headers.Add("totalDigits")
	headers.Add("whiteSpace")
	set getDatatypeHeaders = headers
end function

function getLDMRelationships(packageTreeIDs)
	dim sqlGetData
	sqlGetData = getLDMRelationshipsQuery(packageTreeIDs)
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	dim headers
	set headers = getRelationshipsHeaders()
	'add the headers to the results
	results.Insert 0, headers
	'return array of results
	getLDMRelationships = makeArrayFromArrayLists(results)
end function

function getLDMRelationshipsQuery(packageTreeIDs)
	dim sqlGetData
	sqlGetData = "SELECT c.connector_type, c.sourcecard, s.name as sname,                           " & vbNewLine & _
				" c.name as cName, c.destcard , t.name as tName,                                    " & vbNewLine & _
				" timesliced.Value as T, versioned.VALUE as V                                       " & vbNewLine & _
				" FROM ((((( t_connector c                                                          " & vbNewLine & _     
				" INNER JOIN t_object s on (c.Start_Object_ID = s.Object_ID                         " & vbNewLine & _  
				"						and s.Object_Type = 'Class'))                               " & vbNewLine & _                     
				" INNER JOIN t_object t on (c.End_Object_ID = t.Object_ID                           " & vbNewLine & _                       
				"						and s.Object_Type = 'Class'))                               " & vbNewLine & _
				" INNER JOIN t_package p on s.package_id = p.package_id)                            " & vbNewLine & _
				" LEFT JOIN t_connectortag versioned ON (versioned.[ElementID] = c.[Connector_ID]   " & vbNewLine & _     
				" 										AND versioned.[Property] = 'Versioned'))    " & vbNewLine & _     
				" LEFT JOIN t_connectortag timesliced ON (timesliced.[ElementID] = c.[Connector_ID] " & vbNewLine & _     
				" 										AND timesliced.[Property] = 'Timesliced'))  " & vbNewLine & _     
				" WHERE c.connector_type in ('Association','Generalization')                        " & vbNewLine & _
				" AND (s.Package_ID in (#Branch#) or t.Package_ID in (#Branch#))                    " & vbNewLine & _          
				" ORDER BY 3,4,6                                                                    "
	'set the package tree id's
	sqlGetData = Replace(sqlGetData,"#Branch#",packageTreeIDs)
	'debug
'	dim textFile
'	set textFile = new TextFile
'	textFile.Contents = sqlGetData
'	textFile.FullPath = "H:\temp\debug.txt"
'	textFile.Save
	'return
	getLDMRelationshipsQuery = sqlGetData
end function

function getRelationshipsHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Type")
	headers.Add("SM")
	headers.Add("Source Class")
	headers.Add("Relation name")
	headers.add("TM")
	headers.Add("Target Class")
	headers.Add("T")
	headers.Add("V")
	set getRelationshipsHeaders = headers
end function

function getLDMClassesAndAttributes(packageTreeIDs)
	dim sqlGetData
	sqlGetData = getLDMClassesAndAttributesQuery(packageTreeIDs)
	dim results
	set results = getArrayListFromQuery(sqlGetData)
	'remove the first column (pos)
	dim row
	for each row in results
		row.RemoveAt 0
	next
	dim headers
	set headers = getClassesAndAttributesHeaders()
	'add the headers to the results
	results.Insert 0, headers
	'return array of results
	getLDMClassesAndAttributes = makeArrayFromArrayLists(results)
end function

function getLDMClassesAndAttributesQuery(packageTreeIDs)
	dim sqlGetData
	sqlGetData = "SELECT 0 as Pos, c.Object_Type AS Type,                                                                             " & vbNewLine & _
				"  c.Name AS [ClassName], NULL AS [AttributeName],                                                                    " & vbNewLine & _
				" CASE WHEN c.Abstract = '0' THEN 'no' ELSE 'yes' END AS [Abstract], shared.Value AS [Shared],                        " & vbNewLine & _
				" NULL AS [M], NULL AS [ID], ts.Value AS [T], v.Value AS [V], dc.Value AS [DC], NULL AS [Datatype],                   " & vbNewLine & _
				" p.Name AS [Package], p1.Name AS [Package_1]                                                                         " & vbNewLine & _
				" FROM (((((((t_object c                                                                                              " & vbNewLine & _
				" LEFT JOIN t_objectproperties shared ON (shared.Object_ID = c.Object_ID AND shared.Property = 'Shared'))             " & vbNewLine & _
				" LEFT JOIN t_objectproperties v ON (v.Object_ID = c.Object_ID AND v.Property = 'Versioned'))                         " & vbNewLine & _
				" LEFT JOIN t_objectproperties ts ON (ts.Object_ID = c.Object_ID AND ts.Property = 'Timesliced'))                     " & vbNewLine & _
				" LEFT JOIN t_objectproperties dc ON (dc.Object_ID = c.Object_ID AND dc.Property = 'Atrias::Data Classification'))    " & vbNewLine & _
				" LEFT JOIN t_object op ON op.Object_ID = c.ParentID)                                                                 " & vbNewLine & _
				" INNER JOIN t_package p ON c.Package_ID = p.Package_ID)                                                              " & vbNewLine & _
				" LEFT JOIN t_package p1 ON p1.Package_ID = p.Parent_ID)                                                              " & vbNewLine & _
				" WHERE c.Object_Type = 'Class'                                                                                       " & vbNewLine & _
				" AND p.Package_ID IN (#Branch#)                                                                                      " & vbNewLine & _
				" UNION                                                                                                               " & vbNewLine & _
				" SELECT a.Pos as pos,  'Attribute' AS Type,                                                                          " & vbNewLine & _
				"  c.Name AS [Classname], a.Name AS [Attributename],                                                                  " & vbNewLine & _
				" CASE WHEN c.Abstract = '0' THEN 'no' ELSE 'yes' END AS [Abstract], shared.Value AS [Shared],                        " & vbNewLine & _
				" a.LowerBound + '..' + a.UpperBound AS [M], CASE WHEN x.[Description] IS NULL THEN 'no' ELSE 'yes' END AS [ID],      " & vbNewLine & _
				" ts.Value AS [T], v.VALUE AS [V], dc.Value AS [DC], a.Type AS [Datatype],                                            " & vbNewLine & _
				" p.Name AS [Package], p1.Name AS [Package_1]                                                                         " & vbNewLine & _
				" FROM ((((((((t_attribute a                                                                                          " & vbNewLine & _
				" INNER JOIN t_object c ON a.Object_ID = c.Object_ID)                                                                 " & vbNewLine & _
				" LEFT JOIN t_objectproperties shared ON (shared.Object_ID = c.Object_ID AND shared.Property = 'Shared'))             " & vbNewLine & _
				" LEFT JOIN t_attributetag v ON (v.ElementID = a.ID AND v.Property = 'Versioned'))                                    " & vbNewLine & _
				" LEFT JOIN t_attributetag ts ON (ts.ElementID = a.ID AND ts.Property = 'Timesliced'))                                " & vbNewLine & _
				" LEFT JOIN t_attributetag dc ON (dc.ElementID = a.ID AND dc.Property = 'Atrias::Data Classification'))               " & vbNewLine & _
				" LEFT OUTER JOIN t_xref x ON (x.Client = a.ea_guid                                                                   " & vbNewLine & _
				" 	AND x.Type = 'attribute property'                                                                                 " & vbNewLine & _
				" 	AND x.Description LIKE '%@PROP=@NAME=isID@ENDNAME;@TYPE=Boolean@ENDTYPE;@VALU=1@ENDVALU;%'))                      " & vbNewLine & _
				" INNER JOIN t_package p ON c.Package_ID = p.Package_ID)                                                              " & vbNewLine & _
				" LEFT JOIN t_package p1 ON p1.Package_ID = p.Parent_ID)                                                              " & vbNewLine & _
				" WHERE c.Object_Type = 'Class'                                                                                       " & vbNewLine & _
				" AND p.Package_ID IN (#Branch#)                                                                                      " & vbNewLine & _
				" ORDER BY [ClassName], Pos, [AttributeName]                                                                          "
	'set the package tree id's
	sqlGetData = Replace(sqlGetData,"#Branch#",packageTreeIDs)
	'debug
'	dim textFile
'	set textFile = new TextFile
'	textFile.Contents = sqlGetData
'	textFile.FullPath = "H:\temp\debug.txt"
'	textFile.Save
	'return
	getLDMClassesAndAttributesQuery = sqlGetData
end function

function getClassesAndAttributesHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Type")
	headers.Add("Class name")
	headers.Add("Attribute name")
	headers.add("Abstract")
	headers.Add("Shared")
	headers.Add("M")
	headers.Add("ID")
	headers.Add("T")
	headers.Add("V")
	headers.Add("DC")
	headers.Add("Datatype")
	headers.Add("Package")
	headers.Add("Package + 1")
	set getClassesAndAttributesHeaders = headers
end function

main