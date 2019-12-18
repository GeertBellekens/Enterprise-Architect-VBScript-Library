'[path=\Projects\Project A\Reports and exports]
'[group=Reports and exports]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: TDM-LDM mapping overview
' Author: Geert Bellekens
' Purpose: get an overview of the mapping between TDM and LDM, including the non mapped classes and attributes
' Date: 2019-04-24
'

const outputTabName = "TDM - LDM mapping overview"
const excelTemplate = "G:\Projects\80 Enterprise Architect\Output\Excel export templates\TDM - LDM mapping template.xltx"

'define colors for output
dim ldmLightColor
ldmLightColor = RGB(242,247,252)
dim ldmDarkColor
ldmDarkColor = RGB(221,235,247)
'dim tdmLightColor
'tdmLightColor = RGB(255,250,229)
'dim tdmDarkColor
'tdmDarkColor  = RGB(244,240,220)

function GenerateTDMLDMOverview(TDMPackageGUID, LDMPackageGUID)
	Repository.CreateOutputTab outputTabName
	Repository.ClearOutput outputTabName
	Repository.EnsureOutputVisible outputTabName
	'get source package
	dim sourcePackage as EA.Package
'	msgbox "Please select the source TDM package"
'	set sourcePackage = selectPackage()
	set sourcePackage = Repository.GetPackageByGuid(TDMPackageGUID)
	if sourcePackage is nothing then
		exit function 'exit if not selected
	end if
	'get target package
	dim targetpackage as EA.Package
'	msgbox "Please select the target LDM package"
'	set targetpackage = selectPackage()
	set targetpackage = Repository.GetPackageByGuid(LDMPackageGUID)
	if targetpackage is nothing then
		exit function 'exit if not selected
	end if
	Repository.WriteOutput outputTabName, now() & " Starting TDM - LDM mapping overview" ,0
	Repository.WriteOutput outputTabName, now() & " Getting mapping data" ,0
	'get mapping details
	dim mappingDetails
	set mappingDetails = getMappingDetails(sourcePackage, targetPackage)
	'post processing
	Repository.WriteOutput outputTabName, now() & " Processing mapping data" ,0
	prepareMappingData(mappingDetails)
	'create search output
	dim headers
	set headers = getHeaders()
	'add the headers to the results
	'create the output object
'	dim searchOutput
'	set searchOutput = new SearchResults
'	searchOutput.Name = "Mapping Details"
'	searchOutput.Fields = headers
'	'put the contents in the output
'	dim row
'	for each row in mappingDetails
'		'add row the the output
'		searchOutput.Results.Add row
'	next
'	'show the output
'	searchOutput.Show
	Repository.WriteOutput outputTabName, now() & " Exporting mapping data for Excel" ,0
	'export to excel
	mappingDetails.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'load the template
	excelOutput.NewFile excelTemplate
	Repository.WriteOutput outputTabName, now() & " Getting data into Array" ,0
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(mappingDetails)
	Repository.WriteOutput outputTabName, now() & " Creating Excel sheet" ,0
	'add the output to a sheet in excel
	dim outputSheet
	set outputSheet = excelOutput.createTab("Mapping details", excelContents, true, "TableStyleMedium4")
	Repository.WriteOutput outputTabName, now() & " Formatting Excel sheet" ,0
	'format the output
	formatOutput excelOutput, outputSheet, headers
	Repository.WriteOutput outputTabName, now() & " Saving Excel sheet" ,0
	'save the excel file
	excelOutput.save
	Repository.WriteOutput outputTabName, now() & " Finished TDM - LDM mapping overview" ,0
end function

function formatOutput(excelOutput, outputSheet, headers)
	'set all fields to top align
	excelOutput.setVerticalAlignment outputSheet.UsedRange, xlVAlignTop
	'set headers to atrias red
	dim headerRange
	set headerRange = outputSheet.Range(outputSheet.Cells(1,1), outputSheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'add formatting
	dim i
	for i = 2 to outputSheet.UsedRange.Rows.Count
		dim LDMRange
		set LDMRange = outputSheet.Range(outputSheet.Cells(i,1), outputSheet.Cells(i, 10))
		dim TDMRange
		set TDMRange = outputSheet.Range(outputSheet.Cells(i,11), outputSheet.Cells(i, 19))
		'alternating rows formatting
		if i mod 2 = 0 then
			excelOutput.formatRange LDMRange, ldmDarkColor, "default" , "default", "default", "default", "default"
			'excelOutput.formatRange TDMRange, tdmDarkColor, "default" , "default", "default", "default", "default"
		else
			excelOutput.formatRange LDMRange, ldmLightColor, "default" , "default", "default", "default", "default"
			'excelOutput.formatRange TDMRange, tdmLightColor, "default" , "default", "default", "default", "default"
		end if
	next
	'hide guid columns
	excelOutput.hideColumn outputSheet, 10 'LDM_GUID
	excelOutput.hideColumn outputSheet, 19 'TDM_GUID
end function

function prepareMappingData(mappingDetails)
	'loop mapping details
	dim mappingDetail
	dim i
	i = 0
	for each mappingDetail in mappingDetails
		i = i + 1 'up the counter
		dim mappingNotes
		mappingNotes = mappingDetail(mappingDetail.Count - 1)
		'read it as xml document
		Dim xDoc 
		Set xDoc = CreateObject( "Microsoft.XMLDOM" )
		If xDoc.LoadXML(mappingNotes) Then
			'get the mapping target path node
			dim mappingTargetPath
			set mappingTargetPath =  xDoc.SelectSingleNode("//mappingTargetPath")
			if not mappingTargetPath is nothing then
				'get the second to last guid
				dim mappingPathGuids
				mappingPathGuids = Split (mappingTargetPath.Text, ".")
				dim parentGUID
				parentGUID = mappingPathGuids(Ubound(mappingPathGuids) - 1)
				dim parentElement as EA.Element
				set parentElement  = Repository.GetElementByGuid(parentGUID)
				if not parentElement is nothing then
					
					'replace the Class field with name of the actual parent
					mappingDetail(11) = parentElement.Name
				end if
			end if
		end if
		'remove the mappingNotes from the arraylist
		mappingDetail.RemoveAt mappingDetail.Count - 1
		'format remarks to get rid of things such as &gt;
		mappingDetail(mappingDetail.Count - 1) = Repository.GetFormatFromField("TXT",mappingDetail(mappingDetail.Count - 1))
	next
end function

function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("Domain") '0
	headers.add("Class") '1
	headers.add("Attribute") '2
	headers.add("Data Type") '3
	headers.add("Timesliced") '4
	headers.add("Versioned") '5
	headers.add("Mandatory") '6
	headers.add("Facets") '7
	headers.add("LDM_Type") '8
	headers.add("LDM_GUID") '9
	headers.add("Database") '10
	headers.add("Table") '11
	headers.add("Column") '12
	headers.add("Data_Type") '13
	headers.add("PK") '14
	headers.add("NotNull") '15
	headers.add("Default") '16
	headers.add("TDM_Type") '17
	headers.add("TDM_GUID") '18
	headers.add("Is Empty Mapping") '19
	headers.add("Remark") '20
	set getHeaders = headers
end function


function getMappingDetails(sourcePackage, targetPackage)
	'get the pakage tree id's
	dim sourcePackageTreeIDs
	sourcePackageTreeIDs = getPackageTreeIDString(sourcePackage)
	dim targetPackageTreeIDs
	targetPackageTreeIDs = getPackageTreeIDString(targetPackage)
	dim sqlGetMappingDetails
	sqlGetMappingDetails =  "select                                                                                                                                                             " & vbNewLine & _
							"   ldm.[Domain], ldm.Entity as [Class], ldm.Attribute as Attribute, ldm.AttributeType as AttributeType                                                             " & vbNewLine & _
							" , ldm.TimeSliced, ldm.Versioned, ldm.Mandatory, LEFT(ldm.facets,LEN(ldm.facets)-PATINDEX('%[^'+CHAR(13)+CHAR(10)+']%',REVERSE(ldm.facets))+1) as Facets           " & vbNewLine & _
							" , ldm.Type as LDMType, ldm.CLASSGUID as [Guid]                                                                                                                    " & vbNewLine & _
							" , tdm.[Database], tdm.Entity as [Table], tdm.Attribute as [Column]                                                                                                " & vbNewLine & _
							" , tdm.Datatype, tdm.PK, tdm.NotNull                                                                                                                               " & vbNewLine & _
							" , case when left(convert( varchar(max),tdm.[Default]), 1) = '''' then '''' else '' end + convert(varchar(max),tdm.[Default]) as [Default]                         " & vbNewLine & _
							" , tdm.Type as TDMType, tdm.CLASSGUID as [Guid]                                                                                                                    " & vbNewLine & _
							" , case when p.Package_ID is null then 'false' else 'true' END as EmptyMapping                                                                                     " & vbNewLine & _
							" , '''' + tdm.Comment as Comment, tdm.MappingNotes                                                                                                                                   " & vbNewLine & _
							" from (                                                                                                                                                            " & vbNewLine & _
							" select o.ea_guid AS CLASSGUID, o.Object_Type AS CLASSTYPE, o.name as Entity, null as Attribute, 'Table' as Type, tv.Value as LDMGuid, p1.Name as [Database]       " & vbNewLine & _
							" , CASE WHEN CHARINDEX('<description>',tv.NOTES) > 0 THEN                                                                                                          " & vbNewLine & _
							"   SUBSTRING(tv.NOTES,CHARINDEX('<description>',tv.NOTES)+LEN('<description>'),                                                                                    " & vbNewLine & _
							"              CHARINDEX('</description>',tv.NOTES)-CHARINDEX('<description>',tv.NOTES) - LEN('<description>')) END as Comment                                      " & vbNewLine & _
							", null as Datatype, null as PK, null as NotNull, null as [Default] , tv.Notes as MappingNotes                                                                      " & vbNewLine & _
							" from t_object o                                                                                                                                                   " & vbNewLine & _
							" left join t_objectproperties tv on tv.Object_ID = o.Object_ID                                                                                                     " & vbNewLine & _
							"                                                       and tv.Property in ('sourceElement','linkedAssociation', 'linkedAttribute')                                 " & vbNewLine & _
							" inner join t_package p on p.Package_ID = o.Package_ID                                                                                                             " & vbNewLine & _
							" left join t_package p1 on p1.Package_ID = p.Parent_ID                                                                                                             " & vbNewLine & _
							" where 1 = 1                                                                                                                                                       " & vbNewLine & _
							" and o.Object_Type = 'Class'                                                                                                                                       " & vbNewLine & _
							" and o.Stereotype = 'table'                                                                                                                                        " & vbNewLine & _
							" and o.Package_ID in (" & sourcePackageTreeIDs  & ")                                                                                                               " & vbNewLine & _
							" union all                                                                                                                                                         " & vbNewLine & _
							" select a.ea_guid AS CLASSGUID, 'Attribute' AS CLASSTYPE, o.name as Owner, a.Name as Attribute, 'Column' as Type, tv.VALUE as LDMGuid, p1.Name as [Database]       " & vbNewLine & _
							" , CASE WHEN CHARINDEX('<description>',tv.NOTES) > 0 THEN                                                                                                          " & vbNewLine & _
							"   SUBSTRING(tv.NOTES,CHARINDEX('<description>',tv.NOTES)+LEN('<description>'),                                                                                    " & vbNewLine & _
							"              CHARINDEX('</description>',tv.NOTES)-CHARINDEX('<description>',tv.NOTES) - LEN('<description>')) END as Comment                                      " & vbNewLine & _
							", a.type  + CASE WHEN dt.Size = 1 THEN '(' + cast(a.Length as varchar) + ')'                                                                                       " & vbNewLine & _
							"WHEN dt.Size = 2 THEN '(' + cast(a.Precision as varchar) + ', ' + cast(a.Scale as varchar)  + ')'                                                                  " & vbNewLine & _
							" ELSE '' END as Datatype                                                                                                                                           " & vbNewLine & _
							",a.IsOrdered as PK, a.AllowDuplicates as NotNull, a.[Default]  , tv.Notes as MappingNotes                                                                          " & vbNewLine & _
							" from t_attribute a                                                                                                                                                " & vbNewLine & _
							" inner join t_object o on o.Object_ID = a.Object_ID                                                                                                                " & vbNewLine & _
							"left join t_datatypes dt on dt.ProductName = o.GenType                                                                                                             " & vbNewLine & _
							"                                         and dt.DataType = a.Type                                                                                                  " & vbNewLine & _
							"                                         and dt.Type = 'DDL'                                                                                                       " & vbNewLine & _
							" left join t_attributetag tv on tv.ElementID = a.ID                                                                                                                " & vbNewLine & _
							"                                                       and tv.Property in ('sourceElement','linkedAssociation', 'linkedAttribute')                                 " & vbNewLine & _
							" inner join t_package p on p.Package_ID = o.Package_ID                                                                                                             " & vbNewLine & _
							" left join t_package p1 on p1.Package_ID = p.Parent_ID                                                                                                             " & vbNewLine & _
							" where 1 = 1                                                                                                                                                       " & vbNewLine & _
							" and o.Object_Type = 'Class'                                                                                                                                       " & vbNewLine & _
							" and o.Stereotype = 'table'                                                                                                                                        " & vbNewLine & _
							" and o.Package_ID in (" & sourcePackageTreeIDs  & ")                                                                                                               " & vbNewLine & _
							" )tdm                                                                                                                                                              " & vbNewLine & _
							" full outer join (                                                                                                                                                 " & vbNewLine & _
							" select  o.ea_guid AS CLASSGUID, o.Object_Type AS CLASSTYPE, o.name as Entity, null as Attribute, o.Object_Type as Type, o.ea_guid as LDMGuid, p1.Name as [Domain] " & vbNewLine & _
							" , tc.Value as TimeSliced, vs.Value as Versioned                                                                                                                   " & vbNewLine & _
							" ,null as Mandatory, null as AttributeType, null as facets                                                                                                         " & vbNewLine & _
							" from t_object o                                                                                                                                                   " & vbNewLine & _
							" inner join t_package p on p.Package_ID = o.Package_ID                                                                                                             " & vbNewLine & _
							" left join t_package p1 on p1.Package_ID = p.Parent_ID                                                                                                             " & vbNewLine & _
							" left join t_objectProperties tc on tc.Object_ID = o.Object_ID                                                                                                     " & vbNewLine & _
							"									and tc.Property = 'Timesliced'                                                                                                  " & vbNewLine & _
							" left join t_objectProperties vs on vs.Object_ID = o.Object_ID                                                                                                     " & vbNewLine & _
							"									and vs.Property = 'Versioned'									                                                                " & vbNewLine & _
							" where 1 = 1                                                                                                                                                       " & vbNewLine & _
							" and o.Object_Type = 'Class'                                                                                                                                       " & vbNewLine & _
							" and o.Package_ID in (" & targetPackageTreeIDs  & ")                                                                                                               " & vbNewLine & _
							" union all                                                                                                                                                         " & vbNewLine & _
							" select a.ea_guid AS CLASSGUID, 'Attribute' AS CLASSTYPE, o.name as Owner, a.Name as Attribute, 'Attribute' as Type, a.ea_guid as LDMGuid, p1.Name as [Domain]     " & vbNewLine & _
							" , tc.Value as TimeSliced, vs.Value as Versioned                                                                                                                   " & vbNewLine & _
							" , a.lowerbound as Mandatory, isnull(cl.Name, a.Type) as AttributeType                                                                                             " & vbNewLine & _
							",  case when ftd.Property is not null then ftd.Property + ':' + ftd.Value + char(13) + char(10) else '' end                                                        " & vbNewLine & _
							"+ case when ffd.Property is not null then ffd.Property + ':' + ffd.Value + char(13) + char(10) else '' end                                                         " & vbNewLine & _
							"+ case when fl.Property is not null then fl.Property + ':' + fl.Value + char(13) + char(10) else '' end                                                            " & vbNewLine & _
							"+ case when fmie.Property is not null then fmie.Property + ':' + fmie.Value + char(13) + char(10) else '' end                                                      " & vbNewLine & _
							"+ case when fme.Property is not null then fme.Property + ':' + fme.Value + char(13) + char(10) else '' end                                                         " & vbNewLine & _
							"+ case when fmi.Property is not null then fmii.Property + ':' + fmii.Value + char(13) + char(10) else '' end                                                       " & vbNewLine & _
							"+ case when fmi.Property is not null then fmi.Property + ':' + fmi.Value + char(13) + char(10) else '' end                                                         " & vbNewLine & _
							"+ case when fmil.Property is not null then fmil.Property + ':' + fmil.Value + char(13) + char(10) else '' end                                                      " & vbNewLine & _
							"+ case when fml.Property is not null then fml.Property + ':' + fml.Value + char(13) + char(10) else '' end                                                         " & vbNewLine & _
							"+ case when fp.Property is not null then fp.Property + ':' + fp.Value + char(13) + char(10) else '' end                                                            " & vbNewLine & _
							"as facets                                                                                                                                                          " & vbNewLine & _
							" from t_attribute a                                                                                                                                                " & vbNewLine & _
							" inner join t_object o on o.Object_ID = a.Object_ID                                                                                                                " & vbNewLine & _
							" inner join t_package p on p.Package_ID = o.Package_ID                                                                                                             " & vbNewLine & _
							" left join t_package p1 on p1.Package_ID = p.Parent_ID                                                                                                             " & vbNewLine & _
							"  left join t_attributetag tc on tc.ElementID = a.ID                                                                                                               " & vbNewLine & _
							"									and tc.Property = 'Timesliced'                                                                                                  " & vbNewLine & _
							" left join t_attributetag vs on vs.ElementID = a.ID                                                                                                                " & vbNewLine & _
							"									and vs.Property = 'Versioned'	                                                                                                " & vbNewLine & _
							" left join t_object cl on cl.Object_ID = a.Classifier                                                                                                              " & vbNewLine & _
							" left join t_objectproperties ftd on ftd.Object_ID = cl.Object_ID                                                                                                  " & vbNewLine & _
							"								and ftd.Property = 'totalDigits'                                                                                                    " & vbNewLine & _
							"left join t_objectproperties ffd on ffd.Object_ID = cl.Object_ID                                                                                                   " & vbNewLine & _
							"								and ffd.Property = 'fractionDigits'                                                                                                 " & vbNewLine & _
							"left join t_objectproperties fl on fl.Object_ID = cl.Object_ID                                                                                                     " & vbNewLine & _
							"								and fl.Property = 'length'                                                                                                          " & vbNewLine & _
							"left join t_objectproperties fmie on fmie.Object_ID = cl.Object_ID                                                                                                 " & vbNewLine & _
							"								and fmie.Property = 'minExclusive'                                                                                                  " & vbNewLine & _
							"left join t_objectproperties fme on fme.Object_ID = cl.Object_ID                                                                                                   " & vbNewLine & _
							"								and fme.Property = 'maxExclusive'                                                                                                   " & vbNewLine & _
							"left join t_objectproperties fmii on fmii.Object_ID = cl.Object_ID                                                                                                 " & vbNewLine & _
							"								and fmii.Property = 'minInclusive'                                                                                                  " & vbNewLine & _
							"left join t_objectproperties fmi on fmi.Object_ID = cl.Object_ID                                                                                                   " & vbNewLine & _
							"								and fmi.Property = 'maxInclusive'                                                                                                   " & vbNewLine & _
							"left join t_objectproperties fmil on fmil.Object_ID = cl.Object_ID                                                                                                 " & vbNewLine & _
							"								and fmil.Property = 'minLength'                                                                                                     " & vbNewLine & _
							"left join t_objectproperties fml on fml.Object_ID = cl.Object_ID                                                                                                   " & vbNewLine & _
							"								and fml.Property = 'maxLength'                                                                                                      " & vbNewLine & _
							"left join t_objectproperties fp on fp.Object_ID = cl.Object_ID                                                                                                     " & vbNewLine & _
							"								and fp.Property = 'pattern'                                                                                                         " & vbNewLine & _
							" where 1 = 1                                                                                                                                                       " & vbNewLine & _
							" and o.Object_Type = 'Class'                                                                                                                                       " & vbNewLine & _
							" and o.Package_ID in (" & targetPackageTreeIDs  & ")                                                                                                               " & vbNewLine & _
							" )ldm on ldm.LDMGuid = tdm.LDMGuid                                                                                                                                 " & vbNewLine & _
							" left join t_package p on p.ea_guid = tdm.LDMGuid                                                                                                                  " & vbNewLine & _
							" order by isnull(ldm.Entity, 'ZZ'), ldm.Attribute, tdm.Entity, tdm.Attribute                                                                                       "
    dim queryResults
	'debug
'	Dim debugFile
'	set debugFile = new TextFile
'	debugFile.Contents = sqlGetMappingDetails
'	debugFile.FullPath = "h:\temp\mappingQuery.txt"
'	debugFile.Save
	set queryResults = getArrayListFromQuery(sqlGetMappingDetails)
	'return
	set getMappingDetails = queryResults
end function