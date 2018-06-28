'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' Script Name: Map Business Entities
' Author: Geert Bellekens
' Purpose: maps the attributes of the Business entities related to this message to the attributes of this message
' Date: 2018-05-04
'
'name of the output tab
const outPutName = "Map Business entities"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	
	'get the selected element
	dim selectedElement as EA.Element
	set selectedElement = Repository.GetContextObject
	if selectedElement.ObjectType = otElement _
	  AND isMessageAssembly(selectedElement) then
		'tell the user we are starting
		Repository.WriteOutput outPutName, now() & " Starting Map Business entities for '" & selectedElement.Name & "'", selectedElement.ElementID
		'do the actual work
		mapBusinessEntities(selectedElement)
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Finished Map Business entities for '" & selectedElement.Name & "'", selectedElement.ElementID
	elseif selectedElement.ObjectType = otPackage then
		'tell the user we are starting
		Repository.WriteOutput outPutName, now() & " Starting Map Business entities for package '" & selectedElement.Name & "'", selectedElement.Element.ElementID
		'map the business entities for the selected package and all subPackages
		mapBusinessEntitiesForPackage(selectedElement)
		'tell the user we are finished
		Repository.WriteOutput outPutName, now() & " Finished Map Business entities for package '" & selectedElement.Name & "'", selectedElement.Element.ElementID
	else
		msgbox "Please select a MessageAssembly (element with stereotype «MA»)"
	end if
end sub

function mapBusinessEntitiesForPackage(package)
	dim messageAssemblies
	'get the message assemblies in this package (and subpackages)
	set messageAssemblies = getMessageAssemblies(package)
	dim messageAssembly as EA.Element
	for each messageAssembly in messageAssemblies
		'process each messageAssembly
		mapBusinessEntities(messageAssembly)
	next
end function

function getMessageAssemblies(package)
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	dim sqlGetMessageAssemblies
	sqlGetMessageAssemblies = "select o.Object_ID from t_object o                    " & _
							" inner join t_xref x on x.Client = o.ea_guid            " & _
							" 					and x.Name = 'Stereotypes'           " & _
							" 					and x.Description like '%;Name=MA;%' " & _
							" inner join t_package p on o.Package_ID = p.Package_ID  " & _
							" inner join t_object po on po.ea_guid = p.ea_guid       " & _
							" 						and po.Stereotype = 'DOCLibrary' " & _
							" where o.Package_ID in (" & packageTreeIDs & ") 			 "
	set getMessageAssemblies = getElementsFromQuery(sqlGetMessageAssemblies)
end function

function mapBusinessEntities(selectedElement)
	'tell the user what we are doing
	Repository.WriteOutput outPutName, now() & " Processing MA '" & selectedElement.Name & "'", selectedElement.ElementID
	'first get ID's of the related business entities. These are found on a diagram in the same package as the business entity that is 
	' linked to the «InvEnvelop» element that is linked to the selected «MA»
	dim businessEntityIDs
	businessEntityIDs = getBusinessEntityIDs(selectedElement)
	'get all BBIE' in the same package
	dim bbies
	set bbies = getBBies(selectedElement)
	'loop bbies
	dim bbie as EA.Attribute
	for each bbie in bbies
		Repository.WriteOutput outPutName, now() & " Processing BBIE '" & bbie.Name & "'", bbie.ParentID
		'get the relevant business entity attributes
		dim businessAttributes
		set businessAttributes = getMappedBusinessAttributes(bbie, businessEntityIDs)
		dim businessAttribute as EA.Attribute
		'remove all mapping business attributes
		dim mappingTaggedValue as EA.AttributeTag
		dim i
		for i = bbie.TaggedValues.Count -1 to 0 step -1
			set mappingTaggedValue = bbie.TaggedValues.GetAt(i)
			if lcase(mappingTaggedValue.Name) = lcase("mappedBusinessAttribute") then
				bbie.TaggedValues.DeleteAt i, false
			end if
		next
		for each businessAttribute in businessAttributes
			'inform the user
			Repository.WriteOutput outPutName, now() & " Mapping Business Attribute '" & businessAttribute.Name & "' to BBIE '" & bbie.Name & "'", businessAttribute.ParentID
			'add the tagged value
			
			set mappingTaggedValue = bbie.TaggedValues.AddNew("mappedBusinessAttribute",businessAttribute.AttributeGUID)
			mappingTaggedValue.Update
		next
	next
end function


function getMappedBusinessAttributes(bbie, businessEntityIDs)
	dim sqlGetMappedBusinessattributes
	sqlGetMappedBusinessattributes = "select distinct ba.ID from t_attribute a                                    " & _
									" inner join t_attributetag atv on atv.ElementID = a.ID                       " & _
									" 								and atv.Property = 'sourceAttribute'          " & _
									" inner join t_attribute sa on sa.ea_guid = atv.VALUE                         " & _
									" inner join t_attributetag btv on btv.VALUE = sa.ea_guid                     " & _
									" 								and btv.Property = 'sourceAttribute'          " & _
									" inner join t_attribute ba on ba.ID = btv.ElementID                          " & _
									" inner join t_object be on be.Object_ID = ba.Object_ID                       " & _
									" inner join t_xref x2 on x2.Client = be.ea_guid							  " & _
									" 						and x2.Name = 'Stereotypes'							  " & _
									" 						and (x2.Description like '%;Name=BusinessEntity;%' 	  " & _
									" 							or x2.Description like '%;Name=bEntity;%')		  " & _	
									" where a.ea_guid = '" & bbie.AttributeGUID & "'         					  " & _
									" and be.Object_ID in ("& businessEntityIDs &")                               "
	set getMappedBusinessAttributes = getattributesFromQuery(sqlGetMappedBusinessattributes)
end function

function getBBies(selectedElement)
	dim sqlGetBBies
	sqlGetBBies = " select distinct bbie.ID                                                " & _
				" from ((t_object o                                               " & _
				" inner join t_object abie on (abie.Package_ID = o.Package_ID     " & _
				" 							and abie.Stereotype = 'ABIE'))        " & _
				" inner join t_attribute bbie on (bbie.Object_ID = abie.Object_ID " & _
				" 							and bbie.Stereotype = 'BBIE'))        " & _
				" where o.ea_guid = '" & selectedElement.ElementGUID & "'         "
	set getBBies = getattributesFromQuery(sqlGetBBies)
end function


function getBusinessEntityIDs(selectedElement)
	dim sqlGetBusinessEntities
	sqlGetBusinessEntities = "select distinct dob.Object_ID from (((((((t_object o                             " & _
							" inner join t_connector c on (o.Object_ID = c.End_Object_ID                       " & _
							" 							and c.Connector_Type in ('Association', 'Aggregation') " & _
							" 							and c.DestRole = 'Assembly'))                          " & _
							" inner join t_object env on (env.Object_ID = c.Start_Object_ID                    " & _
							" 						and env.Stereotype = 'InfEnvelope'))                       " & _
							" inner join t_connector c2 on (c2.Start_Object_ID = env.Object_ID                 " & _
							" 							and c2.Connector_Type = 'Dependency'                   " & _
							" 							and c2.Stereotype = 'represents'))                     " & _
							" inner join t_object be on (be.Object_ID = c2.End_Object_ID                       " & _
							" 						and be.Stereotype in ('BusinessEntity','bEntity')))        " & _
							" inner join t_diagram d on (d.Package_ID = be.Package_ID                          " & _
							" 							and d.ParentID = 0))		                           " & _
							" inner join t_diagramobjects do on do.Diagram_ID = d.Diagram_ID)                  " & _
							" inner join t_object dob on (dob.Object_ID = do.Object_ID                         " & _
							" 							and dob.Stereotype in ('BusinessEntity','bEntity')))   " & _
							" where o.ea_guid = '" & selectedElement.ElementGUID & "'                          "
	dim businessEntityIDs 
	businessEntityIDs = getArrayFromQuery(sqlGetBusinessEntities)
	dim id
	dim idList
	set idList = CreateObject("System.Collections.ArrayList")
	for each id in businessEntityIDs
		if len(id) > 0 then
			idList.Add id
		end if
	next
	if idList.Count > 0 then
		getBusinessEntityIDs = Join(idList.ToArray(),",")
	else
		getBusinessEntityIDs = "0"
	end if
end function

function isMessageAssembly(selectedElement)
	'initialize on false
	isMessageAssembly = false
	'check if stereotype «MA» is present
	dim stereotypes
	dim stereotype
	stereotypes = split(selectedElement.StereotypeEx, ",")
	for each stereotype in stereotypes
		if stereotype = "MA" then
			isMessageAssembly = true
			exit for
		end if
	next
end function

main