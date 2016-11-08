'[path=\Projects\Project AC]
'[group=Acerta Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Create default views
' Author: Geert Bellekens
' Purpose: Creates a default view for each table
' Date: 2016-07-14
'

const outPutName = "Create Default Views"


sub main
	'select database
	dim databasePackage as EA.Package
	
	msgbox "select the «Database» or «DataModel» package"
	set databasePackage = selectPackage()
	if not databasePackage is nothing then
		if databasePackage.StereotypeEx = "Database" or databasePackage.StereotypeEx = "DataModel" then
			'create output tab
			Repository.CreateOutputTab outPutName
			Repository.ClearOutput outPutName
			Repository.EnsureOutputVisible outPutName
			'timestamp
			Repository.WriteOutput outPutName, "Starting create default views " & now(), 0
			'get the viewsPackage
			dim viewsPackage
			set viewsPackage = getViewsPackage(databasePackage)
			'find all tables
			dim packageIDstring
			packageIDstring = getPackageTreeIDString(databasePackage)
			dim getElementsQuery
			getElementsQuery = "select * from t_object o " & _
								" where o.Stereotype = 'table' " & _
								" and o.Package_ID in ("& packageIDstring &") "
			dim tables
			set tables = getElementsFromQuery(getElementsQuery)
			dim table
			'loop the tables and make a view for each table
			for each table in tables
				dim viewName
				viewname = left(table.Name,2) & "V" & mid(table.Name,4)
				'check if view doesn't exist yet
				if not elementExistsInPackage (viewsPackage, viewname) then
					'create new view
					Repository.WriteOutput outPutName, "Adding view  " & viewname, 0
					createDefaultView viewsPackage, viewName, table
				end if
			next
			'timestamp end
			Repository.WriteOutput outPutName, "Finished create default views " & now(), 0
		end if
	end if 
end sub

function createDefaultView(viewsPackage, viewName, table)
	dim newView as EA.Element
	set newView = viewsPackage.Elements.AddNew(viewName,"EAUML::view")
	newView.Gentype = table.Gentype
	newView.Update
	'add the definition and owner
	dim taggedValue as EA.TaggedValue
	for each taggedValue in newView.TaggedValues
		if taggedValue.Name = "Owner" then
			taggedValue.Value = getTaggedValueValue(table, "Owner")
			taggedValue.Update
		end if
	next
	'add the viewdef
	set taggedValue = newView.TaggedValues.AddNew("viewdef","")
	taggedValue.Value = "<memo>"
	taggedValue.Notes = "select * from " & getTaggedValueValue(table, "Owner") & "." & table.Name
	taggedValue.Update
end function

function elementExistsInPackage (package, elementName)
	dim element as EA.Element
	elementExistsInPackage = false
	for each element in package.Elements
		if element.Name = elementName then
			elementExistsInPackage = true
			exit for
		end if
	next
end function

'gets the views package for this datase.
'It is the package wiht the name "Views" on either this level or the next level.
'if not found then the datbaase package is returned
function getViewsPackage(databasePackage)
	dim subPackage as EA.Package
	dim subSubPackage as EA.Package
	dim viewsPackage
	'default
	set viewsPackage = databasePackage
	'loop packages
	for each subPackage in databasePackage.Packages
		if subPackage.Name = "Views" then
			set viewsPackage = subPackage
			exit for
		else
			'loop subPackages
			for each subSubPackage in subPackage.Packages
				if subSubPackage.Name = "Views" then
					set viewsPackage = subSubPackage
					exit for
				end if
			next
		end if
	next
	set getViewsPackage = viewsPackage
end function

main