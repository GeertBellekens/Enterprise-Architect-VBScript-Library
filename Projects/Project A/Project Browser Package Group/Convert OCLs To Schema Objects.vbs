'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

' {C2881351-5AA1-44dc-B546-FCD7B87F2761}
' Script Name: Convert OCLs To Schema Objects
' Author: Matthias Van der Elst
' Purpose: Convert the OCL's from the MessageAssemblies in the selected package to Schema Objects
' Date: 2017-09-20

const outPutName = "Create Schema From Package"
dim countAttribute
dim isroot 
isroot = true
dim dict
dim xmlDOM

function Main()

	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	
	if not selectedPackage is nothing then
		'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, now() & " Starting Create Schema From Package '"& selectedPackage.Name &"'", 0
		
		'get all the MA's in the selected package and subpackages
		dim sqlGetClasses
		sqlGetClasses = "select o.Object_ID from t_object o " & _
						"where o.Stereotype = 'MA' " & _
						"and o.Package_ID in (" & getPackageTreeIDString(selectedPackage) & ") order by o.Name"
		
		dim MAs
		set MAs = getElementsFromQuery(sqlGetClasses) 'all the MA's with the OCL's
		
		dim ma as EA.Element
		for each ma in MAs	
			'get the constraints from this MessageAssembly
			Repository.WriteOutput outPutName, now() &  " MessageAssembly: " & ma.Name , ma.ElementID 'ma toevoegen aan de dictionary
			dim constraints
			set constraints = getConstraints(ma)
			
			'get the OCL rules from the constraints
			dim OCLs
			set OCLs = getOCLs(constraints, ma)
			
			dim schema 
			set schema = new Schema
			'set the context
			schema.Context = ma
			Repository.WriteOutput outPutName, now() &  " Creating Schema for: "  & ma.Name, ma.ElementID
			'process the ocl's
			schema.processOCLs OCLs, outputName
			'debug
			debugPrintSchema schema
			'save the schema
			schema.save
		next	
	end if
	Repository.WriteOutput outPutName, now() &  " Finished!"  , 0 

end function

function debugPrintSchema(schema)
	Repository.WriteOutput outPutName, now() & " Schema for: " & schema.Context.Name, 0
	dim element
	for each element in schema.Elements.Items
		Repository.WriteOutput outPutName, now() & "     Element: " & element.Name, 0
		dim schemaProperty
		for each schemaProperty in element.Properties.Items
			Repository.WriteOutput outPutName, now() & "         SchemaProperty: " & schemaProperty.Name, 0
		next
	next
end function

function getOCLs(constraints, context)
	dim OCLs, arrLines
	set OCLs = CreateObject("System.Collections.ArrayList")
	dim constraint as EA.Constraint
	for each constraint in constraints
		Repository.WriteOutput outPutName, now() &  " parsing constraint: " & constraint.Name, 0 
		'first remove the comments
		dim trimmedConstraint
		Dim regExp		
		Set regExp = CreateObject("VBScript.RegExp")
		regExp.Global = True   
		regExp.IgnoreCase = False
		regExp.Pattern = "(--.*)"
		trimmedConstraint = regExp.Replace(Repository.GetFormatFromField("TXT",constraint.Notes), "")
		'then group by individual OCL statement
		dim statements
		statements = split(trimmedConstraint, "inv:")
		dim statement
		dim i
		i = 0
		for each statement in statements
			if len(statement) > 0 then
				set newOCL = new OCLStatement
				newOCL.Context = context
				newOCL.Statement = statement
				if newOCL.IsValid then
					'debug
					'debugPrintOCL newOCL,0
					'only add valid OCLs
					OCLs.Add newOCL
				else
					'report error
					Repository.WriteOutput outPutName, now() &  " ERROR: Could not parse OCL statement: " & newOCL.Statement, 0
				end if
			end if
		next
	next
	set getOCLs = OCLs
end function

function debugPrintOCL(newOCL, indentCount)
	dim indent
	indent = space(indentCount*4) 
	Repository.WriteOutput outPutName, now() & indent & " Statement: " & newOCL.Statement, 0
	Repository.WriteOutput outPutName, now() & indent & " Left: " & newOCL.LeftHand, 0
	Repository.WriteOutput outPutName, now() & indent & " Operator: " & newOCL.Operator, 0
	Repository.WriteOutput outPutName, now() & indent & " Right: " & newOCL.RightHand, 0
	if not newOCL.NextOCLStatement is nothing then
		debugPrintOCL newOCL.NextOCLStatement, indentCount + 4
	end if
end function

function getConstraints(ma)
	dim Constraints 
	set Constraints = CreateObject("System.Collections.ArrayList")
	dim constraint as EA.Constraint
	for each constraint in ma.Constraints
		if constraint.Type = "OCL2.0" then
			dim Facet, Template
			Facet = Left(Trim(constraint.Name),5)
			Template = Left(Trim(constraint.Name),16)
			if not Facet = "Facet" and not Template = "Template Payload" then
				Constraints.Add(constraint)
			end if
		end if
	next
	set getConstraints = Constraints
end function

main