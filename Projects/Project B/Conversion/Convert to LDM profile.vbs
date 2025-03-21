'[path=\Projects\Project B\Conversion]
'[group=Conversion]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Convert to LDM profile
' Author: Geert Bellekens
' Purpose: Converts the model of the selected package tothe LDM profile
' Date: 2025-01-16
'

const outPutName = "Convert to LDM profile"
const classStereotype = "LDM::LDM_Class"
const attributeStereotype = "LDM::LDM_Attribute"
const enumerationStereotype = "LDM::LDM_Enumeration"
const datatypeStereotype = "LDM::LDM_Datatype"
const associationSteretoype = "LDM::LDM_Association"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get the selected package
	dim package as EA.Package
	set package = Repository.GetTreeSelectedPackage()
	'let the user know we started
	Repository.WriteOutput outPutName, now() & " Starting " & outPutName & " for package '"& package.Name &"'", 0
	'do the actual work
	convertPackageToLDMProfile package
	'let the user know it is finished
	Repository.WriteOutput outPutName, now() & " Finished " & outPutName & " for package '"& package.Name &"'", 0
end sub

function convertPackageToLDMProfile(package)
	Repository.WriteOutput outPutName, now() & " Processing package '"& package.Name &"'", 0
	dim element as EA.Element
	'convert elements
	for each element in package.Elements
		convertElementToLDMProfile element
	next
	'convert diagrams
	dim diagram as EA.Diagram
	for each diagram in package.Diagrams
		if diagram.Type = "Logical" _
		  and not instr(diagram.StyleEx, "MDGDgm=LDM::LDM Diagram;") > 0 then
			'report progress
			Repository.WriteOutput outPutName, now() & " Migrating diagram '"& package.Name & "." & diagram.Name &"'", 0
			diagram.Metatype = "LDM::LDM Diagram"
			diagram.StyleEx = setValueForKey(diagram.StyleEx, "MDGDgm", "LDM::LDM Diagram")
			diagram.StyleEx = setValueForKey(diagram.StyleEx, "HideConnStereotype", "1")
			diagram.ExtendedStyle = setValueForKey(diagram.ExtendedStyle, "HideStereo", "1")
			diagram.ExtendedStyle = setValueForKey(diagram.ExtendedStyle, "HideEStereo", "1")
			diagram.Update
		end if
	next
	'process subPackages
	dim subPackage as EA.Package
	for each subPackage in package.Packages
		convertPackageToLDMProfile subPackage
	next
end function

function convertElementToLDMProfile(element)
	Repository.WriteOutput outPutName, now() & " Processing element '"& element.Name &"'", 0
	dim targetStereo
	targetStereo = ""
	'convert element
	select case lcase(element.Type)
		case "class"
			targetStereo = classStereotype
		case "datatype"
			migrateLDMDataType element
			exit function
		case "enumeration"
			targetStereo = enumerationStereotype
	end select
	if len(targetStereo) > 0 _
	  and element.FQStereotype <> targetStereo then
		element.StereotypeEx = targetStereo
		element.Update
	end if
	if element.Type = "Class" then
		'convert attributes
		dim attribute as EA.Attribute
		for each attribute in element.Attributes
			if attribute.FQStereotype <> attributeStereotype then
				attribute.StereotypeEx = attributeStereotype
				attribute.Update
			end if
		next
		'convert associations
		dim connector as EA.Connector
		for each connector in element.Connectors
			if connector.Type = "Association" _
			  and connector.FQStereotype <> associationSteretoype then
				connector.StereotypeEx = associationSteretoype
				connector.Update
			end if
		next
	end if
end function

function migrateLDMDataType(element)
	'skip if already the correct stereotype
	if element.FQStereotype = datatypeStereotype then
		exit function
	end if
	'report progress
	Repository.WriteOutput outPutName, now() & " Migrating Datatype '"& element.Name & "'", 0
	'get tagged value values and delete old tags
	dim fractionDigits
	fractionDigits = getTaggedValueValue(element, "fractionDigits")
	dim maxExclusive
	maxExclusive = getTaggedValueValue(element, "maxExclusive")
	dim maxInclusive
	maxInclusive = getTaggedValueValue(element, "maxInclusive")
	dim maxLength
	maxLength = getTaggedValueValue(element, "maxLength")
	dim minInclusive
	minInclusive = getTaggedValueValue(element, "minInclusive")
	dim minLength
	minLength = getTaggedValueValue(element, "minLength")
	dim pattern
	pattern = getTaggedValueValue(element, "pattern")
	dim totalDigits
	totalDigits = getTaggedValueValue(element, "totalDigits")
	'delete old tags
	deleteTag element, "fractionDigits"
	deleteTag element, "maxExclusive"
	deleteTag element, "maxInclusive"
	deleteTag element, "maxLength"
	deleteTag element, "minInclusive"
	deleteTag element, "minLength"
	deleteTag element, "pattern"
	deleteTag element, "totalDigits"
	'set correct stereotype
	element.StereotypeEx = datatypeStereotype
	element.Update
	'refresh tags
	element.TaggedValues.Refresh
	'copy values from old tags
	setTagValue element, "fractionDigits", fractionDigits
	setTagValue element, "maxExclusive", maxExclusive
	setTagValue element, "maximum", maxInclusive
	setTagValue element, "maxLength", maxLength
	setTagValue element, "minimum", minInclusive
	setTagValue element, "minLength", minLength
	setTagValue element, "pattern", pattern
	setTagValue element, "totalDigits", totalDigits
end function

main