'[path=\Projects\Project AC]
'[group=Acerta Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Import Identifiers
' Author: Geert Bellekens
' Purpose: Import the identifiers exported from MEGA's Candidate key members
' Date: 2016-07-14
'
sub main
	dim mappingFile
	set mappingFile = New TextFile
	'select source logical
	dim logicalPackage as EA.Package
	msgbox "select the logical package root (S-OAA-...)"
	set logicalPackage = selectPackage()
	'first select the mapping file
	if mappingFile.UserSelect("","CSV Files (*.csv)|*.csv") _
	   AND not logicalPackage is nothing then
		'split into lines
		dim lines
		lines = Split(mappingFile.Contents, vbCrLf)
		dim line
		for each line in lines
			'replace any "." with "::" 
			line = Replace(line,".","::")
			'split into logical and physical part
			dim parts
			parts = Split(line,";")
			'there should be 4 parts in the csv file: Identifier - Attribute or Role Name - AttributePath - RoleName + [ID]
			'we should have enough info from the name, and the fact that we know whether it is an attribut or a RoleName
			if Ubound(parts) = 3 then
				dim IdentifierFQN, idName, isAttribute
				IdentifierFQN = parts(0)
				idName = parts(1)
				if len(parts(2)) > 0 then
					isAttribute = true
				else
					isAttribute = false
				end if
				dim classFQN
				'remove the last part of of the IdentifierFQN in order to get the class name
				classFQN = mid(IdentifierFQN , 1 , instrRev(IdentifierFQN,"::") - 1)
				if isAttribute then
					'set identifier on attribute
					setIdentifierAttribute logicalPackage,classFQN,idName
				else
					'set identifier on association end
					setIdenfifierAssociation logicalPackage,classFQN,idName
				end if
			end if
		next
	end if
end sub

function setIdentifierAttribute(logicalPackage,classFQN,idName)
	dim attribute as EA.Attribute
	set attribute = selectObjectFromQualifiedName(logicalPackage,nothing, classFQN & "::" & idName , "::") 
	if not attribute is nothing then
		'set isID property on attribute
		attribute.IsID = true
		attribute.Update
	end if
end function

function setIdenfifierAssociation(logicalPackage,classFQN,idName)
	dim classElement as EA.Element
	set classElement = selectObjectFromQualifiedName(logicalPackage,nothing, classFQN, "::")
	if not classElement is nothing then
		'find the associationEnd
		dim association as EA.Connector
		for each association in classElement.Connectors
			if association.Type = "Association" or association.Type = "Aggregation" then
				dim associationEnd as EA.ConnectorEnd
				set associationEnd = nothing 'initialize to be sure
				if association.ClientID = classElement.ElementID then
					set associationEnd = association.SupplierEnd
				else
					set associationEnd = association.ClientEnd
				end if
				if not associationEnd is nothing then
					if associationEnd.Role = idName then
						'found the correct one
						associationEnd.Constraint = "id"
						associationEnd.Update
						exit for
					end if
				end if
			end if
		next
	end if
end function

main