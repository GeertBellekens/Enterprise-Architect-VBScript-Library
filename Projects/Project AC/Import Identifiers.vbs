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

const outPutName = "Import Identifiers"


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
	   'create output tab
		Repository.CreateOutputTab outPutName
		Repository.ClearOutput outPutName
		Repository.EnsureOutputVisible outPutName
		'set timestamp
		Repository.WriteOutput outPutName, "Starting import identifiers " & now(), 0
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
				'check if the IdentifierFQN is not empty and is a valid FQN
				if len(IdentifierFQN) > 0 AND instrRev(IdentifierFQN,"::") > 1 then
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
			end if
		next
		'set timestamp
		Repository.WriteOutput outPutName, "End import identifiers " & now(), 0
	end if
end sub

function setIdentifierAttribute(logicalPackage,classFQN,idName)
	dim attribute as EA.Attribute
	set attribute = selectObjectFromQualifiedName(logicalPackage,nothing, classFQN & "::" & idName , "::") 
	if not attribute is nothing then
		'set isID property on attribute
		'log progress
		Repository.WriteOutput outPutName, "setting {id} on attribute " & classFQN & "." & attribute.Name,0
		attribute.IsID = true
		attribute.Update
	else
		'log the fact that we didn't find it
		Repository.WriteOutput outPutName, "ERROR: could not find attribute for " & classFQN & "." & idName,0
	end if
	
end function

function setIdenfifierAssociation(logicalPackage,classFQN,idName)
	dim classElement as EA.Element
	set classElement = selectObjectFromQualifiedName(logicalPackage,nothing, classFQN, "::")
	if not classElement is nothing then
		'find the associationEnd
		dim association as EA.Connector
		'register the fact that we found it or not
		dim foundIt
		foundIt = false
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
					
					if associationEnd.Role = idName _
						AND left(associationEnd.Cardinality,1) = "1" then 'only for obligatory associations
						if not foundIt then
							'log progress
							Repository.WriteOutput outPutName, "setting {id} on association " & classFQN & "." & idName,0
							'found the correct one
							associationEnd.Constraint = "id"
							associationEnd.Update
							'register that we found one
							foundIt = true
						else
							Repository.WriteOutput outPutName, "ERROR: found duplicate rolename for " & classFQN & "." & idName,0
						end if
					end if
				end if
			end if
		next
		if not foundIt then
			'log the fact that we didn't find it
			Repository.WriteOutput outPutName, "ERROR: could not find association role for " & classFQN & "." & idName,0
		end if
	end if
end function

main