'[path=\Projects\Project AC]
'[group=Acerta Scripts]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'TODO: finish!
'
' Script Name: Import KBC Codes
' Author: Geert Bellekens
' Purpose: Import the KBC Codes on Attributes
' Date: 2016-07-14
'

const outPutName = "Import KBC Codes"


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
		Repository.WriteOutput outPutName, "Starting import KBC codes " & now(), 0
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
			'there should be 3 parts in the csv file: ClassFQN, attribute name and KBC Code
			if Ubound(parts) = 2 then
				dim IdentifierFQN, idName, isAttribute, KBCCode
				IdentifierFQN = parts(0)
				'log progress
				Repository.WriteOutput outPutName, "Processing " & IdentifierFQN,0
				'check if the IdentifierFQN is not empty and is a valid FQN
				if len(IdentifierFQN) > 0 AND instrRev(IdentifierFQN,"::") > 1 then
					idName = parts(1)
					KBCCode = parts(2)
					'set KBCCode on attribute
					setKBCCodeOnAttribute logicalPackage,IdentifierFQN,idName,KBCCode
				end if
			end if
		next
		'set timestamp
		Repository.WriteOutput outPutName, "Finished import KBC codes  " & now(), 0
	end if
end sub

function setKBCCodeOnAttribute(logicalPackage,classFQN,attributeName, KBCCode)
	dim attribute as EA.Attribute
	set attribute = selectObjectFromQualifiedName(logicalPackage,nothing, classFQN & "::" & attributeName , "::") 
	if not attribute is nothing then
		'set isID property on attribute
		'log progress
		Repository.WriteOutput outPutName, "setting KBC Code &  on attribute " & classFQN & "." & attribute.Name,0
		dim taggedValue as EA.TaggedValue
		set taggedValue = getOrCreateTaggedValue(attribute,"KBC Code")
		taggedValue.Value = KBCCode
		taggedValue.Update
	end if
end function


main