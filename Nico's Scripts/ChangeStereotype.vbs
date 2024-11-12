'[group=Nico's Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript


'
' Script Name: ChangeStereotype
' Author: Nico Celen
' Purpose: Converteren van het stereotype van een element
' Date: 2022-09-14
'

const outPutName = "ChangeStereotype"

'const oldStereotypeName = "Business Capability"
'const newStereotypeName = "Metamodel::Business Capability"
'const newTypeName = "Class"

'const oldStereotypeName = "Business Functie"
'const newStereotypeName = "Metamodel::Business Functie"
'const newTypeName = "Activity"

'const oldStereotypeName = "Business Proces"
'const newStereotypeName = "Metamodel::Business Proces"
'const newTypeName = "Activity"

'const oldStereotypeName = "Business Rol"
'const newStereotypeName = "Metamodel::Business Rol"
'const newTypeName = "Class"

'const oldStereotypeName = "Informatie Concept"
'const newStereotypeName = "Metamodel::Informatie Concept"
'const newTypeName = "Class"

'const oldStereotypeName = "Business Service"
'const newStereotypeName = "Metamodel::Business Service"
'const newTypeName = "Activity"

'const oldStereotypeName = "Informatiestroom"
'const newStereotypeName = "Metamodel::Informatiestroom"
'const newTypeName = "ControlFlow"

'const oldStereotypeName = "Applicatie"
'const newStereotypeName = "Metamodel::Applicatie"
'const newTypeName = "Component"

'const oldStereotypeName = "Applicatie Component"
'const newStereotypeName = "Metamodel::Applicatie Component"
'const newTypeName = "Component"

'const oldStereotypeName = "Technologie Component"
'const newStereotypeName = "Metamodel::Technologie Component"
'const newTypeName = "Activity"

'const oldStereotypeName = "Technologie"
'const newStereotypeName = "Metamodel::Technologie"
'const newTypeName = "Class"

'const oldStereotypeName = "Applicatie Interface"
'const newStereotypeName = "Metamodel::Applicatie Interface"
'const newTypeName = "Interface"

'const oldStereotypeName = "Data Object"
'const newStereotypeName = "Metamodel::Data Object"
'const newTypeName = "Class"

'const oldStereotypeName = "Integratie"
'const newStereotypeName = "Metamodel::Integratie"
'const newTypeName = "Class"


sub main

 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'report progress
 Repository.WriteOutput outPutName, now() & " Start Changing Stereotype" & outputName, 0
 'do the actual work
 changeStereotypes
 'report progress
 Repository.WriteOutput outPutName, now() & " Finished Changing Stereotype " & outputName, 0
end sub

function changeStereotypes
 dim package as EA.Package
 set package = Repository.GetTreeSelectedPackage
 
 if package is nothing then
  exit function
 end if
 'process the package
 changeStereotypesForPackage package
 
 'reload package
 Repository.ReloadPackage package.PackageID
 
end function 

function changeStereotypesForPackage(package)
 Repository.WriteOutput outPutName, now() & " Processing package '" & package.Name & "'",  0
 changeStereotypesForElement package
 'process subpackages
 dim subpackage as EA.Package
 for each subPackage in package.Packages
  changeStereotypesForPackage subPackage
 next
end function

function changeStereotypesForElement(parentElement)
 'process elements
 dim element as EA.Element
 for each element in parentElement.Elements
  Repository.WriteOutput outPutName, now() & "  Element stereotype in package '" & element.Stereotype & "'",  0
  if lcase(element.Stereotype) = lcase(oldStereotypeName) then
   Repository.WriteOutput outPutName, now() & " Updating element '" & element.Name & "'",  0
   element.Type = newTypeName
   element.Update
   element.StereotypeEx = newStereotypeName
   element.Update
  end if
  'process subElements
  changeStereotypesForElement element
 next
end function
 

main