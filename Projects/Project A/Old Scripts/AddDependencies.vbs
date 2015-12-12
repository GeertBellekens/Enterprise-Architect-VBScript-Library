'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
option explicit
 
 
!INC Local Scripts.EAConstants-VBScript
 
'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: AddDependencies
' Purpose: Browse the elements of a complexType, and if an element is itself of a type referring to
'          a complexType, add a dependency link to this complexType.
'          Multiple links are created to dependent complex type if multiple elements are of this complexType.
'          The script also verifies that no dependencies exists to complexType with no corresponding inner element.
' Date:  2014-10-13
'
dim attributeConnectorIds()
'
' Project Browser Script main function
'
sub OnProjectBrowserScript()
            
            ' Get the type of element selected in the Project Browser
            dim treeSelectedType
            treeSelectedType = Repository.GetTreeSelectedItemType()
            
            ' Handling Code: Uncomment any types you wish this script to support
            ' NOTE: You can toggle comments on multiple lines that are currently
            ' selected with [CTRL]+[SHIFT]+[C].
            select case treeSelectedType
            
                        case otElement
                                    ' Code for when an element is selected
                                    dim theElement as EA.Element
                                    set theElement = Repository.GetTreeSelectedObject()
                                    AddDependencyForObject theElement
                                    Session.Prompt "Script finished. Look at output tab for log details.", promptOK
                                    
                        case otPackage
                                    ' Code for when a package is selected
                                    dim thePackage as EA.Package
                                    set thePackage = Repository.GetTreeSelectedObject()
                                    AddDependencyForPackage thePackage
                                    Session.Prompt "Script finished. Look at output tab for log details.", promptOK
                                    
                        case otDiagram
                                    ' Code for when a diagram is selected
                                    dim theDiagram as EA.Diagram
                                    set theDiagram = Repository.GetTreeSelectedObject()
                                    AddDependencyForDiagram theDiagram
									Repository.ReloadDiagram theDiagram.DiagramID
                                    Session.Prompt "Script finished. Look at output tab for log details.", promptOK
                                    
'                       case otAttribute
'                                   ' Code for when an attribute is selected
'                                   dim theAttribute as EA.Attribute
'                                   set theAttribute = Repository.GetTreeSelectedObject()
'                                   
'                       case otMethod
'                                   ' Code for when a method is selected
'                                   dim theMethod as EA.Method
'                                   set theMethod = Repository.GetTreeSelectedObject()
                        
                        case else
                                    ' Error message
                                    Session.Prompt "This script does not support items of this type.", promptOK
                                    
            end select
            
end sub
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Add dependencies for all XSDComplexType of the selected diagram '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub AddDependencyForDiagram(theDiagram)
 
dim EADiagramElement as EA.DiagramObject
dim EAElement as EA.Element
 
            ' Browse all elements of the diagram
            for each EADiagramElement In theDiagram.DiagramObjects
                        set EAElement=Repository.GetElementByID(EADiagramElement.ElementID)
                        AddDependencyForObject(EAElement)
            next
end sub
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Add dependencies for all XSDComplexType of the selected package '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
sub AddDependencyForPackage(thePackage)
 
dim EAElement as EA.Element
dim EAPackage as EA.Package
 
            ' Treat all elements in the package
            for each EAElement in thePackage.Elements                     
                        AddDependencyForObject EAElement
            next
            
            ' Browse sub packages
            for each EAPackage in thePackage.Packages
                        AddDependencyForPackage EAPackage
            next
end sub
 
sub AddDependencyForObject(theElement)
 
dim EAAttribute as EA.Attribute
dim classifierElement as EA.Element
 
' Reset array
ReDim attributeConnectorIds(1)
 
Session.Output "Treating element: " + theElement.Name
if (theElement.Stereotype = "XSDcomplexType") then
            for each EAAttribute in theElement.Attributes
                        if (EAAttribute.ClassifierID <> 0) then
                                    set classifierElement=Repository.GetElementByID(EAAttribute.ClassifierID)
                                    ' Add dependecy only for complex types
                                    if (classifierElement.Stereotype = "XSDcomplexType") then
                                                AddOrUpdateDependencyConnector theElement,EAAttribute,classifierElement                                              
                                    end if
                        end if
            next
            CheckDependencies theElement
end if
 
end sub 
 
sub AddOrUpdateDependencyConnector(sourceElement,sourceAttribute,targetElement)
 
dim EAConnector as EA.Connector
dim EAConnectorEnd as EA.ConnectorEnd
dim cardinality
 
' Check if the connector already exists
set EAConnectorEnd=Nothing
for each EAConnector in sourceElement.Connectors
            if (EAConnector.SupplierID = targetElement.ElementID) then
                        set EAConnectorEnd = EAConnector.SupplierEnd
                        if (EAConnectorEnd.Role = sourceAttribute.Name) then
                                    exit for
                        else
                                    set EAConnectorEnd=Nothing
                        end if
            end if
next
 
' If connector not found create one          
if (EAConnectorEnd Is Nothing) then
            Session.Output "Add dependecy for attribute : " + sourceAttribute.Name
            set EAConnector=sourceElement.Connectors.AddNew("","Dependency")
            EAConnector.SupplierID=targetElement.ElementID
            EAConnector.Update
            set EAConnectorEnd=EAConnector.SupplierEnd
            EAConnectorEnd.Role=sourceAttribute.Name
            EAConnectorEnd.Update
else
            Session.Output "Dependency for attribute : " + sourceAttribute.Name + " already exists. No need to create it."
end if
 
' Set attribute multiplicity to connector
cardinality = ""
select case sourceAttribute.LowerBound
            case "0"
                        cardinality="0.."
            case "1"
                        cardinality="1.."
end select
if (cardinality <> "") then
            select case sourceAttribute.UpperBound
                        case "1"
                                    cardinality = cardinality + "1"
                        case "*"
                                    cardinality = cardinality + "*"
                        case "-1"
                                    cardinality = cardinality + "*"
                        case else
                                    cardinality = ""
            end select
end if
if (cardinality <> "") then
            EAConnectorEnd.Cardinality = cardinality
            EAConnectorEnd.Update
end if
 
' Add the connector to know attribute connector ID list
ReDim Preserve attributeConnectorIds(UBound(attributeConnectorIds) + 1)
attributeConnectorIds(UBound(attributeConnectorIds)) = EAConnector.ConnectorID
 
end sub
 
sub CheckDependencies(theElement)
 
dim EAConnector as EA.Connector
dim EAConnectorEnd as EA.ConnectorEnd
dim connectedElement as EA.Element
dim isConnectorFound
dim I
 
' Verify that each connector of the element is linked with a known attribute
' otheriwse display a warning
for each EAConnector in theElement.Connectors
            if ((EAConnector.Type = "Dependency") and (EAConnector.ClientID = theElement.ElementID)) then
                        isConnectorFound=false
                        for I = 1 to UBound(attributeConnectorIds)
                                    if (EAConnector.ConnectorID = attributeConnectorIds(I)) then
                                                isConnectorFound = true
                                                exit for
                                    end if
                        next
                        
                        if (not isConnectorFound) then
                                    set EAConnectorEnd = EAConnector.SupplierEnd
                                    set connectedElement = Repository.GetElementByID(EAConnector.SupplierID)
                                    Session.Output "!!! WARNING : Connector(" + CStr(EAConnector.ConnectorID) + ") associated with element " + connectedElement.Name + " with role " + EAConnectorEnd.Role + " is not associated with any attribute."
                        end if
            end if
next
 
end sub
 
OnProjectBrowserScript