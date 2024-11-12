'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Create schema from package contents
' Author: Geert Bellekens
' Purpose: Add the contents of the selected package to a new message composer schema
' Date: 2018-07-03
'
const outPutName = "Create Schema From Package"

function Main ()
 
 dim selectedPackage as EA.Package
 set selectedPackage = Repository.GetTreeSelectedPackage()
 if not selectedPackage is Nothing then
  
  'create output tab
  Repository.CreateOutputTab outPutName
  Repository.ClearOutput outPutName
  Repository.EnsureOutputVisible outPutName
  'set timestamp
  Repository.WriteOutput outPutName, "Starting Create Schema From Package at " & now(), 0
  
  'create the artifact
  dim artifact as EA.Element
  set artifact = createArtifact(selectedPackage)
  
  dim xmlDOM 
  set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
  'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
  xmlDOM.validateOnParse = false
  xmlDOM.async = false
  'add schema node
  dim xmlSchema 
  set xmlSchema = xmlDOM.createElement("schema")
   
  dim node 
  set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
  xmlDOM.appendChild node
 '
  dim xmlRoot 
  set xmlRoot = xmlDOM.createElement( "message" )
  xmlDOM.appendChild xmlRoot
  'add description node
  xmlRoot.appendChild createDescriptionNode(xmlDOM, artifact)
  
  'get all elements
  dim sqlGetClasses
  sqlGetClasses = "select o.Object_ID from t_object o " & _
      "where o.Object_Type in ('Class', 'DataType', 'Enumeration') " & _
      "and o.Package_ID in (" & getPackageTreeIDString(selectedPackage) & ") order by o.Name"
  dim allElements
  set allElements = getElementsFromQuery(sqlGetClasses)

  'count attribute
  dim xmlcountAtttr 
  set xmlcountAtttr = xmlDOM.createAttribute("count")
  xmlcountAtttr.nodeValue = allElements.Count
  xmlSchema.setAttributeNode(xmlcountAtttr)
  'add the schema node to the root
  xmlRoot.appendChild xmlSchema
  'create the dictionary holding all processes elements
  dim allNodes
  set allNodes = CreateObject("Scripting.Dictionary")
  'add all the elements to the schema
  dim element as EA.Element
  for each element in allElements
   'add node
   createElementNode xmlDom, xmlSchema, element, allNodes
  next
  'create the actual schema content in the database
  createTDocument artifact, xmlDOM.xml
  'update log
  Repository.WriteOutput outPutName, "Finished Create Schema From Package at " & now(), 0
  'debug
  'writefile "c:\\temp\\schemaContents.xml", xmlDOM.xml
 end if
end function

function createElementNode(xmlDOM, xmlSchema, element, allNodes)
 'update log
 Repository.WriteOutput outPutName, "Processing element " & element.Name, element.ElementID
 'check if the element does not exist yet
 if not allNodes.Exists(element.ElementID) then
  'add the node to the list
  allNodes.Add element.ElementID, element
 else
  exit function
 end if
 
 dim xmlClass
 set xmlClass = xmlDOM.createElement( "class" )
 
 'name attribute
 dim xmlNameAtttr 
 set xmlNameAtttr = xmlDOM.createAttribute("name")
 xmlNameAtttr.nodeValue = element.Name
 xmlClass.setAttributeNode(xmlNameAtttr)
 
 'guid attribute
 dim xmlguidAtttr 
 set xmlguidAtttr = xmlDOM.createAttribute("guid")
 xmlguidAtttr.nodeValue = element.ElementGUID
 xmlClass.setAttributeNode(xmlguidAtttr)
 
 'ancestry
 addAncestry xmlClass, xmlDOM, xmlSchema, element, allNodes

 'add propertiesnode
 dim xmlProperties
 set xmlProperties= xmlDOM.createElement("properties")
 
 'add attributes
 dim attribute as EA.Attribute
 for each attribute in element.Attributes
  xmlProperties.appendChild createPropertyNode (xmlDOM, attribute.AttributeGUID, "attribute")
  'add an element node for the type of this attribute
  if attribute.ClassifierID > 0 then
   dim attributeType as EA.Element
   set attributeType = Repository.GetElementByID(attribute.ClassifierID)
   'add the node for the attributeType 
   createElementNode xmlDOM, xmlSchema, attributeType, allNodes
  end if
 next
 
 'add associations only if they start at the given element
 dim relation as EA.Connector
 for each relation in element.Connectors
  if (relation.Type = "Association" _
  or relation.Type = "Aggregation" ) then
   'add association node
   xmlProperties.appendChild createPropertyNode (xmlDOM, relation.ConnectorGUID, "association")
   'add element node for the target of the relation
   dim targetElement as EA.Element
   set targetElement = Repository.GetElementByID(relation.SupplierID)
   'add the node for the target
   createElementNode xmlDOM, xmlSchema, targetElement, allNodes
  end if
 next
 
 'add xmlProperties to class node
 xmlClass.appendChild xmlProperties
 
 'add the xmlClass node to the schema
 xmlSchema.appendChild xmlClass
end function

function addAncestry(xmlClass, xmlDOM, xmlSchema, element, allNodes)
 'not for XSDSimpletypes
 if element.HasStereotype("XSDsimpleType") then
  exit function
 end if
 'loop base elements
 dim sqlGetBaseElements
 sqlGetBaseElements = "select c.End_Object_ID as Object_ID from t_connector c " & _
      " where c.Connector_Type = 'Generalization' " & _
      " and c.Start_Object_ID = " & element.ElementID
 dim baseElements
 set baseElements = getElementsFromQuery(sqlGetBaseElements)
 if baseElements.Count > 0 then
  'composition attribute
  dim xmlCompositionAttr
  set xmlCompositionAttr = xmlDOM.createAttribute("composition")
  xmlCompositionAttr.nodeValue = "inherit"
  xmlClass.setAttributeNode(xmlCompositionAttr)
  'add ancestry node
  dim xmlAncestry
  set xmlAncestry = xmlDOM.createElement("ancestry")
  'loop base elements
  dim baseElement as EA.Element
  for each baseElement in baseElements
   'create ancesterNode
   dim xmlAncestor
   set xmlAncestor = xmlDOM.createElement("ancestor")
   'name attribute
   dim xmlNameAtttr 
   set xmlNameAtttr = xmlDOM.createAttribute("name")
   xmlNameAtttr.nodeValue = baseElement.Name
   xmlAncestor.setAttributeNode(xmlNameAtttr)
   'guid attribute
   dim xmlguidAtttr 
   set xmlguidAtttr = xmlDOM.createAttribute("guid")
   xmlguidAtttr.nodeValue = baseElement.ElementGUID
   xmlAncestor.setAttributeNode(xmlguidAtttr)
   'add to ancestry node
   xmlAncestry.appendChild xmlAncestor
   'create element node for ancestor
   createElementNode xmlDOM, xmlSchema, baseElement, allNodes
  next
  'add to xmlClassNode
  xmlClass.appendChild xmlAncestry
 end if
end function

function createPropertyNode (xmlDOM, guid, propertyType)
 dim xmlProperty
 set xmlProperty = xmlDOM.createElement("property")
 
 'guid attribute
 dim xmlguidAtttr 
 set xmlguidAtttr = xmlDOM.createAttribute("guid")
 xmlguidAtttr.nodeValue = guid
 xmlProperty.setAttributeNode(xmlguidAtttr)
 
 'type attribute
 dim xmltypeAtttr 
 set xmltypeAtttr = xmlDOM.createAttribute("type")
 xmltypeAtttr.nodeValue = propertyType
 xmlProperty.setAttributeNode(xmltypeAtttr)
 
 'return node
 set createPropertyNode = xmlProperty
end function


function createDescriptionNode(xmlDOM, selectedElement)
 dim xmlDescription
 set xmlDescription = xmlDOM.createElement( "description" )
 
 'name attribute
 dim xmlNameAtttr 
 set xmlNameAtttr = xmlDOM.createAttribute("name")
 xmlNameAtttr.nodeValue = selectedElement.Name
 xmlDescription.setAttributeNode(xmlNameAtttr)
 
 'namespace attribute
 dim xmlnamespaceAtttr 
 set xmlnamespaceAtttr = xmlDOM.createAttribute("namespace")
 xmlnamespaceAtttr.nodeValue = ""
 xmlDescription.setAttributeNode(xmlnamespaceAtttr)
 
 'schemaset attribute
 dim xmlschemasetAtttr 
 set xmlschemasetAtttr = xmlDOM.createAttribute("schemaset")
 xmlschemasetAtttr.nodeValue = "ECDM Message Composer"
 xmlDescription.setAttributeNode(xmlschemasetAtttr)
 
 'provider attribute
 dim xmlproviderAtttr 
 set xmlproviderAtttr = xmlDOM.createAttribute("provider")
 xmlproviderAtttr.nodeValue = "ECDM Message Composer"
 xmlDescription.setAttributeNode(xmlproviderAtttr)
 
 'model attribute
 dim xmlmodelAtttr 
 set xmlmodelAtttr = xmlDOM.createAttribute("model")
 xmlmodelAtttr.nodeValue = Repository.ProjectGUID
 xmlDescription.setAttributeNode(xmlmodelAtttr)
 
 'modelURL attribute
 dim xmlmodelURLAtttr 
 set xmlmodelURLAtttr = xmlDOM.createAttribute("modelURL")
 xmlmodelURLAtttr.nodeValue = ""
 xmlDescription.setAttributeNode(xmlmodelURLAtttr)
 
 'version attribute
 dim xmlversionAtttr 
 set xmlversionAtttr = xmlDOM.createAttribute("version")
 xmlversionAtttr.nodeValue = "13.5.1351.1351"
 xmlDescription.setAttributeNode(xmlversionAtttr)
 
 'xmlns attribute
 dim xmlxmlnsAtttr 
 set xmlxmlnsAtttr = xmlDOM.createAttribute("xmlns")
 xmlxmlnsAtttr.nodeValue = "Der:"
 xmlDescription.setAttributeNode(xmlxmlnsAtttr)
 
 'type attribute
 dim xmltypeAtttr 
 set xmltypeAtttr = xmlDOM.createAttribute("type")
 xmltypeAtttr.nodeValue = "schema"
 xmlDescription.setAttributeNode(xmltypeAtttr)
 
 'auxiliary node
 dim xmlAuxiliary
 set xmlAuxiliary = xmlDOM.createElement( "auxiliary" )
 
 'xmlns attribute
 dim xmlxmlnsAtttrA 
 set xmlxmlnsAtttrA = xmlDOM.createAttribute("xmlns")
 xmlxmlnsAtttrA.nodeValue = ""
 xmlAuxiliary.setAttributeNode(xmlxmlnsAtttrA)
 'add auxiliary node
 xmlDescription.appendChild xmlAuxiliary
 
 'return node
 set createDescriptionNode = xmlDescription
end function

function writefile(filename, contents)
 dim fileSystemObject
 dim outputFile
  
 set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
 set outputFile = fileSystemObject.CreateTextFile(filename, true )
 outputFile.Write contents
 outputFile.Close
end function 

function createTDocument(artifact, xmlString)
 dim timestamp
 timestamp = Year(now()) & "-" & Month(now()) & "-" & Day(now()) & " " & Hour(now()) & ":" & Minute(now) & ":" & Second(now())
 dim sqlCreateSchemaDocument
 sqlCreateSchemaDocument = " INSERT INTO [dbo].[t_document]             " & vbNewLine & _
       "            ([DocID]                        " & vbNewLine & _
       "            ,[DocName]                      " & vbNewLine & _
       "            ,[Notes]                        " & vbNewLine & _
       "            ,[Style]                        " & vbNewLine & _
       "            ,[ElementID]                    " & vbNewLine & _
       "            ,[ElementType]                  " & vbNewLine & _
       "            ,[StrContent]                   " & vbNewLine & _
       "            ,[BinContent]                   " & vbNewLine & _
       "            ,[DocType]                      " & vbNewLine & _
       "            ,[Author]                       " & vbNewLine & _
       "            ,[Version]                      " & vbNewLine & _
       "            ,[IsActive]                     " & vbNewLine & _
       "            ,[Sequence]                     " & vbNewLine & _
       "            ,[DocDate])                     " & vbNewLine & _
       "      VALUES                                " & vbNewLine & _
       "            ('" & CreateGuid() & "'         " & vbNewLine & _
       "            ,'" & artifact.Name & "'        " & vbNewLine & _
       "            ,NULL                           " & vbNewLine & _
       "            ,NULL                           " & vbNewLine & _
       "            ,'" & artifact.ElementGUID & "' " & vbNewLine & _
       "            ,'SC_MessageProfile'            " & vbNewLine & _
       "            ,N'" & xmlString & "'           " & vbNewLine & _
       "            ,NULL                           " & vbNewLine & _
       "            ,'SC_MessageProfile'            " & vbNewLine & _
       "            ,'OCL to Schema Script'         " & vbNewLine & _
       "            ,NULL                           " & vbNewLine & _
       "            ,1                              " & vbNewLine & _
       "            ,0                              " & vbNewLine & _
       "            ,'" & timestamp & "')                "
  Repository.Execute sqlCreateSchemaDocument
end function

private function createArtifact(ownerPackage)
 'add new artifact in owner package
 dim artifact as EA.Element
 set artifact = ownerPackage.Elements.AddNew(ownerPackage.Name & "_Schema", "Artifact")
 artifact.Update
 'save the Schemacomposer property in the Style settings
 Repository.Execute "update t_object set Style = 'MessageProfile=1;' where ea_guid = '" & artifact.ElementGUID & "'"
 set createArtifact = artifact
end function

' Returns a unique Guid on every call. Removes any cruft.
Function CreateGuid()
    CreateGuid = Left(CreateObject("Scriptlet.TypeLib").Guid,38)
End Function

'returns an ArrayList with the elements accordin tot he ObjectID's in the given query
function getElementsFromQuery(sqlQuery)
 dim elements 
 set elements = Repository.GetElementSet(sqlQuery,2)
 dim result
 set result = CreateObject("System.Collections.ArrayList")
 dim element
 for each element in elements
  result.Add Element
 next
 set getElementsFromQuery = result
end function

'get the package id string of the given package tree
function getPackageTreeIDString(package)
 'initialize at "0"
 getPackageTreeIDString = "0"
 dim packageTree
 dim currentPackage as EA.Package
 if not package is nothing then
  'get the whole tree of the selected package
  set packageTree = getPackageTree(package)
  ' get the id string of the tree
  getPackageTreeIDString = makePackageIDString(packageTree)
 end if 
end function

'returns an ArrayList of the given package and all its subpackages recursively
function getPackageTree(package)
 dim packageList
 set packageList = CreateObject("System.Collections.ArrayList")
 addPackagesToList package, packageList
 set getPackageTree = packageList
end function

'add the given package and all subPackges to the list (recursively
function addPackagesToList(package, packageList)
 dim subPackage as EA.Package
 'add the package itself
 packageList.Add package
 'add subpackages
 for each subPackage in package.Packages
  addPackagesToList subPackage, packageList
 next
end function

'make an id string out of the package ID of the given packages
function makePackageIDString(packages)
 dim package as EA.Package
 dim idString
 idString = ""
 dim addComma 
 addComma = false
 for each package in packages
  if addComma then
   idString = idString & ","
  else
   addComma = true
  end if
  idString = idString & package.PackageID
 next 
 'if there are no packages then we return "0"
 if idString = "" then
  idString = "0"
 end if
 'return idString
 makePackageIDString = idString
end function

'test
main
