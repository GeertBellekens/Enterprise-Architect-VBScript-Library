'[path=\Projects\Project DL\Template fragments]
'[group=Template fragments]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: DMNDecisionTable
' Author: Geert Bellekens
' Purpose: Return the decision table for the given element, to be used by a script fragment
' Date: 2021-09-24
'

function MyRtfData (objectID )
 
 dim xmlDOM 
 set  xmlDOM = CreateObject( "Microsoft.XMLDOM" )
 'set  xmlDOM = CreateObject( "MSXML2.DOMDocument.4.0" )
 xmlDOM.validateOnParse = false
 xmlDOM.async = false
  
 dim node 
 set node = xmlDOM.createProcessingInstruction( "xml", "version='1.0'")
    xmlDOM.appendChild node
'
 dim xmlRoot 
 set xmlRoot = xmlDOM.createElement( "EADATA" )
 xmlDOM.appendChild xmlRoot

 dim xmlDataSet
 set xmlDataSet = xmlDOM.createElement( "Dataset_0" )
 xmlRoot.appendChild xmlDataSet
  
 dim xmlData 
 set xmlData = xmlDOM.createElement( "Data" )
 xmlDataSet.appendChild xmlData
 
 'create the for this decision table
 createRows xmlDOM, xmlData, objectID
 
 MyRtfData = xmlDOM.xml
end function

function createRows(xmlDOM, xmlData,objectID)
 'get element
 dim decision as EA.Element
 set decision = Repository.GetElementByID(objectID)
 'get decisionLogic
 dim decisionLogic
 decisionLogic = getDecisionLogic(decision)
 'load xML
 dim xmlDecision 
 set xmlDecision = CreateObject("MSXML2.DOMDocument")
 If xmlDecision.LoadXML(decisionLogic) Then
  dim headers
  set headers = getHeaders(xmlDecision)
  createRow  xmlDOM, xmlData, headers
 end if
 dim rules
 set rules = getRules(xmlDecision)
 dim rule
 for each rule in rules
  createRow  xmlDOM, xmlData, rule
 next
end function

function getRules(xmlDOM)
 dim rules
 set rules = CreateObject("System.Collections.ArrayList")
 dim ruleNodes
 set ruleNodes = xmlDOM.SelectNodes("//dmn:rule")
 dim ruleNode
 dim i
 i = 0
 for each ruleNode in ruleNodes
  dim rule
  set rule = CreateObject("System.Collections.ArrayList")
  'add U
  i = i + 1
  rule.add i
  dim textNodes
  set textNodes = ruleNode.SelectNodes(".//dmn:text")
  dim textNode
  for each textNode in textNodes
   rule.add textNode.Text
  next
  rules.add rule
 next
 'return
 set getRules = rules
end function

function getHeaders(xmlDOM)
 dim headers
 set headers = CreateObject("System.Collections.ArrayList")
 'add the U column
 headers.add "U"
 dim childNode
 for each childNode in xmlDOM.documentElement.ChildNodes
  if childNode.NodeName = "dmn:input" then
   'get label
   headers.add childNode.getAttribute("label")
  elseif childNode.NodeName = "dmn:output" then
   'get name
   headers.add childNode.getAttribute("name")
  end if
 next
 'return
 set getHeaders = headers
end function

function createRow (xmlDOM, xmlData, rowData)
 'create row in xml
 dim xmlRow
 set xmlRow = xmlDOM.createElement( "Row" )
 xmlData.appendChild xmlRow
 dim field
 dim i
 i = 0
 'Add fields
 for each field in rowData
  dim xmlField
  if i = 0 then
   set xmlField = xmlDOM.createElement( "U" ) 
  else
   set xmlField = xmlDOM.createElement( "Field" & i ) 
  end if
  xmlField.text = field
  xmlRow.appendChild xmlField
  i = i + 1
 next
end function

function getDecisionLogic(element)
 dim tag as EA.TaggedValue
 dim decisionLogic
 decisionLogic = ""
 for each tag in element.TaggedValues
  if lcase(tag.Name) = "decisionlogic" then
   decisionLogic = tag.Notes
   exit for
  end if
 next
 'return logic
 getDecisionLogic = decisionLogic
end function

'msgbox MyPackageRtfData(3357,"")
function test
 dim outputString
 dim fileSystemObject
 dim outputFile
 
 dim element as EA.Element
 set element = Repository.GetElementByGuid("{9F3E8E21-1703-45de-BBB9-3DE29A383448}")
 
 outputString =  MyRtfData(element.ElementID)
 
 set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
 set outputFile = fileSystemObject.CreateTextFile( "H:\temp\decisionTest.xml", true )
 outputFile.Write outputString
 outputFile.Close
end function 

'test