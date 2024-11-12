'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Create PUR Diagrams
' Author: Geert Bellekens
' Purpose: Create a diagram at each level, showing the business process a that level with its parent and child processes
' Date: 2018-10-24
'

const outPutName = "Create PUR Diagrams"

sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName

 'get the selected package
 dim selectedPackage as EA.Package
 set selectedPackage = Repository.GetTreeSelectedPackage
 if not selectedPackage is nothing then
  'set timestamp for start
  Repository.WriteOutput outPutName,now() & " Start Create PUR Diagrams"  , 0
  'do the actual work
  addDiagramsToPackage selectedPackage
  'set timestamp for start
  Repository.WriteOutput outPutName,now() & " Finished Create PUR Diagrams"  , 0
 end if
end sub

function addDiagramsToPackage(selectedPackage)
 'only if there are no diagrams in the package
 if selectedPackage.Diagrams.Count = 0 then
  'Add diagram of type Archimate Business Process Diagram
  dim newDiagram as EA.Diagram
  set newDiagram = selectedPackage.Diagrams.AddNew(selectedPackage.Name, "Archimate2::Business")
  if not newDiagram is nothing then
   newDiagram.Update
  end if
  'get the business process elements that should be on this diagram
  addBusinessProcesses newDiagram, selectedPackage
 end if
 'process subPackages
 dim subPackage as EA.Package
 for each subPackage in selectedPackage.Packages
  addDiagramsToPackage(subPackage)
 next
end function



function addBusinessProcesses(newDiagram, selectedPackage)
 dim businessProcesses
 dim businessProces as EA.Element
 dim x
 x = 200
 dim y
 y = 200
 'get the business processes
 set businessProcesses = getBusinessprocesses(selectedPackage)
 for each businessProces in businessProcesses
  'add the business processes to the diagram
  Repository.WriteOutput outPutName,now() & " Adding business process '" & businessProces.Name & "' to diagram '" & newDiagram.Name & "'"    , 0
  addElementToDiagram businessProces, newDiagram, y, x
  x = x + 10
  y = y + 10
  'if the business process is part of this diagram then we set the composite diagram
  if businessProces.PackageID = selectedPackage.PackageID then
   setCompositeDiagram businessProces, newDiagram
  end if
 next
 'layout the diagram
 dim XMLdiagramID
 XMLdiagramID = Repository.GetProjectInterface().GUIDtoXML(newDiagram.DiagramGUID)
    Repository.GetProjectInterface().LayoutDiagramEx XMLdiagramID,lsDiagramDefault, 4, 20, 20, false
 'save the diagram
 newDiagram.Update
 'close the diagram
 Repository.CloseDiagram newDiagram.DiagramID
end function

function getBusinessprocesses(selectedPackage)
 dim sqlGetBusinessProcesses
 sqlGetBusinessProcesses = "select o.Object_ID from t_object o                                  " & vbNewLine & _
       " where o.Package_ID = " & selectedPackage.PackageID &"                " & vbNewLine & _
       " and o.Stereotype = 'ArchiMate_BusinessProcess'                       " & vbNewLine & _
       " union                                                                " & vbNewLine & _
       " select o2.Object_ID from t_object o                                  " & vbNewLine & _
       " inner join t_connector c on o.Object_ID = c.Start_Object_ID          " & vbNewLine & _
       "        and c.Stereotype = 'ArchiMate_Composition' " & vbNewLine & _
       " inner join t_object o2 on o2.Object_ID = c.End_Object_ID             " & vbNewLine & _
       " where o.Package_ID = " & selectedPackage.PackageID &"                " & vbNewLine & _
       " and o.Stereotype = 'ArchiMate_BusinessProcess'                       " & vbNewLine & _
       " union                                                                " & vbNewLine & _
       " select o2.Object_ID from t_object o                                  " & vbNewLine & _
       " inner join t_connector c on o.Object_ID = c.End_Object_ID            " & vbNewLine & _
       "        and c.Stereotype = 'ArchiMate_Composition' " & vbNewLine & _
       " inner join t_object o2 on o2.Object_ID = c.Start_Object_ID           " & vbNewLine & _
       " where o.Stereotype = 'ArchiMate_BusinessProcess'                     " & vbNewLine & _
       " and o.Package_ID = " & selectedPackage.PackageID
 set getBusinessprocesses = getElementsFromQuery(sqlGetBusinessProcesses)
end function

main