'[path=\Projects\Project A\Template fragments]
'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Atrias Scripts.Util
'
sub main
 MyRtfData("5209")
end sub


function MyRtfData(DiagramID)
 dim diagram as EA.Diagram
 set diagram = Repository.GetDiagramByID(DiagramID)
 dim trans as EA.Connector
 dim source as EA.Element
 dim target as EA.Element
 dim diagramLink as EA.DiagramLink
 for each diagramLink in diagram.DiagramLinks
  set trans = Repository.GetConnectorByID(diagramLink.ConnectorID)
  set source = Repository.GetElementByID(trans.ClientID)
  set target = Repository.GetElementByID(trans.SupplierID)
  Session.Output source.Name & " - " & target.Name & " - " & trans.TransitionGuard
  
  'get the trigger names and specifications
  
  
 next
 
 
 
 
 'set states = Repository.getE
 
end function

main