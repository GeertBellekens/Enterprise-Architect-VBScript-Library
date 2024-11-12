'[group=De Lijn - Element Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC De Lijn Scripts.Genereer Proceskaart - Algemeen

'***************************************************************************************************
' Author:   Alain Van Goethem
' Purpose:   Generate a virtual document of "Proceskaart" based on select BusinessProcess
' Creation Date: 04/10/2018
'***************************************************************************************************

sub main()
 Session.Output "*******************************************"
 Session.Output "Initiating script 'Genereer Proceskaart'..."
 
 ' Get a reference to the selected element
 dim currentElement as EA.Element
 set currentElement = Repository.GetContextObject()
  
 if not currentElement is nothing then
  'Session.Prompt "BusinessProcess found", promptOK
  createBPDocument(currentElement)
  Session.Prompt "Proceskaart gegenereerd.", promptOK
 else
  Session.Output "Error - No Business Process element selected."
  Session.Prompt "Gelieve een element van type BusinessProcess te selecteren.", promptOK
 end if

 Session.Output "Done"
end sub

main
