'[path=\Projects\Project A\Project Browser Package Group]
'[group=Project Browser Package Group]

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Remove Attribute Type Dependencies
' Author: Geert Bellekens
' Purpose: This script will remove all attribute type dependencies created by the EA Message Composer
'    script can be executed on a package an will remove the dependencies from the selected package and all sub packages.
		
' Date: '2017-04-20
'
sub main

	dim messagePackage as EA.Package
	set messagePackage = Repository.GetTreeSelectedPackage()
	' ask the user for confirmation before updating the tags just in case something went wrong with the generation of the subset and we have the wrong diagram
	dim response
	response = msgbox("Remove all attribute type dependencies in '" & messagePackage.Name & "'?", vbYesNo+vbQuestion, "Remove attribute type dependencies?")
	if response = vbYes then
		dim packageIDString
		packageIDString = getPackageTreeIDString(messagePackage)
		dim sqlDeleteDependencies
		sqlDeleteDependencies = " delete t_connector where Connector_ID in                   " & _                            
								" (select c.Connector_ID from ((t_connector c      			 " & _	
								" inner join t_object so on c.Start_Object_ID = so.Object_ID)" & _
								" inner join t_attribute a on (a.Object_ID = so.Object_ID    " & _
								" 						and a.Classifier = c.End_Object_ID)) " & _
								" where c.Connector_Type = 'dependency'                      " & _
								" and so.Package_ID in (" & packageIDString & "))            "
		Repository.Execute sqlDeleteDependencies
		msgbox "Finished!"
	end if
	
end sub

main