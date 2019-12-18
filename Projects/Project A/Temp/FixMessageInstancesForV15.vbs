'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: FixMessageInstancesForV15
' Author: Geert Bellekens
' Purpose: Fix the message instances for v15 by updating their t_xref, making sure the stereotype only reeds
' Date: 2019-10-02
'

'name of the output tab
const outPutName = "Fix Message Instances"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if selectedPackage is nothing then
		msgbox "Please select a package in the project browser before running this script",vbOKOnly+vbExclamation,"No package selected!"
		exit sub
	end if
	'get confirmation
	dim userInput
	userinput = MsgBox( "Fix message instances for package '"& selectedPackage.Name &"'?", vbYesNo + vbQuestion, "Fix message instances?")
	'save the schema content
	if userinput = vbYes then
		'report progress
		Repository.WriteOutput outPutName, now() & " Starting Fix message instances for '"& selectedPackage.Name &"'", 0
		'actually fix the message instances
		fixMessageInstances selectedPackage
		'report progress
		Repository.WriteOutput outPutName, now() & " Finished Fix message instances for '"& selectedPackage.Name &"'", 0
	end if 
end sub

function fixMessageInstances (package)
	'get the package tree id's
	dim packageTreeIDs
	packageTreeIDs = getPackageTreeIDString(package)
	dim sqlUpdateStereotypes
	sqlUpdateStereotypes = "update x set x.Description = '@STEREO;Name=Message;GUID={BD861500-3029-4bda-97A1-BC677549C506};@ENDSTEREO;'                " & vbNewLine & _
						" from t_object o                                                                                                            " & vbNewLine & _
						" inner join t_xref x on x.Client = o.ea_guid                                                                                " & vbNewLine & _
						" 				and x.Name = 'Stereotypes'                                                                                   " & vbNewLine & _
						" where o.Object_Type = 'Object'                                                                                             " & vbNewLine & _
						" and CONVERT(varchar(max), x.Description) = '@STEREO;Name=Message;GUID={BD861500-3029-4bda-97A1-BC677549C506};FQName=BPMN2.0::Message;@ENDSTEREO;' " & vbNewLine & _
						" and o.Package_ID in (" & packageTreeIDs & ")                                                                               "
	'execute the update query
	Repository.Execute sqlUpdateStereotypes
	'reload package if v15 or higher
	if Repository.LibraryVersion >= 1500 then
		Repository.ReloadPackage package.PackageID
	end if
end function

main