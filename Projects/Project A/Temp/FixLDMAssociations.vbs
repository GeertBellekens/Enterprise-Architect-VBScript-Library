'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: FixLDMAssociations
' Author: Geert Bellekens
' Purpose: Fixes the LDM associations by settign the navigability and removing the direction arrows.
' Date: 2018-08-29
'

'name of the output tab
const outPutName = "Fix LDM associatons"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'set timestamp for start
	Repository.WriteOutput outPutName,now() & " Starting to fix the LDM associations"  , 0
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	'keep track of all processed associations
	dim processedAssociations
	set processedAssociations = CreateObject("Scripting.Dictionary")
	'fix associations for selected package
	fixAssocations selectedPackage, processedAssociations
	'let the user know we are finished
	Repository.WriteOutput outPutName,now() & " Finished! Fixed " & processedAssociations.Count & " associations"  , 0
end sub

function fixAssocations(selectedPackage, processedAssociations)
	dim element as EA.Element
	'process owned elements
	for each element in selectedPackage.Elements
		if element.Type = "Class" then
			'inform user of progress
			Repository.WriteOutput outPutName,now() & " Processing associations of '" & element.Name & "'"  , 0
			dim slqGetAssociations
			slqGetAssociations = " select c.Connector_ID from t_connector c                 " & _
								" inner join t_object o on c.End_Object_ID = o.Object_ID   " & _
								" where c.Connector_Type in ('Association', 'Aggregation') " & _
								" and o.Object_Type = 'Class'                              " & _
								" and c.Start_Object_ID = " & element.ElementID
			dim associations 
			set associatons = getConnectorsFromQuery(slqGetAssociations)
			dim connector as EA.Connector
			for each connector in associatons
				if not processedAssociations.Exists(connector.ConnectorGUID) then
					'remove direction arrow
					dim sqlUpdateArrows
					 sqlUpdateArrows = "update t_diagramlinks set Geometry = replace (convert(varchar(max),Geometry), 'DIR=1','DIR=0') " & _
									   " where Geometry like '%DIR=1%' " & _
									   " and ConnectorID =" & connector.ConnectorID
					Repository.Execute sqlUpdateArrows
					'set navigability
					if connector.Direction <> "Source -> Destination" then
						connector.Direction = "Source -> Destination"
						connector.Update
					end if
					'store in processed associations
					processedAssociations.Add connector.ConnectorGUID, connector
				end if
			next
		end if
	next
	'process subpackages
	dim subPackage
	for each subPackage in selectedPackage.Packages
		fixAssocations subPackage, processedAssociations
	next
end function

main