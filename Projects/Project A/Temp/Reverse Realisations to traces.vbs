'[path=\Projects\Project A\Temp]
'[group=Temp]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Reverse Realisations to traces
' Author: Geert Bellekens
' Purpose: Finds all Realisation relationships in the selected package tree and converts them to traces in the opposite direction
' Date: 2019-02-12
'
sub main
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage()
	'get the connectors
	dim sqlGetConnectors
	sqlGetConnectors = "select c.Connector_ID from t_object o                                                " & _
						" inner join t_connector c on c.Start_Object_ID = o.Object_ID                        " & _
						" 							and c.Connector_Type in ('Realisation', 'Realization')   " & _
						" where o.Package_ID in (" & getPackageTreeIDString(selectedPackage) &")                                                           "
	
	dim connectors
	set connectors = getConnectorsFromQuery(sqlGetConnectors)
	dim connector as EA.Connector
	for each connector in connectors
		'switch source and target
		dim temp
		temp = connector.ClientID
		connector.ClientID = connector.SupplierID
		connector.SupplierID = temp
		'set to trace
		connector.Type = "Abstraction"
		connector.Stereotype = "trace"
		connector.Update
	next	
end sub

main