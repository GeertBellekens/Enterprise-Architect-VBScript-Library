'[path=\Projects\Project B\Model Management]
'[group=Model Management]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Remove Checkout-offline
' Author: Geert Bellekens
' Purpose: Removes the flag "Checkeout offline" from the selected package
' Date: 2019-04-03
'
const outPutName = "Remove Checkout Offline"

sub main
	'create output tab
	Repository.CreateOutputTab outPutName
	Repository.ClearOutput outPutName
	Repository.EnsureOutputVisible outPutName
	'get selected package
	dim selectedPackage as EA.Package
	set selectedPackage = Repository.GetTreeSelectedPackage
	if not selectedPackage is nothing then
		'ask for confirmation
		dim userIsSure
		userIsSure = Msgbox("Are you sure you want to remove the checkout-offline from the package '" &selectedPackage.Name & "' '", vbYesNo+vbExclamation, "Remove Checkout-offline" )
		if userIsSure = vbYes then
			'Repository.WriteOutput outPutName, now() & " Starting delete package tree for package '"& selectedPackage.Name &"'", 0
			'delete the package using it's parent
			removeCheckoutOffline selectedPackage
'			'refresh
'			Repository.RefreshModelView 0
			'let user know
			Repository.WriteOutput outPutName, now() & " Removed checkout-offline from '"& selectedPackage.Name &"'", 0
		end if
	end if
end sub

function removeCheckoutOffline(package)
	dim sqlUpdate
	sqlUpdate = "update p set p.PackageFlags = LEFT (p.PackageFlags, patindex('%CheckedOutOffline=1%',p.PackageFlags)-1)    " & vbNewLine & _
				" from t_package p                                                                                          " & vbNewLine & _
				" where p.PackageFlags like '%CheckedOutOffline=1%'                                                         " & vbNewLine & _
				" and p.ea_guid = '" & package.PackageGUID & "'                                                             "
	Repository.Execute sqlUpdate
	'Reload package
	Repository.ReloadPackage package.PackageID
end function

main