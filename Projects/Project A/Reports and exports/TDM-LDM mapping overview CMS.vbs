'[path=\Projects\Project A\Reports and exports]
'[group=Reports and exports]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Reports and exports.TDM-LDM main

'
' Script Name: TDM-LDM mapping overview CMS
' Author: Geert Bellekens
' Purpose: get an overview of the mapping between TDM and LDM, including the non mapped classes and attributes
' Date: 2019-04-24
'
sub main
	dim TDMPackageGUID
	TDMPackageGUID = "{BA787F2C-3E47-43b5-9A4A-280D5199FB30}" 'TDM CMS
	dim LDMPackageGUID 
	LDMPackageGUID = "{0C10A9E8-3B77-42f2-B989-E4FA1F97F5F9}" 'LDM CMS
	'call main function
	GenerateTDMLDMOverview TDMPackageGUID, LDMPackageGUID
end sub

main