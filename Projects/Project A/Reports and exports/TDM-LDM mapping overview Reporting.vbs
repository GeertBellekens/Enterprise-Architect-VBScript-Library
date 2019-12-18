'[path=\Projects\Project A\Reports and exports]
'[group=Reports and exports]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Reports and exports.TDM-LDM main

'
' Script Name: Reporting TDM-LDM mapping overview
' Author: Geert Bellekens
' Purpose: get an overview of the mapping between TDM and LDM, including the non mapped classes and attributes
' Date: 2019-04-24
'
sub main
 dim TDMPackageGUID
 TDMPackageGUID = "{60506E74-4E10-48a1-8D21-9C000DFF5702}" 'BI_View_Table
 dim LDMPackageGUID 
 LDMPackageGUID = "{B460AE6A-19A2-42ed-BD43-0221BA94B3B9}" 'LDM DW
 'call main function
 GenerateTDMLDMOverview TDMPackageGUID, LDMPackageGUID
end sub

main