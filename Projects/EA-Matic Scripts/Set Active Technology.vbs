'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]

'option explicit
'
'!INC Local Scripts.EAConstants-VBScript

'
' Script Name: ActivateTechnology
' Author: Geert Bellekens
' Purpose: Sets the given technology as the "Active" technology
' Date: 2020-03-31
'
'EA-Matic

function EA_FileOpen()
	Repository.ActivateTechnology "TRD"
end function