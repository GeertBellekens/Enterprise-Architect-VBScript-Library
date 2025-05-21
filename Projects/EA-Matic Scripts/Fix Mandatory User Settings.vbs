'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Fix Mandatory User Settings
' Author: Geert Bellekens
' Purpose: Check the mandatory user settings in the registry and set them correctly if needed
' Date: 2021-05-28
'

'EA-Matic

const REG_SZ = "REG_SZ"
const REG_DWORD = "REG_DWORD"
const REG_BINARY = "REG_BINARY"

function fixSettings
	dim regPath
	Dim regkey
	dim regValue
	dim existingValue
	'place in the registry that contains all of the user settings
	regPath = "HKEY_CURRENT_USER\Software\Sparx Systems\EA400\EA\OPTIONS\"
	'get the EA version
	'set the default diagram layout
	Repository.Execute  "update s set s.[Value] = 'l=20;c=20;d=1;cr=1;la=2;i=1;it=4;a=0;' from usys_system s where s.[Property] = 'Diagram_Layout'"
	'fix registry settings
	dim settingsValid
	settingsValid = true
	settingsValid = settingsValid AND validateRegValue(regPath, "SORT_FEATURES","1", REG_DWORD) 'commented out because of conflicts with TMF model. Has to be uncommented for Baloise
	settingsValid = settingsValid AND validateRegValue(regPath, "TREE_SORT","0", REG_DWORD)
	settingsValid = settingsValid AND validateRegValue(regPath, "XMI_ReportXrefDeletion","0", REG_DWORD)
	settingsValid = settingsValid AND validateRegValue(regPath, "JET4","1", REG_DWORD)
	validateRegValue regPath, "XSDFileName","", REG_SZ
		
	if not settingsValid then
		msgbox "Mandatory user settings have been corrected." & vbNewLine & "Please restart EA",vbOKOnly+vbExclamation,"Corrected mandatory user settings!" 
		Repository.Exit
	end if
		
end function

function validateRegValue(regPath, regKey, regValue, regType)
	Dim shell
	' Create the Shell object
	Set shell = CreateObject("WScript.Shell")
	dim existingValue
	on error resume next
	'read registry value
	existingValue = shell.RegRead(regPath & regkey)
	'if the key doesn't exist then RegRead throws an error
	If Err.Number <> 0 Then
		existingValue = ""
		Err.Clear
	end if
	on error goto 0
	'check the value in the registry with the desired value
	if Cstr(existingValue) <> regValue then
		'write the correct value to the registry
		shell.RegWrite regPath & regkey, regValue, regType
		'return false
		validateRegValue = false
	else
		'value was already OK, return true
		validateRegValue = true
	end if
end function

function EA_FileOpen()
	 fixSettings
end function
