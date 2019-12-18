'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Fix Mandatory User Settings
' Author: Geert Bellekens
' Purpose: Check the mandatory user settings in the registry and set them correctly if needed
' Date: 2019-11-05
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
	dim eaVersion
	eaVersion = Repository.LibraryVersion
	
	dim settingsValid
	settingsValid = true
	'Fontname13 is only relevant for V15
	if eaVersion > 1300 then
		settingsValid = settingsValid AND validateRegValue(regPath, "FONTNAME13","Arial", REG_SZ)
	else
		settingsValid = settingsValid AND validateRegValue(regPath, "FONTNAME","Arial", REG_SZ)
	end if
	settingsValid = settingsValid AND validateRegValue(regPath, "SAVE_CLIP_FRAME","1", REG_DWORD)
	settingsValid = settingsValid AND validateRegValue(regPath, "PRINT_IMAGE_FRAME","1", REG_DWORD)
	settingsValid = settingsValid AND validateRegValue(regPath, "SAVE_IMAGE_FRAME","1", REG_DWORD)
	settingsValid = settingsValid AND validateRegValue(regPath, "SORT_FEATURES","0", REG_DWORD)
	settingsValid = settingsValid AND validateRegValue(regPath, "ALLOW_DUPLICATE_TAGS","1", REG_DWORD)
	
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