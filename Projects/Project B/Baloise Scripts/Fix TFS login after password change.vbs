'[path=\Projects\Project B\Baloise Scripts]
'[group=Baloise Scripts]

option explicit


!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Fix TFS login error after password change
' Author: Geert Bellekens
' Purpose: Remove cached credentials for TFS command line utility
' Error:
'   Error while initializing Version Control provider:

'   TFS reports the following error:
' Date: 2023-11-14
'

Const HKEY_CURRENT_USER = 2147483649

sub main
    Dim shell
    ' We need to delete the Token registry entry under the path
    ' HKEY_CURRENT_USER\SOFTWARE\Microsoft\VSCommon\14.0\ClientServices\TokenStorage\VisualStudio\VssApp\32d4baa406d0478cb61b00666d6aeb46\Token
    ' but the subkey containing the token is a random value. Therefore we first get the subkey name, and then delete the Token value.
    ' this deletes the cached credentials, and forces a login prompt when opening an EA model that is linked to TFS.
    Set shell = CreateObject("WScript.Shell")
    dim subKey
    subkey = getSubKeyName()
    shell.RegDelete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\VSCommon\14.0\ClientServices\TokenStorage\VisualStudio\VssApp\"& subkey & "\Token"
    msgbox "Please reload EA to trigger the TFS login again"
end sub

function getSubKeyName
    getSubKeyName = ""
    Dim reg  
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    dim subKeys
    dim subKey
    reg.EnumKey HKEY_CURRENT_USER, "SOFTWARE\Microsoft\VSCommon\14.0\ClientServices\TokenStorage\VisualStudio\VssApp", subKeys
    ' Loop through each key
    For Each subKey In subKeys
        getSubKeyName = subKey
        exit for
    Next
end function