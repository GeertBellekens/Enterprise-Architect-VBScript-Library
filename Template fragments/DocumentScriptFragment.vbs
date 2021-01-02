'[group=Template fragments]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: DocumentScriptFragment
' Author: Geert Bellekens
' Purpose: test a document script fragment by returning raw RTF
' Date: 2020-11-27
'
function MyRTFData()
	MyRTFData = "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1036{\fonttbl{\f0\fnil\fcharset0 Calibri;}}" & vbNewLine & _
				"{\*\generator Riched20 10.0.18362}\viewkind4\uc1" & vbNewLine & _
				"\pard\sa200\sl276\slmult1\f0\fs22\lang9\par" & vbNewLine & _
				"\par" & vbNewLine & _
				"line 1\par" & vbNewLine & _
				"line 2\par" & vbNewLine & _
				"\par" & vbNewLine & _
				"}" & vbNewLine
end function

