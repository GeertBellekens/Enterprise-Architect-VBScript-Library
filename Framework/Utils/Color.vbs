'[path=\Framework\Utils]
'[group=Utils]

!INC Local Scripts.EAConstants-VBScript

' From http://www.sparxsystems.com.au/enterprise_architect_user_guide/11/automation_and_scripting/diagramobjects.html
' The color value is a decimal representation of the hex RGB value, where Red=FF, Green=FF00 and Blue=FF0000
' Who would write an RGB as BGR. YAEAB
function SparxColorFromRGB(red, green, blue)
	SparxColorFromRGB = CLng("&h" & blue & green & red)
end function
