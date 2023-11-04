'[path=\Framework\Utils]
'[group=Utils]

!INC Local Scripts.EAConstants-VBScript

' Sparx "default" color to reset color back to diagram object default colors
dim sparxDefaultColor
sparxDefaultColor 					= -1

' Convert a (red, green, blue) color (for example, (230, 255, 230)) into a Sparx Color 
' 
' From http://www.sparxsystems.com.au/enterprise_architect_user_guide/11/automation_and_scripting/diagramobjects.html
' The color value is a decimal representation of the hex RGB value, where Red=FF, Green=FF00 and Blue=FF0000
' See https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/676k6dt6(v=vs.84)
function RgbColorToSparxColor(red, green, blue)
	RgbColorToSparxColor = RGB(red, green, blue)
end function

' Convert a Sparx color value into an array[red, green, blue] 
' the values of red, green, blue are decimal values from 0x00 to 0xFF
function SparxColorToRgbColor(color)
	dim red, green, blue

	red = color mod 16^2
	green = (color \ (16^2)) mod 16^2
	blue = (color \ (16^4)) mod 16^2

	SparxColorToRgbColor = Array(red, green, blue)
end function

' Convert a hex color value (for examnple, &HE6FFE6) into a SparxColor
function HexColorToSparxColor(hexValue)
	dim red, green, blue

	blue = hexValue mod 16^2
	green = (hexValue \ (16^2)) mod 16^2
	red = (hexValue \ (16^4)) mod 16^2
	
	HexColorToSparxColor = RGB(red, green, blue)
end function

' Convert a SparxColor into a hex color value
function SparxColorToHexColor(color)
	dim rgb
	rgb = SparxColorToRgbColor(color)
	
	SparxColorToHexColor = CLng("&H" & HEX(rgb(0)) & HEX(rgb(1)) & HEX(rgb(2)))
end function

