option explicit

'[path=\Framework\Utils]
'[group=Utils]

!INC Local Scripts.EAConstants-VBScript
!INC Logging.LogManager
!INC Assert.Assert
!INC Utils.Color

dim logger
set logger = LogManager.getLogger("Util.Color Test")

sub TestSparxColorToRGB
	dim rgb
	rgb = SparxColorToRgbColor(15138790)
	assertEquals "rgb(230,255,230) = #E6FFE6 = SparxColor(15138790)", 230, rgb(0)
	assertEquals "rgb(230,255,230) = #E6FFE6 = SparxColor(15138790)", 255, rgb(1)
	assertEquals "rgb(230,255,230) = #E6FFE6 = SparxColor(15138790)", 230, rgb(2)
end sub

sub TestRgbColorToSparxColor
	' https://convertacolor.com/#/hex/69BE63
	assertEquals "rgb(230,255,230) = #E6FFE6 = SparxColor(15138790)", 15138790, RgbColorToSparxColor(230, 255, 230)
end sub

sub TestHexColorToSparxColor
	' https://convertacolor.com/#/hex/69BE63
	assertEquals "rgb(230,255,230) = #E6FFE6 = SparxColor(15138790)", 15138790, HexColorToSparxColor(&HE6FFE6)
end sub

sub TestSparxColorToHexColor
	' https://convertacolor.com/#/hex/69BE63
	assertEquals "rgb(230,255,230) = #E6FFE6 = SparxColor(15138790)", &HE6FFE6, SparxColorToHexColor(15138790)
end sub

sub main
	TestSparxColorToRGB
	TestSparxColorToHexColor
	TestRgbColorToSparxColor
	TestHexColorToSparxColor
	logger.INFO "Tests Completed"
end sub

main