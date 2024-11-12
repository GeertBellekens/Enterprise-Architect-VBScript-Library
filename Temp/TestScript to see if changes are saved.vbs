'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

sub main
 dim m_ExcelApp
 Session.Output "starting createObject excel created"
 set m_ExcelApp = CreateObject("Excel.Application")
 Session.Output "excel created"
 m_ExcelApp.Visible = true
 m_ExcelApp.Quit
end sub
'blabla, and more blabla

main