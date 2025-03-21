'[path=\Projects\Project B\Package Group]
'[group=Package Group]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
!INC Baloise Scripts.MessageDetailsMain

' Script Name: Export Message Content 
' Author: Geert Bellekens
' Purpose: Export the contents of the messages in the selected package to excel.
' Date: 2018-09-03
'


const outPutName = "Export Message Content"


sub main()
	getMessageDetailsMain regularMessageContent
end sub


main