'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript


sub main
	Repository.Execute "update t_diagram set locked = 0 where locked = 1"
end sub

main