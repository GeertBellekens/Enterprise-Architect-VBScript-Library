'[path=\Projects\EA-Matic Scripts]
'[group=EA-Matic]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

function EA_OnPostNewConnector(Info)
	 'Add code here
	 msgbox "gelukt met includes wan getWC() =" & getWC()
end function