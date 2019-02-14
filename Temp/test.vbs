'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

sub main
	getEAAddingFrameworkModel()
end sub

'gets a new instance of the EAAddinFramework and initializes it with the EA.Repository
function getEAAddingFrameworkModel()
	'Initialize the EAAddinFramework model
    dim model 
    set model = CreateObject("TSF.UmlToolingFramework.Wrappers.EA.Model")
    model.initialize(Repository)
	set getEAAddingFrameworkModel = model
end function

main