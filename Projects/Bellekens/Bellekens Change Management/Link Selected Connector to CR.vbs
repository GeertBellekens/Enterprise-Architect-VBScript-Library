'[path=\Projects\Bellekens\Bellekens Change Management]
'[group=Link Group]
option explicit

!INC Bellekens Change Management.LinkToCRMain

'This script only calls the function defined in the main script.
'Ths script is to be copied in the appropriate groups

'Execute main function defined in LinkToCRMain
sub main
 dim selectedItem
 set selectedItem = Repository.GetContextObject
 linkItemToCR selectedItem, nothing
end sub

main