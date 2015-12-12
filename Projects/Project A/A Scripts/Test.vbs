'[path=\Projects\Project A\A Scripts]
'[group=Atrias Scripts]
option explicit

sub main
	dim sqlUpdate
	sqlUpdate = "update t_attribute set name = 'CorrectName' where name = 'WrongName' and stereotype = 'column'"
	Repository.Execute sqlUpdate
end sub

main