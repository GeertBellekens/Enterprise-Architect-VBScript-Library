//[group=Temp]
!INC Local Scripts.EAConstants-JScript

function main()
{
	Session.Output(Repository.ConnectionString);
	
	var test = Repository.GetElementByID(222);
	Session.Output(test.name);
}

main();