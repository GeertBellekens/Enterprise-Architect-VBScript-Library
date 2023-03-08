## VBScript Template

```
option explicit
'[path=\path\to\script]
'[group=group name]

'!INC Group.Subject

const tabName = "Subject"

sub main
       'create output tab
       Repository.CreateOutputTab tabName
       Repository.ClearOutput tabName
       Repository.EnsureOutputVisible tabName

       'set timestamp
       Repository.WriteOutput tabName, now() & " Starting Subject", 0

       '''gather required details

       'for example: get selected package
       dim package as EA.Package
       set package = Repository.GetTreeSelectedPackage()

       '''validate details and provide error message

       'exit if not selected
       if package is nothing then
              msgbox "Please select a package before running this script"
              exit sub
       end if

       '''Delegate actual work to included script
       doSubject(package)

       'set timestamp
       Repository.WriteOutput tabName, now() & " Finished Subject", 0

end sub

main
```
