//[group=Project Browser]
!INC Local Scripts.EAConstants-JScript

// Author: Jelle Hellemans (jelle.hellemans@scarlet.be) (v1.2)

function OnProjectBrowserScript()
{
 ClearOutput ("Script");
 
 var name = Session.Input("Go to Node Path:");
 
 if (name == '') {
  return;
 }

 var EOrA = ["Packages", "Elements"];
 name = name.replace(/(\w|\))\.(\w|\()/g, '\$1&££&\$2');
 Session.Output(name);
 var split = name.split("&££&");
 var current = Repository.Models.GetByName(split[0]);
 for (var i = 1; i < split.length-1; i++) {
  var part = split[i];
  Session.Output("at: "+part);
  var res = null;
  for (var j = 0; j < EOrA.length; j++) {
   try {
    res = current[EOrA[j]].GetByName(part);
    if (res) {
     break;
    }
   } catch (e) {
    Session.Output("warning: "+e.message);
   }
  }
  if (res == null) {
   Repository.EnsureOutputVisible( "Script" );
   Session.Output("not found!");
   return;
  }
  current = res;
 }

 var result = null;
 var stuff = ["Elements", "Packages", "Diagrams"];
 for(var i = 0; i < stuff.length; i++) {
  try {
   result = current[stuff[i]].GetByName(split[split.length-1]);
   if (result) {
    break;
   }
  } catch (e) {
   Session.Output("warning: "+e.message);
  } 
 }
 if (result == null) {
  Repository.EnsureOutputVisible( "Script" );
  Session.Output("not found!");
  return;
 }
 Session.Output("found "+result.Name+" "+result.ObjectType);
 Repository.ShowInProjectView(result);
 
 if (result.ObjectType == otDiagram) {
  Repository.OpenDiagram(result.DiagramID);
 }
}
OnProjectBrowserScript();