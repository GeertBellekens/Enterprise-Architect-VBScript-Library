//[group=Tims Scripts]
!INC Local Scripts.EAConstants-JScript

/*
 * Script Name: 
 * Author: 
 * Purpose: 
 * Date: 
 */
 
function main()
{
 appComps = Repository.GetPackageByGuid("{89CC255F-72C3-486f-90C4-20D45E5AD61E}");
 Session.Output(appComps.Name);
 for (i = 0; i < appComps.Elements.Count; i++) {
  element = appComps.Elements.GetAt(i);
//  element.StereotypeEx = "Applicatie,ArchiMate_ApplicationComponent";
  element.StereotypeEx = "ArchiMate_ApplicationComponent";
  Session.Output(element.Name);
  for (j = 0; j < element.Elements.Count; j++) {
   subElement = element.Elements.GetAt(j);
//   subElement.StereotypeEx = "Applicatie,ArchiMate_ApplicationComponent";
   subElement.StereotypeEx = "ArchiMate_ApplicationComponent";
   Session.Output(subElement.Name);
   subElement.Update();
  }
  element.Update();
 }
}

main();