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
 integraties = Repository.GetPackageByGuid("{999E603A-E75B-4f07-A013-87D44B5BBC41}");
 Session.Output(integraties.Name);
 for (i = 0; i < integraties.Elements.Count; i++) {
  element = integraties.Elements.GetAt(i);
//  element.StereotypeEx = "Integratie,ArchiMate_ApplicationInteraction";
  element.StereotypeEx = "ArchiMate_ApplicationInteraction";
  Session.Output(element.Name);
  element.Update();
 }
}

main();
