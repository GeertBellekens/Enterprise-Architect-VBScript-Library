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
// Session.Output(integraties.Name);
 i = 1;
 found = false;
 IL = 0;
 while (i <= integraties.elements.Count && !found) {
  element = integraties.elements.GetAt(i-1);
  IL = element.Name.substring(3,6);
  if (i != IL) {
   found = true;
  }
  i++;
 }
 Session.Prompt(i-1,1);
}

main();