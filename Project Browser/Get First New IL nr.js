//[group=Project Browser]
!INC Local Scripts.EAConstants-JScript

/*
 * This code has been included from the default Project Browser template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.   
 * 
 * Script Name:
 * Author:
 * Purpose:
 * Date:
 */

/*
 * Project Browser Script main function
 */
function OnProjectBrowserScript()
{
 integraties = Repository.GetPackageByGuid("{999E603A-E75B-4f07-A013-87D44B5BBC41}");
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

OnProjectBrowserScript();
