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
 if (Session.Prompt("Are you sure? This will block your session for a while...", 4) == 1) {
  Repository.GetProjectInterface().ExportPackageXMI("{A860D908-E8EA-4baa-BE4B-3E6A4D0E1284}",0,0,0,0,0,"H:\\autoexport.xmi");
 }
}

OnProjectBrowserScript();
