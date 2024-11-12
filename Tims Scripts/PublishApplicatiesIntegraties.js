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
 Repository.GetProjectInterface().RunHTMLReport("{89CC255F-72C3-486f-90C4-20D45E5AD61E}", "H:\\_Projecten\\EA\\Sparx\\htmlexports\\applicaties\\", "png", "<default>", "html");
// Repository.GetProjectInterface().RunHTMLReport("{999E603A-E75B-4f07-A013-87D44B5BBC41}", "H:\\_Projecten\\EA\\Sparx\\htmlexports\\integraties", "PNG", "<default>", ".htm");
// Repository.GetProjectInterface().RunReport("{89CC255F-72C3-486f-90C4-20D45E5AD61E}", "Diagram Report", "H:\\_Projecten\\EA\\Sparx\\htmlexports\\applicaties\\test.pdf");
}

main();