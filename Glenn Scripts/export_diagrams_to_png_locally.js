//[group=Glenn Scripts]
!INC Local Scripts.EAConstants-JScript

/*
 * Script Name: 
 * Author: 
 * Purpose: 
 * Date: 
 */
 
function exportAllToOneDrive()
{
 exportDiagramArrayToSubfolder(getApplicationDiagramsToExport(), "Application");
 exportDiagramArrayToSubfolder(getProjectDiagramsToExport(), "Project");
 exportDiagramArrayToSubfolder(getIntegratielijnenDiagramsToExport(), "Integratielijn");
}

function exportDiagramArrayToSubfolder(arr, subfolder) {
 for (var i = 0; i < arr.length; i++) {
  var diagramName = Repository.GetDiagramByGUID(arr[i]).Name;
  Session.Output("Exporting diagram for: " + diagramName);
  Repository.GetProjectInterface().PutDiagramImageToFile(arr[i], getOneDriveBasePath() + subfolder + "\\" + diagramName + ".png", 1);
 }
}

function getOneDriveBasePath() {
 return "C:\\Users\\35899\\OneDrive - DeLijn\\Diagrams\\";
}

function getApplicationDiagramsToExport() {
 return new Array("{05C56C76-EC21-46c0-ADEF-45CBB2987019}", "{5160FE9B-5945-42c4-8B5D-D2BC90F0358F}", "{F0BE8AEE-42DD-4f72-BB4E-9BBD26F12CCA}", "{50376EDA-0BFA-42c4-8E4A-E383D8623388}","{246DFC71-B004-41e7-82DC-3105845930D9}", "{1B2E364B-2DB8-476a-BF13-C3CAE1B5D1FD}", "{2BF8ADCB-500D-4d59-8D65-74BCCFE36438}", "{3E98F71D-073F-4caa-A8A0-37CDA2F790F0}", "{6689BFA6-5194-48dd-95E3-AA5693976D52}", "{60DE3EBB-B32D-45f7-92BB-179A1A774EF4}", "{C14F3E93-EB87-4b65-9D64-98BA6788B332}", "{463CB518-7AC9-4c50-B5A8-086B805881E3}", "{E95495A2-79BA-4f02-8219-3F334A7C6271}", "{5C366369-D419-4159-ABE8-F3FE49BB5442}");
}

function getProjectDiagramsToExport() {
 return new Array("{5C7CE95B-D74D-426e-9D27-DCC412B623C3}");
}

exportAllToOneDrive();

function getIntegratielijnenDiagramsToExport() {
 Session.Output("Getting all exportable diagrams for ILs.. This might take a while..");
 var diagramsArray = [];
 
 var potentialIls = Repository.GetPackageByGuid("{999E603A-E75B-4f07-A013-87D44B5BBC41}").Elements;
 for(var i = 0; i < potentialIls.Count; i++) {
  var potentialIl = potentialIls.GetAt(i);
  
  if("AA::Integratielijn" !== potentialIl.FQStereotype)
   continue;
  
  if("Glenn Heylen" !== potentialIl.Author && "Johan Sonck" !== potentialIl.Author)
   continue;
  
  if("Obsolete" === potentialIl.Status)
   continue;
  
  var compositeDiagram = potentialIl.CompositeDiagram;
  if(!compositeDiagram || !compositeDiagram.Name) 
   continue;
  
  diagramsArray.push(compositeDiagram.DiagramGUID);  
 }
 
 return diagramsArray;
}