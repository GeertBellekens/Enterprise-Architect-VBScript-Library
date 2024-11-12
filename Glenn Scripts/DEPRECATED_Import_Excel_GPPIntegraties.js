//[group=Glenn Scripts]
!INC Local Scripts.EAConstants-JScript

/*
 * Script Name: 
 * Author: Tim & Glenn
 * Purpose: updaten van bestaande integraties en hun flow-relaties naar bestaande applicaties vanuit csv
 * Date: 
 */
 
/*
 code snippets:
 
 if (guid == "") {
  integratie = integraties.Elements.AddNew(naam,"AA::Integratie");
 } else {
  integratie = Repository.GetElementByGuid("{" + guid + "}");
 }
 
  if (guid == "") {
  integraties.Elements.Refresh();
 }
 
 
 var integraties = Repository.GetPackageByGuid("{999E603A-E75B-4f07-A013-87D44B5BBC41}");
*/

function test() {
 //var integratie = Repository.GetElementByGuid("{6768F265-34DE-405f-A01F-15F6CBF35E7C}");
 
 Session.Output(determineStatus("TEST"));
 
// var integration = Repository.GetElementByGuid("{64EE8379-0165-438b-873B-16410A81B5CA}");
// var sourceApplication = Repository.GetElementByGuid("{DB4F168F-FC0B-453b-BD46-F2B6C0CFB0E9}");
// var destinationApplication = Repository.GetElementByGuid("{1820DE5E-9EDD-4d52-B7CF-C91ABFC71D72}");
 
 //var integration = Repository.GetElementByGuid("{644E41F6-CDFB-48e3-AF60-8F7CE9DD8F04}");
 //var destinationApplication = Repository.GetElementByGuid("{1820DE5E-9EDD-4d52-B7CF-C91ABFC71D72}");
 
 //var bool = controlFlowExists(integration, destinationApplication);
 
 updateIntegration(
  "{644E41F6-CDFB-48e3-AF60-8F7CE9DD8F04}",
  "IL.603",
  "UKS - Selligent (Mail Notification)",
  "UKS moet email notificaties kunnen sturen naar de klant en gebruikt hiervoor Selligent Campaigner",
  "PLANNED",
  "{C96F344F-E161-488d-9D3D-C7C8A1C6902C}",
  "{1820DE5E-9EDD-4d52-B7CF-C91ABFC71D72}"
 );
}

function determineStatus(lifecycle) {
 switch(lifecycle) {
  case "PLANNED":
   return "Validated";
  case "ACTIVE":
   return "Implemented";
  case "PHASEOUT":
  case "DECOMMISSIONED":
   return "Obsolete";
  case "UNKNOWN":
  default:
   return "Proposed";
 }
}

function updateIntegrationTaggedValues(integratie) {
// var aanspreekpunt = findTaggedValue(integratie, "aanspreekpunt");
// aanspreekpunt.Value = "Glenn Heylen";
// if (!aanspreekpunt.Update()) {
//  Session.Output("Error updating TaggedValue aanspreekpunt. Skipping..");
//  Session.Output(integratie.GetLastError());
// }
}

function findTaggedValue(element, name) {
 for (var i = 0; i < element.TaggedValues.Count; i++) {
        var taggedValue= element.TaggedValues.GetAt(i);
  if(taggedValue.Name === name) {
   return taggedValue;
  }
    }
}

function updateIntegrationProperties(integration, nummer, titel, beschrijving, status) {
 integration.Name = nummer + " " + titel;
 integration.Alias = nummer;
 integration.Notes = beschrijving;
 integration.Status = determineStatus(status);
 
 updateIntegrationTaggedValues(integration);
 
 if (!integration.Update()) {
  Session.Output("Error updating integration properties with number " + nummer + ". Skipping..");
  Session.Output(integration.GetLastError());
 }
}

function updateIntegration(guid, nummer, titel, beschrijving, status, sourceGuid, destinationGuid) { 
 var integration = Repository.GetElementByGuid(guid);
 var sourceApplication = Repository.GetElementByGuid(sourceGuid);
 var destinationApplication = Repository.GetElementByGuid(destinationGuid);
 
 if(integration == null){
  Session.Output("ERROR: Cannot find integration with guid " + guid + ", nummer " + nummer + ". Skipping..");
  return;
 }
 
 updateIntegrationProperties(integration, nummer, titel, beschrijving, status);
 updateRelations(integration, sourceApplication, destinationApplication);
 
 Session.Output("Updated: " + integration.Name);
}

function updateRelations(integration, sourceApplication, destinationApplication) {
 updateRelation(sourceApplication, integration);
 updateRelation(integration, destinationApplication);
}

function updateRelation(src, dest) {
 var connectors = findAllConnectors(src, dest);
 
 if(connectors.length == 1) {
  return;
 } else if(connectors.length < 1) {
  createControlFlow(src, dest); 
 } else {
  cleanupConnectors(connectors, src);
 }
}

function cleanupConnectors(connectors, src) {
 for (var i = 1; i < connectors.length; i++) {
  var stillRemoving = true;
  while(stillRemoving){
   for (var j = 0; j < src.Connectors.Count; j++) {
    var conn = src.Connectors.GetAt(j);
    if(conn.ConnectorGUID === connectors[i].ConnectorGUID) {
     src.Connectors.Delete(j);
     src.Connectors.Refresh();
     break;
    }
    
    if(j == src.Connectors.Count - 1) {
     stillRemoving = false;
     break;
    }
   }
  }
 }
}

function findAllConnectors(src, dest) {
 var controlFlows = [];
 for (var i = 0; i < src.Connectors.Count; i++) {
        var connector = src.Connectors.GetAt(i);
        if (connector.StereoType == "ArchiMate_Flow") {
            if (connector.SupplierID === dest.ElementID) {
                controlFlows.push(connector);
            }
        }
    }
    return controlFlows;
}

function createControlFlow(src, dest) {
 var flow = src.Connectors.AddNew("","ControlFlow");
 flow.ClientID = src.ElementID;
 flow.SupplierID = dest.ElementID;
 flow.Direction = "Source -> Destination";
 flow.StereotypeEx = "ArchiMate3::ArchiMate_Flow";
 if (!flow.Update()) {
  Session.Output("Error creating ControlFlow..");
  Session.Output(flow.GetLastError());
 }
 src.Connectors.Refresh();
}

function updateIntegrations()
{
 var fso = new ActiveXObject("Scripting.FileSystemObject");
 var ts = fso.openTextFile("H:\\Downloads\\delijn-integrations.csv", 1, true);
 var line = ts.ReadLine(); //read heading away

 while (!ts.AtEndOfStream) {
  var line = ts.ReadLine();
  
  if (/(\r?\n){2}$/.test(line)) {
   Session.Output("Skipping Empty line..");
  }
  
  Session.Output("Parsing: " + line);
  var values = line.split(";");
  var nummer = values[0];
  var titel = values[1];
  var beschrijving = values[2];
  var guid = "{" + values[3] + "}";
  var bronguid = "{" + values[4] + "}";
  var doelguid = "{" + values[5] + "}";
  var status = values[6];
  var verdict = values[7];
  
  if((!guid || guid.length === 0 )) {
   continue; //only update existing!
  }
  
  updateIntegration(guid, nummer, titel, beschrijving, status, bronguid, doelguid);
 }
}

function main() {
 //test();
 updateIntegrations();
}
main();