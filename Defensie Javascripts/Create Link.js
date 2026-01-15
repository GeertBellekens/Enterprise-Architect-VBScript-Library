//[group=Defensie Javascripts]
!INC Local Scripts.EAConstants-JavaScript

/*
 * Script Name: This script will create a link between tables (entities) of the DATAMART 
   and corresponding views with the same name in the APP (system) folders. 
 * Author: Frederik Stouten
 * Purpose: 
 * Date: 30/06/2025
 */
 
function fastCollectionToHash(collection) {
    var hash = {};
    var count = collection.Count;

    for (var i = 0; i < count; i++) {
        var element = collection.GetAt(i);
		Session.Output("processing ["+element.Name+"] i="+i.toString());
        if (element) {
            var name = element.Name;
            if (name) {
                hash[name] = element;
            }
        }
    }

    return hash;
}
 
function convertElementsToMap(elements) {
    const elementMap = new Map();

    elements.forEach(element => {
        if (element && element.name) {
            elementMap.set(element.name, element);
        }
    });

    return elementMap;
}
 
 function findPackageByNameInRepository(targetName) {
    for (var i = 0; i < Repository.Models.Count; i++) {
        var model = Repository.Models.GetAt(i);
        var result = findPackageByName(model, targetName);
        if (result !== null) {
            return result;
        }
    }
    return null;
}

function findPackageByName(pkg, targetName) {
    if (pkg.Name === targetName) {
        return pkg;
    }

    for (var i = 0; i < pkg.Packages.Count; i++) {
        var child = pkg.Packages.GetAt(i);
        var result = findPackageByName(child, targetName);
        if (result !== null) {
            return result;
        }
    }

    return null;
}
 
/*	Gets all elements as an EA.Collection in a given sub-package "subpackageName", in a given package
	"packageName.
*/
function getElementsFromNamedSubpackage(packageName, subpackageName, stereoType) {
	
    var parentPackage = findPackageByNameInRepository(packageName);
    if (!parentPackage) {
        Session.Output("Parent package '" + packageName + "' not found.");
        return [];
    }
	
    var subpackage = findPackageByName(parentPackage, subpackageName);
    if (!subpackage) {
        Session.Output("Subpackage '" + subpackageName + "' not found under '" + packageName + "'.");
        return [];
    }
	
	if (stereoType === "table") {
		Session.Output("start converting ...");
		return fastCollectionToHash(subpackage.Elements);
		Session.Output("end converting.");
	}
	if (stereoType === "view") {
		return subpackage.Elements;
	}
}

/* Links one table element to another view element. 
*/
function linkTableToView(tableElement, viewElement) {

    if (!tableElement || !viewElement) {
        Session.Output("Error: One or both elements not found.");
        return;
    }

    // Create a new connector from the table to the view
    var connector = tableElement.Connectors.AddNew("Table to View Link", "Dependency");
    connector.SupplierID = viewElement.ElementID;

    if (!connector.Update()) {
        Session.Output("Error: Failed to create connector.");
        return;
    }

    tableElement.Connectors.Refresh();
    Session.Output("Connector created between '" + tableElement.Name + "' and '" + viewElement.Name + "'.");
}

function replaceStart(original, targetStart, replacement) {
  if (original.startsWith(targetStart)) {
    return replacement + original.slice(targetStart.length);
  }
  return original;
}

/* Links all tables in "tablesMap" to all corresponding views in "views" based on same name.
*/
function linkTablesToViews(tablesMap, views) {
	for (var i = 0; i < views.Count; i++) {
		var view = views.GetAt(i);
		var viewName = view.Name.toLowerCase();
		Session.Output("processing view ["+viewName+"].");
		// for ALICE replace prefix v_datamart by alice!
		var viewNameAdapted = replaceStart(viewName, "v_datamart", "netmanto");
		// for datamart manager add "datamart_ to string.
		//viewNameAdapted = viewName;
		var table = tablesMap[viewNameAdapted];
		linkTableToView(table, view);
	}
}
 
function main()
{
	var dataMartPackageName = "DATAMART";
	var dataMartViewsPackageName = "datamart_views";
	//var applicationPackageName = "APP_Service_Catalog";
	//var applicationPackageName = "APP_ALICEV2";
	//var applicationPackageName = "APP_Datamart_manager";
	var applicationPackageName = "APP_NETMANTO";
	var applicationViewsPackageName = "Views";
	
	//Get the list of tables inside the Views package in the DATAMART package.
	var tablesMap = getElementsFromNamedSubpackage(dataMartPackageName, dataMartViewsPackageName, "table")
	
	Session.Output("Found " + Object.keys(tablesMap).length + " tables in datamart.");
	
	//Get the list of views inside the Views package in the application package.
	var views = getElementsFromNamedSubpackage(applicationPackageName, applicationViewsPackageName, "view");
	
	Session.Output("Found " + views.Count + " views in application.");
	
	linkTablesToViews(tablesMap, views);
}

main();