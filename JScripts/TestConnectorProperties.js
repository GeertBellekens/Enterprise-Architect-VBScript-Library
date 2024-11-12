//[group=JScripts]
function main()
{
	var conn as EA.Connector;
	
	conn = Repository.GetConnectorByGuid("{3410D31C-97AA-4b70-AB0A-469C637B67C0}")
	Session.Output("conn name " + conn.Name );
	Session.Output("   objType " + conn.ObjectType );
	Session.Output("   stereo " + conn.Stereotype );
	Session.Output("   metaType " + conn.MetaType );
	conn.Properties;
	if(conn.Properties.Count > 0) 
	{                                                // this does exist
		for(var k = 0; k < conn.Properties.Count; k++) 
		{
			Session.Output(" prop name " + conn.Properties.Item(k).Name); 
			Session.Output(" prop Value " + conn.Properties.Item(k).Value); 
			Session.Output(" prop Type " + conn.Properties.Item(k).Type	); 
			Session.Output(" prop Validation " + conn.Properties.Item(k).Validation	); 
		}
	}
}
main();

/*
function dumpConnectorInfo() {
	var elem                                                                              as EA.Element;
	var conn                                                                              as EA.Connector;
	var selectedObjects                                     as EA.Collection;
	var diagObj                                                                        as EA.DiagramObject;
	var conns                                                                            as EA.Collection;
	var diag                                                                                as EA.Diagram;
	var props                                                                             as EA.Properties;           // <--- no such EA element exists
	var prop                                                                               as EA.Property;                               // <-- no such EA element exists
   
	Repository.EnsureOutputVisible("Script");//opens the output window
	Repository.ClearOutput("Script"); // clears the output window
   
	diag  = Repository.GetCurrentDiagram();
	selectedObjects = diag.SelectedObjects();
	if(selectedObjects.Count > 0) {
		for(var i = 0; i < selectedObjects.Count; i++) {
			diagObj  = selectedObjects.GetAt(i);
			elem = Repository.getElementByID(diagObj.ElementID);
			conns = elem.Connectors;
			if(conns.Count > 0) {
				for(var j = 0; j < conns.Count; j++) {
					conn = conns.GetAt(j);
					Session.Output("conn name " + conn.Name );
					Session.Output("   objType " + conn.ObjectType );
					Session.Output("   stereo " + conn.Stereotype );
					Session.Output("   metaType " + conn.MetaType );
					conn.Properties;
					if(conn.Properties.Count > 0) {                                                // this does exist
						for(var k = 0; k < conn.Properties.Count; k++) {
										Session.Output(" prop name " + conn.Properties.Name);  // << this breaks also
						}
					}
				}
			}
		}
	}
	Session.Output("Done.");
}       
*/