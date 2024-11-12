//[group=Diagram Group]
!INC Local Scripts.EAConstants-JavaScript

/*
 * This code has been included from the default Diagram Script template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.
 *
 * Script Name:
 * Author:
 * Purpose:
 * Date:
 */

/*
 * Diagram Script main function
 */
function OnDiagramScript()
{
	// Get a reference to the current diagram
	var currentDiagram as EA.Diagram;
	currentDiagram = Repository.GetCurrentDiagram();
	currentDiagram = Repository.GetDiagramByGuid("{5814202A-5F51-4ff1-BE8D-3AC2B55949BF}");

	if ( currentDiagram != null )
	{
		// Get a reference to any selected connector/objects
		var selectedConnector as EA.Connector;
		var selectedObjects as EA.Collection;
		selectedConnector = currentDiagram.SelectedConnector;
		selectedObjects = currentDiagram.SelectedObjects;

		if ( selectedConnector != null )
		{
			// A connector is selected
		}
		else if ( selectedObjects.Count > 0 )
		{
			// One or more diagram objects are selected
			var objectCount = 0;
			
			while (objectCount < selectedObjects.Count)
			//while (objectCount < selectedObjects.Count())
				{
					Session.Output("Enumerating");
					var currentDiagElement as EA.DiagramObject;
					var currentElement as EA.Element;
					currentDiagElement = selectedObjects.GetAt(objectCount);
					currentElement = Repository.GetElementByID(currentDiagElement.ElementID);
					//Session.Prompt(currentElement.Type, promptOK); //[1]
					//Session.Prompt(currentElement.Name, promptOK); //[2]
					//Session.Prompt(currentElement.Stereotype, promptOK); //[3]						
					currentElement.Type = "Class";
					currentElement.blabla = "";
					Session.Prompt( "blabla", promptOK);
					currentElement.Update();
					//currentElement.Stereotype = "ArchiMate_Requirement"; //[5]
					currentElement.Stereotype = "ArchiMate3::ArchiMate_Requirement"; //[6]
					currentElement.Update();
					//Session.Prompt(currentElement.Type, promptOK); //[7]
					//Session.Prompt(currentElement.Name, promptOK); //[8]
					//Session.Prompt(currentElement.Stereotype, promptOK); //[9]					
					objectCount = objectCount + 1;
				}
		}
		else
		{
			// Nothing is selected
		}
	}
	else
	{
		Session.Prompt( "This script requires a diagram to be visible.", promptOK)
	}
	Repository.ReloadDiagram(currentDiagram.DiagramID);
}

OnDiagramScript();
