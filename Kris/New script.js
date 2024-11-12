//[group=Kris]
!INC Local Scripts.EAConstants-JScript

/*
 * Script Name: 
 * Author: 
 * Purpose: 
 * Date: 
 */
 
function main()
{
 var elementID;
 var newStereotypeName;
 
 elementID = "{1A332897-86EE-4201-BF5A-DA527945EC82}";
 
 item = Repository.GetElementByGuid(elementID);
 Session.Output("Name:             " + item.Name);
 Session.Output("Abstract:         " + item.Abstract);
 Session.Output("ActionFlags:      " + item.ActionFlags);
 Session.Output("Alias:            " + item.Alias);
 Session.Output("AssociationClassConnectorID: " + item.AssociationClassConnectorID);
 Session.Output("Attributes:                  " + item.Attributes);
 Session.Output("AttributesEx:                " + item.AttributesEx);
 Session.Output("Author:       " + item.Author);
 Session.Output("BaseClasses:                 " + item.BaseClasses);
 Session.Output(item.ClassfierID);
 Session.Output(item.ClassifierID);
 Session.Output(item.ClassifierName);
 Session.Output(item.ClassifierType);
 Session.Output(item.Complexity);
 Session.Output(item.CompositeDiagram);
 Session.Output(item.Connectors);
 Session.Output(item.Constraints);
 Session.Output(item.ConstraintsEx);
 Session.Output(item.Created);
 Session.Output(item.CustomProperties);
 Session.Output(item.Diagrams);
 Session.Output(item.Difficulty);
 Session.Output(item.Efforts);
 Session.Output(item.ElementGUID);
 Session.Output(item.ElementID);
 Session.Output(item.Elements);
 Session.Output(item.EmbeddedElements);
 Session.Output(item.EventFlags);
 Session.Output(item.ExtensionPoints);
 Session.Output(item.Files);
 Session.Output(item.FQName);
 Session.Output(item.FQStereotype);
 Session.Output(item.GenFile);
 Session.Output(item.Genlinks);
 Session.Output(item.GenType);
 Session.Output(item.Header1);
 Session.Output(item.Header2);
 Session.Output(item.IsActive);
 Session.Output(item.IsComposite);
 Session.Output(item.IsLeaf);
 Session.Output(item.IsNew);
 Session.Output(item.IsRoot);
 Session.Output(item.IsSpec);
 Session.Output(item.Issues);
 Session.Output(item.Locked);
 Session.Output(item.MetaType);
 Session.Output(item.Methods);
 Session.Output(item.MethodsEx);
 Session.Output(item.Metrics);
 Session.Output(item.MiscData);
 Session.Output(item.Modified);
 Session.Output(item.Multiplicity);
 Session.Output(item.Name);
 Session.Output(item.Notes);
 Session.Output(item.ObjectType);
 Session.Output(item.PackageID);
 Session.Output(item.ParentID);
 Session.Output(item.Partitions);
 Session.Output(item.Persistence);
 Session.Output(item.Phase);
 Session.Output(item.Priority);
 Session.Output(item.Properties);
 Session.Output(item.PropertyType);
 Session.Output(item.PropertyTypeName);
 Session.Output(item.Realizes);
 Session.Output(item.Requirements);
 Session.Output(item.RequirementsEx);
 Session.Output(item.Resources);
 Session.Output(item.Risks);
 Session.Output(item.RunState);
 Session.Output(item.Scenarios);
 Session.Output(item.StateTransitions);
 Session.Output(item.Status);
 Session.Output(item.Stereotype);
 Session.Output(item.StereotypeEx);
 Session.Output(item.StyleEx);
 Session.Output(item.Subtype);
 Session.Output(item.Tablespace);
 Session.Output(item.Tag);
 Session.Output(item.TaggedValues);
 Session.Output(item.TaggedValuesEx);
 Session.Output(item.TemplateParameters);
 Session.Output(item.Tests);
 Session.Output(item.TreePos);
 Session.Output(item.Type);
 Session.Output(item.TypeInfoProperties);
 Session.Output(item.Version);
 Session.Output(item.Visibility);
 }

main();