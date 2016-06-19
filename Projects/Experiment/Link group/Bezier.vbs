'[path=\Projects\Experiment\Link group]
'[group=Link group]

sub main
	dim selectedConnector as EA.Connector
	set selectedConnector = Repository.GetContextObject
	dim taggedValue as EA.ConnectorTag
	set taggedValue = selectedConnector.TaggedValues.AddNew("_Bezier","1")
	taggedValue.Update
end sub

main