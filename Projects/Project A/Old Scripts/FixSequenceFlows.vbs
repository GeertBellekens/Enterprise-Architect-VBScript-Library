'[path=\Projects\Project A\Old Scripts]
'[group=Old Scripts]
Sub Main
	dim updateSequenceFlowsQuery
	updateSequenceFlowsQuery = "update t_xref set description = '@STEREO;Name=SequenceFlow;GUID={D48F475E-6647-4e93-9439-753FFCB06902};FQName=BPMN2.0::SequenceFlow;@ENDSTEREO;' where description like '@STEREO;Name=SequenceFlow;GUID={12BE6A97-43D3-4184-BA61-77D61267EB62};FQName=BPMN::SequenceFlow;@ENDSTEREO;'"
	Repository.Execute updateSequenceFlowsQuery
End Sub

Main