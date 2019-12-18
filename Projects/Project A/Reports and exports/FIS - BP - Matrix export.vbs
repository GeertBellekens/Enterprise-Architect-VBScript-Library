'[path=\Projects\Project A\Reports and exports]
'[group=Reports and exports]

option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include

'
' Script Name: Create FIS - BP - Matrix export
' Author: Geert Bellekens
' Purpose: Create an excel export with the data of the query FIS-BP-Matrix
' Date: 2019-10-16
'

const excelTemplate = "G:\Projects\80 Enterprise Architect\Output\Excel export templates\UMIG DGO - IM - XD - 05 - FIS-Business Process Matrix template.xltx"

sub main
	'get the results
	dim sqlGetResults
	sqlGetResults = getSQLQuery()
	dim queryResult
	set queryResult = getArrayListFromQuery(sqlGetResults)
	'get the headers
	dim headers
	set headers = getHeaders()
	'add the headers to the results
	queryResult.Insert 0, headers
	'create the excel file
	dim excelOutput
	set excelOutput = new ExcelFile
	'load the template
	excelOutput.NewFile excelTemplate
	'create a two dimensional array
	dim excelContents
	excelContents = makeArrayFromArrayLists(queryResult)
	'add the output to a sheet in excel
	dim sheet
	set sheet = excelOutput.createTab("FIS-BP Matrix", excelContents, true, "TableStyleMedium4")
	'set headers to atrias red
	dim headerRange
	set headerRange = sheet.Range(sheet.Cells(1,1), sheet.Cells(1, headers.Count))
	excelOutput.formatRange headerRange, atriasRed, "default", "default", "default", "default", "default"
	'save the excel file
	excelOutput.save
end sub


function getSQLQuery()
	getSQLQuery = 	"select stuff( COALESCE(', ' + mx.NA, '')                                                                                                                                " & vbNewLine & _
					" 			+ COALESCE(', ' + mx.UMIG, '')                                                                                                                               " & vbNewLine & _
					" 			+ COALESCE(', ' + mx.[UMIG AO], '')                                                                                                                          " & vbNewLine & _
					" 			+ COALESCE(', ' + mx.[UMIG DGO], '')                                                                                                                         " & vbNewLine & _
					" 			+ COALESCE(', ' + mx.[UMIG PaP], '')                                                                                                                         " & vbNewLine & _
					" 			+ COALESCE(', ' + mx.[UMIG PPP], '')                                                                                                                         " & vbNewLine & _
					" 			+ COALESCE(', ' + mx.[UMIG TPDA], '')                                                                                                                        " & vbNewLine & _
					" 			+ COALESCE(', ' + mx.[UMIG TSO], '')                                                                                                                         " & vbNewLine & _
					" 			,1,2,'') as UMIG                                                                                                                                             " & vbNewLine & _
					" , mx.Message, mx.FIS, mx.FISDirection                                                                                                                                  " & vbNewLine & _
					" , mx.BusinessProcess, mx.SubProcess_1, mx.SubProcess_2, mx.SubProcess_3, mx.SubProcess_4, mx.SubProcess_5                                                              " & vbNewLine & _
					" , mx.SourcePool, mx.SourceLane, mx.TargetPool, mx.TargetLane                                                                                                           " & vbNewLine & _
					" from                                                                                                                                                                   " & vbNewLine & _
					" (                                                                                                                                                                      " & vbNewLine & _
					" select                                                                                                                                                                 " & vbNewLine & _
					" case when tvU.Value like '1%' then 'N/A' end as NA,                                                                                                                    " & vbNewLine & _
					" case when tvU.Value like '_,1%' then 'UMIG' end as UMIG,                                                                                                               " & vbNewLine & _
					" case when tvU.Value like '_,_,1%' then 'UMIG AO' end as [UMIG AO],                                                                                                     " & vbNewLine & _
					" case when tvU.Value like '_,_,_,1%' then 'UMIG DGO' end as [UMIG DGO],                                                                                                 " & vbNewLine & _
					" case when tvU.Value like '_,_,_,_,1%' then 'UMIG PaP' end as [UMIG PaP],                                                                                               " & vbNewLine & _
					" case when tvU.Value like '_,_,_,_,_,1%' then 'UMIG PPP' end as [UMIG PPP],                                                                                             " & vbNewLine & _
					" case when tvU.Value like '_,_,_,_,_,_,1%' then 'UMIG TPDA' end as [UMIG TPDA],                                                                                         " & vbNewLine & _
					" case when tvU.Value like '_,_,_,_,_,_,_,1%' then 'UMIG TSO' end as [UMIG TSO],                                                                                         " & vbNewLine & _
					" mtr.*                                                                                                                                                                  " & vbNewLine & _
					" from                                                                                                                                                                   " & vbNewLine & _
					" (                                                                                                                                                                      " & vbNewLine & _
					" SELECT                                                                                                                                                                 " & vbNewLine & _
					" p.ea_guid AS CLASSGUID, p.Object_Type AS CLASSTYPE,                                                                                                                    " & vbNewLine & _
					" m.name AS FIS, m.Object_ID as FisID,                                                                                                                                   " & vbNewLine & _
					" (CASE WHEN mtv.PropertyID is not null and mtv.Value IS NULL THEN 'In' ELSE mtv.Value END) as FISDirection,                                                             " & vbNewLine & _
					" t.name as Message,                                                                                                                                                     " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name  is null THEN (CASE WHEN sourcePCL.Name is null THEN sourceCL.Name ELSE sourcePCL.Name END) ELSE sourcePoolCL.Name END) as SourcePool,    " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name is not null THEN sourcePCL.Name  ELSE null END) as SourceLane,                                                                            " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name  is null THEN (CASE WHEN targetPCL.Name is null THEN targetCL.Name ELSE targetPCL.Name END)ELSE targetPoolCL.Name END) as TargetPool,     " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name is not null THEN targetPCL.Name  ELSE null END)  as TargetLane,                                                                           " & vbNewLine & _
					" p.Name as BusinessProcess, null as SubProcess_1,null as SubProcess_2,null as SubProcess_3,null as SubProcess_4,null as SubProcess_5                                    " & vbNewLine & _
					" FROM t_connector c                                                                                                                                                     " & vbNewLine & _
					" inner join t_connectortag ct on ct.ElementID = c.Connector_ID                                                                                                          " & vbNewLine & _
					"                                                and ct.Property = 'MessageRef'                                                                                          " & vbNewLine & _
					" inner join t_object m on m.ea_guid = ct.VALUE                                                                                                                          " & vbNewLine & _
					" left join t_objectproperties mtv on mtv.Object_ID = m.Object_ID                                                                                                        " & vbNewLine & _
					"                                                              and mtv.Property = 'Atrias::Direction'                                                                    " & vbNewLine & _
					" inner join t_diagramlinks dl on dl.ConnectorID = c.Connector_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d on d.Diagram_ID = dl.DiagramID                                                                                                                  " & vbNewLine & _
					" inner join t_object p on p.Object_ID = d.ParentID                                                                                                                      " & vbNewLine & _
					" inner join t_object source on source.Object_ID = c.Start_Object_ID                                                                                                     " & vbNewLine & _
					" left join t_object  sourceCl on sourceCL.Object_ID = source.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object sourceP on sourceP.Object_ID = source.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object sourcePCL on sourcePCL.Object_ID = sourceP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object sourcePool on sourcePool.Object_ID = sourceP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and sourcePool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object sourcePoolCL on sourcePoolCL.Object_ID = sourcePool.Classifier                                                                                      " & vbNewLine & _
					" inner join t_object target on target.Object_ID = c.End_Object_ID                                                                                                       " & vbNewLine & _
					" left join t_object  targetCl on targetCL.Object_ID = target.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object targetP on targetP.Object_ID = target.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object targetPCL on targetPCL.Object_ID = targetP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object targetPool on targetPool.Object_ID = targetP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and targetPool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object targetPoolCL on targetPoolCL.Object_ID = targetPool.Classifier                                                                                      " & vbNewLine & _
					" left join t_connector mtt on m.object_ID = mtt.End_object_ID                                                                                                           " & vbNewLine & _
					"                                         and mtt.Connector_Type in( 'Realization', 'Realisation')                                                                       " & vbNewLine & _
					" left join t_object t on mtt.Start_Object_ID = t.Object_ID                                                                                                              " & vbNewLine & _
					"                                         and t.stereotype = 'message'                                                                                                   " & vbNewLine & _
					" union                                                                                                                                                                  " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" SELECT --level1                                                                                                                                                        " & vbNewLine & _
					" bp1.ea_guid AS CLASSGUID, bp1.Object_Type AS CLASSTYPE,                                                                                                                " & vbNewLine & _
					" m.name AS FIS, m.Object_ID as FisID,                                                                                                                                   " & vbNewLine & _
					" (CASE WHEN mtv.PropertyID is not null and mtv.Value IS NULL THEN 'In' ELSE mtv.Value END) as FISDirection,                                                             " & vbNewLine & _
					" t.name as Message,                                                                                                                                                     " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name  is null THEN (CASE WHEN sourcePCL.Name is null THEN sourceCL.Name ELSE sourcePCL.Name END) ELSE sourcePoolCL.Name END) as SourcePool,    " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name is not null THEN sourcePCL.Name  ELSE null END) as SourceLane,                                                                            " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name  is null THEN (CASE WHEN targetPCL.Name is null THEN targetCL.Name ELSE targetPCL.Name END)ELSE targetPoolCL.Name END) as TargetPool,     " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name is not null THEN targetPCL.Name  ELSE null END)  as TargetLane,                                                                           " & vbNewLine & _
					" bp1.Name as BusinessProcess, p.name as SubProcess_1,null as SubProcess_2,null as SubProcess_3,null as SubProcess_4,null as SubProcess_5                                " & vbNewLine & _
					" FROM t_connector c                                                                                                                                                     " & vbNewLine & _
					" inner join t_connectortag ct on ct.ElementID = c.Connector_ID                                                                                                          " & vbNewLine & _
					"                                                and ct.Property = 'MessageRef'                                                                                          " & vbNewLine & _
					" inner join t_object m on m.ea_guid = ct.VALUE                                                                                                                          " & vbNewLine & _
					" left join t_objectproperties mtv on mtv.Object_ID = m.Object_ID                                                                                                        " & vbNewLine & _
					"                                                              and mtv.Property = 'Atrias::Direction'                                                                    " & vbNewLine & _
					" inner join t_diagramlinks dl on dl.ConnectorID = c.Connector_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d on d.Diagram_ID = dl.DiagramID                                                                                                                  " & vbNewLine & _
					" inner join t_object p on p.Object_ID = d.ParentID                                                                                                                      " & vbNewLine & _
					"                                   and p.Stereotype in ('Activity','BusinessProcess','ArchiMate_BusinessProcess')                                                       " & vbNewLine & _
					" inner join t_objectProperties tvca1 on tvca1.Value = p.ea_guid                                                                                                         " & vbNewLine & _
					"                                                                    and tvca1.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p1 on p1.Object_ID = tvca1.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p1.Stereotype in ('Activity','BusinessProcess','ArchiMate_BusinessProcess')                                                " & vbNewLine & _
					" inner join t_diagramObjects do1 on do1.Object_ID = p1.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d1 on d1.Diagram_ID = do1.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp1 on bp1.Object_ID = d1.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_object source on source.Object_ID = c.Start_Object_ID                                                                                                     " & vbNewLine & _
					" left join t_object  sourceCl on sourceCL.Object_ID = source.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object sourceP on sourceP.Object_ID = source.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object sourcePCL on sourcePCL.Object_ID = sourceP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object sourcePool on sourcePool.Object_ID = sourceP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and sourcePool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object sourcePoolCL on sourcePoolCL.Object_ID = sourcePool.Classifier                                                                                      " & vbNewLine & _
					" inner join t_object target on target.Object_ID = c.End_Object_ID                                                                                                       " & vbNewLine & _
					" left join t_object  targetCl on targetCL.Object_ID = target.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object targetP on targetP.Object_ID = target.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object targetPCL on targetPCL.Object_ID = targetP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object targetPool on targetPool.Object_ID = targetP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and targetPool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object targetPoolCL on targetPoolCL.Object_ID = targetPool.Classifier                                                                                      " & vbNewLine & _
					" left join t_connector mtt on m.object_ID = mtt.End_object_ID                                                                                                           " & vbNewLine & _
					"                                         and mtt.Connector_Type in( 'Realization', 'Realisation')                                                                       " & vbNewLine & _
					" left join t_object t on mtt.Start_Object_ID = t.Object_ID                                                                                                              " & vbNewLine & _
					"                                         and t.stereotype = 'message'                                                                                                   " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" union                                                                                                                                                                  " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" SELECT --level2                                                                                                                                                        " & vbNewLine & _
					" bp2.ea_guid AS CLASSGUID, bp2.Object_Type AS CLASSTYPE,                                                                                                                " & vbNewLine & _
					" m.name AS FIS, m.Object_ID as FisID,                                                                                                                                   " & vbNewLine & _
					" (CASE WHEN mtv.PropertyID is not null and mtv.Value IS NULL THEN 'In' ELSE mtv.Value END) as FISDirection,                                                             " & vbNewLine & _
					" t.name as Message,                                                                                                                                                     " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name  is null THEN (CASE WHEN sourcePCL.Name is null THEN sourceCL.Name ELSE sourcePCL.Name END) ELSE sourcePoolCL.Name END) as SourcePool,    " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name is not null THEN sourcePCL.Name  ELSE null END) as SourceLane,                                                                            " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name  is null THEN (CASE WHEN targetPCL.Name is null THEN targetCL.Name ELSE targetPCL.Name END)ELSE targetPoolCL.Name END) as TargetPool,     " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name is not null THEN targetPCL.Name  ELSE null END)  as TargetLane,                                                                           " & vbNewLine & _
					" bp2.Name as BusinessProcess, bp1.name as SubProcess_1, p.Name as SubProcess_2,null as SubProcess_3,null as SubProcess_4,null as SubProcess_5                           " & vbNewLine & _
					" FROM t_connector c                                                                                                                                                     " & vbNewLine & _
					" inner join t_connectortag ct on ct.ElementID = c.Connector_ID                                                                                                          " & vbNewLine & _
					"                                                and ct.Property = 'MessageRef'                                                                                          " & vbNewLine & _
					" inner join t_object m on m.ea_guid = ct.VALUE                                                                                                                          " & vbNewLine & _
					" left join t_objectproperties mtv on mtv.Object_ID = m.Object_ID                                                                                                        " & vbNewLine & _
					"                                                              and mtv.Property = 'Atrias::Direction'                                                                    " & vbNewLine & _
					" inner join t_diagramlinks dl on dl.ConnectorID = c.Connector_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d on d.Diagram_ID = dl.DiagramID                                                                                                                  " & vbNewLine & _
					" inner join t_object p on p.Object_ID = d.ParentID                                                                                                                      " & vbNewLine & _
					"                                   and p.Stereotype in ('Activity','BusinessProcess','ArchiMate_BusinessProcess')                                                       " & vbNewLine & _
					" inner join t_objectProperties tvca1 on tvca1.Value = p.ea_guid                                                                                                         " & vbNewLine & _
					"                                                                    and tvca1.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p1 on p1.Object_ID = tvca1.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p1.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do1 on do1.Object_ID = p1.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d1 on d1.Diagram_ID = do1.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp1 on bp1.Object_ID = d1.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca2 on tvca2.Value = bp1.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca2.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p2 on p2.Object_ID = tvca2.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p2.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do2 on do2.Object_ID = p2.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d2 on d2.Diagram_ID = do2.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp2 on bp2.Object_ID = d2.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_object source on source.Object_ID = c.Start_Object_ID                                                                                                     " & vbNewLine & _
					" left join t_object  sourceCl on sourceCL.Object_ID = source.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object sourceP on sourceP.Object_ID = source.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object sourcePCL on sourcePCL.Object_ID = sourceP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object sourcePool on sourcePool.Object_ID = sourceP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and sourcePool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object sourcePoolCL on sourcePoolCL.Object_ID = sourcePool.Classifier                                                                                      " & vbNewLine & _
					" inner join t_object target on target.Object_ID = c.End_Object_ID                                                                                                       " & vbNewLine & _
					" left join t_object  targetCl on targetCL.Object_ID = target.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object targetP on targetP.Object_ID = target.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object targetPCL on targetPCL.Object_ID = targetP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object targetPool on targetPool.Object_ID = targetP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and targetPool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object targetPoolCL on targetPoolCL.Object_ID = targetPool.Classifier                                                                                      " & vbNewLine & _
					" left join t_connector mtt on m.object_ID = mtt.End_object_ID                                                                                                           " & vbNewLine & _
					"                                         and mtt.Connector_Type in( 'Realization', 'Realisation')                                                                       " & vbNewLine & _
					" left join t_object t on mtt.Start_Object_ID = t.Object_ID                                                                                                              " & vbNewLine & _
					"                                         and t.stereotype = 'message'                                                                                                   " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" union                                                                                                                                                                  " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" SELECT --level3                                                                                                                                                        " & vbNewLine & _
					" bp3.ea_guid AS CLASSGUID, bp3.Object_Type AS CLASSTYPE,                                                                                                                " & vbNewLine & _
					" m.name AS FIS, m.Object_ID as FisID,                                                                                                                                   " & vbNewLine & _
					" (CASE WHEN mtv.PropertyID is not null and mtv.Value IS NULL THEN 'In' ELSE mtv.Value END) as FISDirection,                                                             " & vbNewLine & _
					" t.name as Message,                                                                                                                                                     " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name  is null THEN (CASE WHEN sourcePCL.Name is null THEN sourceCL.Name ELSE sourcePCL.Name END) ELSE sourcePoolCL.Name END) as SourcePool,    " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name is not null THEN sourcePCL.Name  ELSE null END) as SourceLane,                                                                            " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name  is null THEN (CASE WHEN targetPCL.Name is null THEN targetCL.Name ELSE targetPCL.Name END)ELSE targetPoolCL.Name END) as TargetPool,     " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name is not null THEN targetPCL.Name  ELSE null END)  as TargetLane,                                                                           " & vbNewLine & _
					" bp3.Name as BusinessProcess, bp2.name as SubProcess_1, bp1.Name as SubProcess_2, p.name as SubProcess_3, null as SubProcess_4, null as SubProcess_5                    " & vbNewLine & _
					" FROM t_connector c                                                                                                                                                     " & vbNewLine & _
					" inner join t_connectortag ct on ct.ElementID = c.Connector_ID                                                                                                          " & vbNewLine & _
					"                                                and ct.Property = 'MessageRef'                                                                                          " & vbNewLine & _
					" inner join t_object m on m.ea_guid = ct.VALUE                                                                                                                          " & vbNewLine & _
					" left join t_objectproperties mtv on mtv.Object_ID = m.Object_ID                                                                                                        " & vbNewLine & _
					"                                                              and mtv.Property = 'Atrias::Direction'                                                                    " & vbNewLine & _
					" inner join t_diagramlinks dl on dl.ConnectorID = c.Connector_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d on d.Diagram_ID = dl.DiagramID                                                                                                                  " & vbNewLine & _
					" inner join t_object p on p.Object_ID = d.ParentID                                                                                                                      " & vbNewLine & _
					"                                   and p.Stereotype in ('Activity','BusinessProcess','ArchiMate_BusinessProcess')                                                       " & vbNewLine & _
					" inner join t_objectProperties tvca1 on tvca1.Value = p.ea_guid                                                                                                         " & vbNewLine & _
					"                                                                    and tvca1.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p1 on p1.Object_ID = tvca1.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p1.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do1 on do1.Object_ID = p1.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d1 on d1.Diagram_ID = do1.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp1 on bp1.Object_ID = d1.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca2 on tvca2.Value = bp1.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca2.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p2 on p2.Object_ID = tvca2.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p2.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do2 on do2.Object_ID = p2.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d2 on d2.Diagram_ID = do2.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp2 on bp2.Object_ID = d2.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca3 on tvca3.Value = bp2.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca3.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p3 on p3.Object_ID = tvca3.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p3.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do3 on do3.Object_ID = p3.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d3 on d3.Diagram_ID = do3.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp3 on bp3.Object_ID = d3.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_object source on source.Object_ID = c.Start_Object_ID                                                                                                     " & vbNewLine & _
					" left join t_object  sourceCl on sourceCL.Object_ID = source.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object sourceP on sourceP.Object_ID = source.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object sourcePCL on sourcePCL.Object_ID = sourceP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object sourcePool on sourcePool.Object_ID = sourceP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and sourcePool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object sourcePoolCL on sourcePoolCL.Object_ID = sourcePool.Classifier                                                                                      " & vbNewLine & _
					" inner join t_object target on target.Object_ID = c.End_Object_ID                                                                                                       " & vbNewLine & _
					" left join t_object  targetCl on targetCL.Object_ID = target.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object targetP on targetP.Object_ID = target.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object targetPCL on targetPCL.Object_ID = targetP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object targetPool on targetPool.Object_ID = targetP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and targetPool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object targetPoolCL on targetPoolCL.Object_ID = targetPool.Classifier                                                                                      " & vbNewLine & _
					" left join t_connector mtt on m.object_ID = mtt.End_object_ID                                                                                                           " & vbNewLine & _
					"                                         and mtt.Connector_Type in( 'Realization', 'Realisation')                                                                       " & vbNewLine & _
					" left join t_object t on mtt.Start_Object_ID = t.Object_ID                                                                                                              " & vbNewLine & _
					"                                         and t.stereotype = 'message'                                                                                                   " & vbNewLine & _
					" union                                                                                                                                                                  " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" SELECT --level4                                                                                                                                                        " & vbNewLine & _
					" bp4.ea_guid AS CLASSGUID, bp4.Object_Type AS CLASSTYPE,                                                                                                                " & vbNewLine & _
					" m.name AS FIS, m.Object_ID as FisID,                                                                                                                                   " & vbNewLine & _
					" (CASE WHEN mtv.PropertyID is not null and mtv.Value IS NULL THEN 'In' ELSE mtv.Value END) as FISDirection,                                                             " & vbNewLine & _
					" t.name as Message,                                                                                                                                                     " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name  is null THEN (CASE WHEN sourcePCL.Name is null THEN sourceCL.Name ELSE sourcePCL.Name END) ELSE sourcePoolCL.Name END) as SourcePool,    " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name is not null THEN sourcePCL.Name  ELSE null END) as SourceLane,                                                                            " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name  is null THEN (CASE WHEN targetPCL.Name is null THEN targetCL.Name ELSE targetPCL.Name END)ELSE targetPoolCL.Name END) as TargetPool,     " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name is not null THEN targetPCL.Name  ELSE null END)  as TargetLane,                                                                           " & vbNewLine & _
					" bp4.Name as BusinessProcess, bp3.name as SubProcess_1, bp2.Name as SubProcess_2, bp1.name  as SubProcess_3, p.name as SubProcess_4,null as SubProcess_5                " & vbNewLine & _
					" FROM t_connector c                                                                                                                                                     " & vbNewLine & _
					" inner join t_connectortag ct on ct.ElementID = c.Connector_ID                                                                                                          " & vbNewLine & _
					"                                                and ct.Property = 'MessageRef'                                                                                          " & vbNewLine & _
					" inner join t_object m on m.ea_guid = ct.VALUE                                                                                                                          " & vbNewLine & _
					" left join t_objectproperties mtv on mtv.Object_ID = m.Object_ID                                                                                                        " & vbNewLine & _
					"                                                              and mtv.Property = 'Atrias::Direction'                                                                    " & vbNewLine & _
					" inner join t_diagramlinks dl on dl.ConnectorID = c.Connector_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d on d.Diagram_ID = dl.DiagramID                                                                                                                  " & vbNewLine & _
					" inner join t_object p on p.Object_ID = d.ParentID                                                                                                                      " & vbNewLine & _
					"                                   and p.Stereotype in ('Activity','BusinessProcess','ArchiMate_BusinessProcess')                                                       " & vbNewLine & _
					" inner join t_objectProperties tvca1 on tvca1.Value = p.ea_guid                                                                                                         " & vbNewLine & _
					"                                                                    and tvca1.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p1 on p1.Object_ID = tvca1.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p1.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do1 on do1.Object_ID = p1.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d1 on d1.Diagram_ID = do1.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp1 on bp1.Object_ID = d1.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca2 on tvca2.Value = bp1.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca2.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p2 on p2.Object_ID = tvca2.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p2.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do2 on do2.Object_ID = p2.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d2 on d2.Diagram_ID = do2.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp2 on bp2.Object_ID = d2.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca3 on tvca3.Value = bp2.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca3.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p3 on p3.Object_ID = tvca3.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p3.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do3 on do3.Object_ID = p3.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d3 on d3.Diagram_ID = do3.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp3 on bp3.Object_ID = d3.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca4 on tvca4.Value = bp3.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca4.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p4 on p4.Object_ID = tvca4.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p4.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do4 on do4.Object_ID = p4.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d4 on d4.Diagram_ID = do4.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp4 on bp4.Object_ID = d4.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_object source on source.Object_ID = c.Start_Object_ID                                                                                                     " & vbNewLine & _
					" left join t_object  sourceCl on sourceCL.Object_ID = source.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object sourceP on sourceP.Object_ID = source.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object sourcePCL on sourcePCL.Object_ID = sourceP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object sourcePool on sourcePool.Object_ID = sourceP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and sourcePool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object sourcePoolCL on sourcePoolCL.Object_ID = sourcePool.Classifier                                                                                      " & vbNewLine & _
					" inner join t_object target on target.Object_ID = c.End_Object_ID                                                                                                       " & vbNewLine & _
					" left join t_object  targetCl on targetCL.Object_ID = target.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object targetP on targetP.Object_ID = target.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object targetPCL on targetPCL.Object_ID = targetP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object targetPool on targetPool.Object_ID = targetP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and targetPool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object targetPoolCL on targetPoolCL.Object_ID = targetPool.Classifier                                                                                      " & vbNewLine & _
					" left join t_connector mtt on m.object_ID = mtt.End_object_ID                                                                                                           " & vbNewLine & _
					"                                         and mtt.Connector_Type in( 'Realization', 'Realisation')                                                                       " & vbNewLine & _
					" left join t_object t on mtt.Start_Object_ID = t.Object_ID                                                                                                              " & vbNewLine & _
					"                                         and t.stereotype = 'message'                                                                                                   " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" union                                                                                                                                                                  " & vbNewLine & _
					"                                                                                                                                                                        " & vbNewLine & _
					" SELECT --level5                                                                                                                                                        " & vbNewLine & _
					" bp5.ea_guid AS CLASSGUID, bp5.Object_Type AS CLASSTYPE,                                                                                                                " & vbNewLine & _
					" m.name AS FIS, m.Object_ID as FisID,                                                                                                                                   " & vbNewLine & _
					" (CASE WHEN mtv.PropertyID is not null and mtv.Value IS NULL THEN 'In' ELSE mtv.Value END) as FISDirection,                                                             " & vbNewLine & _
					" t.name as Message,                                                                                                                                                     " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name  is null THEN (CASE WHEN sourcePCL.Name is null THEN sourceCL.Name ELSE sourcePCL.Name END) ELSE sourcePoolCL.Name END) as SourcePool,    " & vbNewLine & _
					" (CASE WHEN sourcePoolCL.Name is not null THEN sourcePCL.Name  ELSE null END) as SourceLane,                                                                            " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name  is null THEN (CASE WHEN targetPCL.Name is null THEN targetCL.Name ELSE targetPCL.Name END)ELSE targetPoolCL.Name END) as TargetPool,     " & vbNewLine & _
					" (CASE WHEN targetPoolCL.Name is not null THEN targetPCL.Name  ELSE null END)  as TargetLane,                                                                           " & vbNewLine & _
					" bp5.Name as BusinessProcess, bp4.name as SubProcess_1, bp3.Name as SubProcess_2, bp2.name as SubProcess_3, bp1.name as SubProcess_4, p.Name as SubProcess_5            " & vbNewLine & _
					" FROM t_connector c                                                                                                                                                     " & vbNewLine & _
					" inner join t_connectortag ct on ct.ElementID = c.Connector_ID                                                                                                          " & vbNewLine & _
					"                                                and ct.Property = 'MessageRef'                                                                                          " & vbNewLine & _
					" inner join t_object m on m.ea_guid = ct.VALUE                                                                                                                          " & vbNewLine & _
					" left join t_objectproperties mtv on mtv.Object_ID = m.Object_ID                                                                                                        " & vbNewLine & _
					"                                                              and mtv.Property = 'Atrias::Direction'                                                                    " & vbNewLine & _
					" inner join t_diagramlinks dl on dl.ConnectorID = c.Connector_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d on d.Diagram_ID = dl.DiagramID                                                                                                                  " & vbNewLine & _
					" inner join t_object p on p.Object_ID = d.ParentID                                                                                                                      " & vbNewLine & _
					"                                   and p.Stereotype in ('Activity','BusinessProcess','ArchiMate_BusinessProcess')                                                       " & vbNewLine & _
					" inner join t_objectProperties tvca1 on tvca1.Value = p.ea_guid                                                                                                         " & vbNewLine & _
					"                                                                    and tvca1.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p1 on p1.Object_ID = tvca1.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p1.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do1 on do1.Object_ID = p1.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d1 on d1.Diagram_ID = do1.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp1 on bp1.Object_ID = d1.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca2 on tvca2.Value = bp1.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca2.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p2 on p2.Object_ID = tvca2.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p2.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do2 on do2.Object_ID = p2.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d2 on d2.Diagram_ID = do2.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp2 on bp2.Object_ID = d2.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca3 on tvca3.Value = bp2.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca3.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p3 on p3.Object_ID = tvca3.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p3.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do3 on do3.Object_ID = p3.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d3 on d3.Diagram_ID = do3.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp3 on bp3.Object_ID = d3.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca4 on tvca4.Value = bp3.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca4.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p4 on p4.Object_ID = tvca4.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p4.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do4 on do4.Object_ID = p4.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d4 on d4.Diagram_ID = do4.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp4 on bp4.Object_ID = d4.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_objectProperties tvca5 on tvca5.Value = bp4.ea_guid                                                                                                       " & vbNewLine & _
					"                                                                    and tvca5.Property = 'CalledActivityRef'                                                            " & vbNewLine & _
					" inner join t_object p5 on p5.Object_ID = tvca5.Object_ID                                                                                                               " & vbNewLine & _
					"                                         and p5.Stereotype = 'Activity'                                                                                                 " & vbNewLine & _
					" inner join t_diagramObjects do5 on do5.Object_ID = p5.Object_ID                                                                                                        " & vbNewLine & _
					" inner join t_diagram d5 on d5.Diagram_ID = do5.Diagram_ID                                                                                                              " & vbNewLine & _
					" inner join t_object bp5 on bp5.Object_ID = d5.ParentID                                                                                                                 " & vbNewLine & _
					" inner join t_object source on source.Object_ID = c.Start_Object_ID                                                                                                     " & vbNewLine & _
					" left join t_object  sourceCl on sourceCL.Object_ID = source.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object sourceP on sourceP.Object_ID = source.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object sourcePCL on sourcePCL.Object_ID = sourceP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object sourcePool on sourcePool.Object_ID = sourceP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and sourcePool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object sourcePoolCL on sourcePoolCL.Object_ID = sourcePool.Classifier                                                                                      " & vbNewLine & _
					" inner join t_object target on target.Object_ID = c.End_Object_ID                                                                                                       " & vbNewLine & _
					" left join t_object  targetCl on targetCL.Object_ID = target.Classifier                                                                                                 " & vbNewLine & _
					" left join t_object targetP on targetP.Object_ID = target.ParentID                                                                                                      " & vbNewLine & _
					" left join t_object targetPCL on targetPCL.Object_ID = targetP.Classifier                                                                                               " & vbNewLine & _
					" left join t_object targetPool on targetPool.Object_ID = targetP.ParentID                                                                                               " & vbNewLine & _
					"                                                       and targetPool.Stereotype = 'Pool'                                                                               " & vbNewLine & _
					" left join t_object targetPoolCL on targetPoolCL.Object_ID = targetPool.Classifier                                                                                      " & vbNewLine & _
					" left join t_connector mtt on m.object_ID = mtt.End_object_ID                                                                                                           " & vbNewLine & _
					"                                         and mtt.Connector_Type in( 'Realization', 'Realisation')                                                                       " & vbNewLine & _
					" left join t_object t on mtt.Start_Object_ID = t.Object_ID                                                                                                              " & vbNewLine & _
					"                                         and t.stereotype = 'message'                                                                                                   " & vbNewLine & _
					" ) mtr                                                                                                                                                                  " & vbNewLine & _
					" left join t_objectproperties tvU on tvU.Object_ID = mtr.FisID                                                                                                          " & vbNewLine & _
					"                                    and tvU.Property = 'Atrias::UMIG'                                                                                                   " & vbNewLine & _
					" ) mx                                                                                                                                                                   " & vbNewLine & _
					" ORDER BY Message, FIS, BusinessProcess, SubProcess_1, SubProcess_2, SubProcess_3, SubProcess_4, SubProcess_5                                                           "
end function

function getHeaders()
	dim headers
	set headers = CreateObject("System.Collections.ArrayList")
	headers.add("UMIG")
	headers.add("Message")
	headers.add("FIS")
	headers.add("FIS direction")
	headers.add("Business Process")
	headers.add("SubProcess 1")
	headers.add("SubProcess 2")
	headers.add("SubProcess 3")
	headers.add("SubProcess 4")
	headers.add("SubProcess 5")
	headers.add("Source Pool")
	headers.add("Source Lane")
	headers.add("Target Pool")
	headers.add("Target Lane")
	set getHeaders = headers
end function

main