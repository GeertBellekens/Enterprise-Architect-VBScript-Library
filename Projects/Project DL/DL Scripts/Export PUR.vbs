'[path=\Projects\Project DL\DL Scripts]
'[group=De Lijn Scripts]

Option Explicit

!INC Local Scripts.EAConstants-VBScript
!INC Wrappers.Include
'
' Script Name: Export PUR
' Author: Geert Bellekens
' Purpose: Export the PUR to an Excel file
' Date: 2018-11-21
'
'** guid aanpassen indien de top level package van de PUR wijzigt.
Const PURpackageGUID = "{7E289E8A-07AD-4ebd-9C07-2E06DA1283C1}"


const outPutName = "Export PUR"

Private Sub main
 'create output tab
 Repository.CreateOutputTab outPutName
 Repository.ClearOutput outPutName
 Repository.EnsureOutputVisible outPutName
 'inform user
 Repository.WriteOutput outPutName, now() & " Working on PUR export to Excel. Please hold on...", 0
 'first fix PUR compositions
 fixPurCompositions
 'inform user
 Repository.WriteOutput outPutName, now() & " Getting PUR Data", 0
 ' Get the results
 Dim sqlGetPURExport
 sqlGetPURExport = getSQLPurExport()
 Dim purResults
 Set purResults = getArrayListFromQuery(sqlGetPURExport)
 ' Update results
 updateResults(purResults)
 'inform user
 Repository.WriteOutput outPutName, now() & " Exporting PUR data to Excel", 0
 ' Get the headers
 Dim headers
 Set headers = getHeaders()
 ' Add the headers to the results
 purResults.Insert 0, headers
 ' Create the Excel file
 Dim excelOutput
 Set excelOutput = new ExcelFile
 ' Create a two dimensional array
 Dim excelContents
 excelContents = makeArrayFromArrayLists(purResults)
 ' Add the output to a sheet in excel
 dim ws
 set ws = excelOutput.createTabWithFormulas("PUR", excelContents, true, "TableStyleMedium4")
 'format columns
 formatColumns ws, excelOutput
 'process E2E processes to a new sheet
 'inform user
 Repository.WriteOutput outPutName, now() & " Getting E2E Data", 0
 ' Get the results
 Dim sqlGetE3EExport
 sqlGetE3EExport = getSQLE2Eprocesses()
 Dim E2EResults
 Set E2EResults = getArrayListFromQuery(sqlGetE3EExport)
 ' Update results
 updateResults(E2EResults)
 'inform user
 Repository.WriteOutput outPutName, now() & " Exporting E2E data to Excel", 0
 Set headers = getE2EHeaders()
 ' Add the headers to the results
 E2EResults.Insert 0, headers
 excelContents = makeArrayFromArrayLists(E2EResults)
 ' Add the output to a sheet in excel
 set ws = excelOutput.createTabWithFormulas("E2E processen", excelContents, true, "TableStyleMedium4")
 'format columns
 formatColumns ws, excelOutput
 ' Ask for location, then Save the excel file
 excelOutput.save
 'inform user
 Repository.WriteOutput outPutName, now() & " Finished successfully. Please review result in Excel.", 0
End Sub

function fixPurCompositions()
 'report progess
 Repository.WriteOutput outPutName, now() & " Correcting Composition directions", 0
 dim i
 i = 0
 dim sqlGetData
 sqlGetData = "select c.Connector_ID                                                                      " & vbNewLine & _
    " from t_connector c                                                                        " & vbNewLine & _
    " inner join t_object os on os.Object_ID = c.Start_Object_ID                                " & vbNewLine & _
    "       and os.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')   " & vbNewLine & _
    "       and c.DestIsAggregate = 2                                           " & vbNewLine & _
    " inner join t_object oe on oe.Object_ID = c.End_Object_ID                                  " & vbNewLine & _
    "       and oe.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')   " & vbNewLine & _
    " where c.Stereotype = 'ArchiMate_Composition'                                              " & vbNewLine & _
    " union                                                                                     " & vbNewLine & _
    " select c.Connector_ID                                                                     " & vbNewLine & _
    " from t_connector c                                                                        " & vbNewLine & _
    " inner join t_object os on os.Object_ID = c.Start_Object_ID                                " & vbNewLine & _
    "       and os.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')   " & vbNewLine & _
    "                                                                           " & vbNewLine & _
    " inner join t_object oe on oe.Object_ID = c.End_Object_ID                                  " & vbNewLine & _
    "       and oe.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')   " & vbNewLine & _
    " where c.Stereotype = 'ArchiMate_Composition'                                              " & vbNewLine & _
    " and isnull(c.Direction, '') <> 'Unspecified'                                               "
 dim compositions
 set compositions = getConnectorsFromQuery(sqlGetData)
 dim composition
 dim total
 total = compositions.Count()
 for each composition in compositions
  'report progess
  Repository.WriteOutput outPutName, now() & " Correcting " & i & " of " & total & "  Compositions", 0
  i = i + correctCompositionDirection(composition)
 next
end function


function formatColumns(ws, excelOutput)
 'set width of columns 2-5 to fixed width
 ws.Cells(1,2).ColumnWidth = 20
 ws.Cells(1,3).ColumnWidth = 20
 ws.Cells(1,4).ColumnWidth = 20
 ws.Cells(1,5).ColumnWidth = 20
 'set fontcolor to match background for all but the last level of process (making the name invisible, but still there to be filtered upon)
 dim white
 white = RGB(255,255,255)
 dim grey
 grey = RGB(237,237,237)
 dim color
 dim row
 dim range
 for row = 2 to ws.UsedRange.Rows.Count
  dim toColumn
  if len(ws.Cells(row, 5).Value) > 0  then 
   toColumn = 4
  elseif len(ws.Cells(row, 4).Value) > 0  then 
   toColumn = 3
  elseif len(ws.Cells(row, 3).Value) > 0  then
   toColumn = 2 
  else 
   toColumn = 0
  end if
  'determine color based on odd or even row
  if row Mod 2 = 1 then
   color = white
  else
   color = grey
  end if
  if toColumn > 0 then
   'set font color
   set range = ws.Range(ws.Cells(row, 2), ws.Cells(row, toColumn))
   excelOutput.formatRange range, "default", color, "default", "default", "default", "default"
  end if
 next
end function

Private Function updateResults(purResults)
 Dim row
 For Each row in purResults
  ' Add a single quote to the ref field to get Excel to accept it as text
  row(0) = "'" & row(0)
  'format the description as text
  if len(row(6)) > 0 then
   row(6) = Repository.GetFormatFromField("TXT", row(6))
  end if
  ' Remove the orderField located at index 5
  row.RemoveAt 5
 Next
End Function


Private Function getSQLPurExport()
 getSQLPurExport =   "select distinct bpr.ref, bpr.Category, bpr.processGroup, bpr.macroProcess, bpr.process, bpr.orderField, o.Note as Toelichting                                           " & vbNewLine & _
      " ,tvo.Value as Owner, tvef.Value as EigenaarFunctie, tven.Value as EigenaarNaam, tvbf.Value as BeheerderFunctie, tvbn.Value as BeheerderNaam                            " & vbNewLine & _
      " , tvpuf.Value as ProcesuitvoerderFunctie, tvpun.Value as ProcesuitvoerderNaam                                                                                          " & vbNewLine & _
      " , '=HYPERLINK(""https://webea.apps.addelijn.be?m=1&o=' + SUBSTRING(bpr.ea_guid, 2, 36) + '"")' as webEAURL                                                            " & vbNewLine & _
      " from (                                                                                                                                                                 " & vbNewLine & _
      " select cat.Alias as ref, cat.name as Category, '' as processGroup, '' as macroProcess, '' as process                                                                   " & vbNewLine & _
      " , cat.Alias as orderField, cat.ea_guid as ea_guid, cat.Object_ID                                                                                                       " & vbNewLine & _
      " from t_object cat                                                                                                                                                      " & vbNewLine & _
      " inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                                " & vbNewLine & _
      " inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                                 " & vbNewLine & _
      " where cat.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                                                   " & vbNewLine & _
      " and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                              " & vbNewLine & _
      " and cat.ParentID = 0                                                                                                                                                   " & vbNewLine & _
      " and not exists (                                                                                                                                                       " & vbNewLine & _
      "  select c.Connector_ID from t_connector c                                                                                                                             " & vbNewLine & _
      "  inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                             " & vbNewLine & _
      "       and o.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                                 " & vbNewLine & _
      "  where c.Stereotype = 'ArchiMate_Composition'                                                                                                                         " & vbNewLine & _
      "  and c.End_Object_ID = cat.Object_ID)                                                                                                                                 " & vbNewLine & _
      " union all                                                                                                                                                              " & vbNewLine & _
      " select pg.alias as ref, cat.Name as Category, pg.name as processGroup, '' as macroProcess, '' as process                                                               " & vbNewLine & _
      " , pg.Alias as orderField, pg.ea_guid, pg.Object_ID                                                                                                                     " & vbNewLine & _
      " from t_object pg                                                                                                                                                       " & vbNewLine & _
      " inner join t_connector c on c.End_Object_ID = pg.Object_ID                                                                                                             " & vbNewLine & _
      "        and c.Stereotype = 'ArchiMate_Composition'                                                                                                   " & vbNewLine & _
      " inner join (                                                                                                                                                           " & vbNewLine & _
      "  select cat.Object_ID, cat.Alias, cat.Name                                                                                                                            " & vbNewLine & _
      "  from t_object cat                                                                                                                                                    " & vbNewLine & _
      "  inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                              " & vbNewLine & _
      "  inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                               " & vbNewLine & _
      "  where cat.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                                                 " & vbNewLine & _
      "  and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                            " & vbNewLine & _
      "  and cat.ParentID = 0                                                                                                                                                 " & vbNewLine & _
      "  and not exists (                                                                                                                                                     " & vbNewLine & _
      "   select c.Connector_ID from t_connector c                                                                                                                         " & vbNewLine & _
      "   inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                         " & vbNewLine & _
      "        and o.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                             " & vbNewLine & _
      "   where c.Stereotype = 'ArchiMate_Composition'                                                                                                                     " & vbNewLine & _
      "   and c.End_Object_ID = cat.Object_ID)) cat on cat.Object_ID = c.Start_Object_ID                                                                                   " & vbNewLine & _
      " where                                                                                                                                                                  " & vbNewLine & _
      " pg.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                                                          " & vbNewLine & _
      " union                                                                                                                                                                  " & vbNewLine & _
      " select mp.Alias as ref, pg.Category as Category, pg.Name as processGroup, mp.name as macroProcess, '' as process                                                       " & vbNewLine & _
      " , COALESCE(CASE WHEN TRIM(mp.Alias) = '' THEN null ELSE mp.Alias END, pg.Alias + '-' + FORMAT(isnull(mpp.TPos,0)*10, '000')) as orderField, mp.ea_guid, mp.Object_ID   " & vbNewLine & _
      " from t_object mp                                                                                                                                                       " & vbNewLine & _
      "         inner join t_package mpp on mp.Package_ID =  mpp.Package_ID                                                                                                    " & vbNewLine & _
      " inner join t_connector c on c.End_Object_ID = mp.Object_ID                                                                                                             " & vbNewLine & _
      "        and c.Stereotype = 'ArchiMate_Composition'                                                                                                   " & vbNewLine & _
      " inner join (                                                                                                                                                           " & vbNewLine & _
      "  select pg.Object_ID, pg.Alias, pg.Name, cat.Name as Category                                                                                                         " & vbNewLine & _
      "  from t_object pg                                                                                                                                                     " & vbNewLine & _
      "  inner join t_connector c on c.End_Object_ID = pg.Object_ID                                                                                                           " & vbNewLine & _
      "         and c.Stereotype = 'ArchiMate_Composition'                                                                                               " & vbNewLine & _
      "  inner join (                                                                                                                                                         " & vbNewLine & _
      "   select cat.Object_ID, cat.Alias, cat.Name                                                                                                                        " & vbNewLine & _
      "   from t_object cat                                                                                                                                                " & vbNewLine & _
      "   inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                          " & vbNewLine & _
      "   inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                           " & vbNewLine & _
      "   where cat.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                                             " & vbNewLine & _
      "   and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                        " & vbNewLine & _
      "   and cat.ParentID = 0                                                                                                                                             " & vbNewLine & _
      "   and not exists (                                                                                                                                                 " & vbNewLine & _
      "    select c.Connector_ID from t_connector c                                                                                                                     " & vbNewLine & _
      "    inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                     " & vbNewLine & _
      "         and o.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                         " & vbNewLine & _
      "    where c.Stereotype = 'ArchiMate_Composition'                                                                                                                 " & vbNewLine & _
      "    and c.End_Object_ID = cat.Object_ID)                                                                                                                         " & vbNewLine & _
      "    ) cat on cat.Object_ID = c.Start_Object_ID                                                                                                                   " & vbNewLine & _
      "   ) pg on pg.Object_ID = c.Start_Object_ID                                                                                                                         " & vbNewLine & _
      " where                                                                                                                                                                  " & vbNewLine & _
      " mp.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                                                          " & vbNewLine & _
      " union                                                                                                                                                                  " & vbNewLine & _
      " select pr.Alias as ref, mp.Category as Category, mp.ProcessGroup as processGroup, mp.Name as macroProcess, pr.Name as process                                          " & vbNewLine & _
      " , COALESCE( CASE WHEN TRIM(pr.Alias) = '' THEN null ELSE pr.Alias END, COALESCE( CASE WHEN TRIM(mp.mpAlias) = '' THEN null ELSE mp.mpAlias END,                        " & vbNewLine & _
      "   mp.pgAlias + '-' + FORMAT(isnull(mp.mppTPos,0)*10, '000')) + '-' + FORMAT(isnull(pr.TPos, 0)*10, '000')) as orderField                                               " & vbNewLine & _
      " , pr.ea_guid, pr.Object_ID                                                                                                                                             " & vbNewLine & _
      " from t_object pr                                                                                                                                                       " & vbNewLine & _
      " inner join t_connector c on c.End_Object_ID = pr.Object_ID                                                                                                             " & vbNewLine & _
      "        and c.Stereotype = 'ArchiMate_Composition'                                                                                                   " & vbNewLine & _
      " inner join (                                                                                                                                                           " & vbNewLine & _
      "  select mp.Object_ID, mpp.TPos as mppTPos, mp.Alias as mpAlias, mp.Name, pg.Alias as pgAlias, pg.Category, pg.Name as ProcessGroup                                    " & vbNewLine & _
      "  from t_object mp                                                                                                                                                     " & vbNewLine & _
      "   inner join t_package mpp on mp.Package_ID =  mpp.Package_ID                                                                                                          " & vbNewLine & _
      "  inner join t_connector c on c.End_Object_ID = mp.Object_ID                                                                                                           " & vbNewLine & _
      "         and c.Stereotype = 'ArchiMate_Composition'                                                                                               " & vbNewLine & _
      "  inner join (                                                                                                                                                         " & vbNewLine & _
      "   select pg.Object_ID, pg.Alias, pg.Name, cat.Name as Category                                                                                                     " & vbNewLine & _
      "   from t_object pg                                                                                                                                                 " & vbNewLine & _
      "   inner join t_connector c on c.End_Object_ID = pg.Object_ID                                                                                                       " & vbNewLine & _
      "          and c.Stereotype = 'ArchiMate_Composition'                                                                                           " & vbNewLine & _
      "   inner join (                                                                                                                                                     " & vbNewLine & _
      "    select cat.Object_ID, cat.Alias, cat.Name                                                                                                                    " & vbNewLine & _
      "    from t_object cat                                                                                                                                            " & vbNewLine & _
      "    inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                      " & vbNewLine & _
      "    inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                       " & vbNewLine & _
      "    where cat.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                                         " & vbNewLine & _
      "    and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                    " & vbNewLine & _
      "    and cat.ParentID = 0                                                                                                                                         " & vbNewLine & _
      "    and not exists (                                                                                                                                             " & vbNewLine & _
      "     select c.Connector_ID from t_connector c                                                                                                                 " & vbNewLine & _
      "     inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                 " & vbNewLine & _
      "          and o.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces')                                                                     " & vbNewLine & _
      "     where c.Stereotype = 'ArchiMate_Composition'                                                                                                             " & vbNewLine & _
      "     and c.End_Object_ID = cat.Object_ID)                                                                                                                     " & vbNewLine & _
      "     ) cat on cat.Object_ID = c.Start_Object_ID                                                                                                               " & vbNewLine & _
      "    ) pg on pg.Object_ID = c.Start_Object_ID                                                                                                                     " & vbNewLine & _
      "   ) mp on mp.Object_ID = c.Start_Object_ID                                                                                                                         " & vbNewLine & _
      " where                                                                                                                                                                  " & vbNewLine & _
      " pr.Stereotype in ('ArchiMate_BusinessProcess','DeLijnProces') ) bpr                                                                                                    " & vbNewLine & _
      " inner join t_object o on o.Object_ID = bpr.Object_ID                                                                                                                   " & vbNewLine & _
      " left join t_objectproperties tvo on bpr.Object_ID = tvo.Object_ID                                                                                                      " & vbNewLine & _
      "           and tvo.Property = 'eigenaar directie'                                                                                              " & vbNewLine & _
      " left join t_objectproperties tvef on bpr.Object_ID = tvef.Object_ID                                                                                                    " & vbNewLine & _
      "           and tvef.Property = 'eigenaar functie'                                                                                              " & vbNewLine & _
      " left join t_objectproperties tven on bpr.Object_ID = tven.Object_ID                                                                                                    " & vbNewLine & _
      "           and tven.Property = 'eigenaar naam'                                                                                                 " & vbNewLine & _
      " left join t_objectproperties tvbf on bpr.Object_ID = tvbf.Object_ID                                                                                                    " & vbNewLine & _
      "           and tvbf.Property = 'beheerder functie'                                                                                             " & vbNewLine & _
      " left join t_objectproperties tvbn on bpr.Object_ID = tvbn.Object_ID                                                                                                    " & vbNewLine & _
      "           and tvbn.Property = 'beheerder naam'                                                                                                " & vbNewLine & _
      " left join t_objectproperties tvrv on bpr.Object_ID = tvrv.Object_ID                                                                                                    " & vbNewLine & _
      "           and tvrv.Property = 'raakvlakken'                                                                                                   " & vbNewLine & _
      " left join t_objectproperties tvpuf on bpr.Object_ID = tvpuf.Object_ID                                                                                                  " & vbNewLine & _
      "           and tvpuf.Property = 'procesuitvoerder functie'                                                                                     " & vbNewLine & _
      " left join t_objectproperties tvpun on bpr.Object_ID = tvpun.Object_ID                                                                                                  " & vbNewLine & _
      "           and tvpun.Property = 'procesuitvoerder naam'                                                                                        " & vbNewLine & _
      " order by bpr.orderField, macroProcess, process                                                                                                                         "
End Function

Private Function getHeaders()
 dim headers
 set headers = CreateObject("System.Collections.ArrayList")
 headers.add("Ref")
 headers.add("Categorie")
 headers.Add("Proces Groep")
 headers.Add("Macro Proces")
 headers.Add("Proces")
 headers.Add("Toelichting")
 headers.Add("Eigenaar Directie")
 headers.Add("Eigenaar Functie")
 headers.Add("Eigenaar Naam")
 headers.Add("Beheerder Functie")
 headers.Add("Beheerder Naam")
 headers.Add("Procesuitvoerder Functie")
 headers.Add("Procesuitvoerder Naam")
 headers.Add("webEAUrl")
 set getHeaders = headers
End Function

function getSQLE2Eprocesses()
 getSQLE2Eprocesses = "select distinct bpr.ref, bpr.Category, bpr.processGroup, bpr.macroProcess, bpr.process, bpr.orderField, o.Note as Toelichting                                                                 " & vbNewLine & _
     " ,tvo.Value as BusinessOwner, tvsp.Value as SPOCProces                                                                                                                                             " & vbNewLine & _
     " , '=HYPERLINK(""https://webea.apps.addelijn.be?m=1&o=' + SUBSTRING(bpr.ea_guid, 2, 36) + '"")' as webEAURL                                                                                       " & vbNewLine & _
     " , '=HYPERLINK(""https://webea.apps.addelijn.be?m=1&o=' + SUBSTRING(od.ea_guid, 2, 36) + '"")' as OverzichtURL                                                                                    " & vbNewLine & _
     " , '=HYPERLINK(""https://webea.apps.addelijn.be?m=1&o=' + SUBSTRING(dd.ea_guid, 2, 36) + '"")' as DetailUrl                                                                                       " & vbNewLine & _
     " from (                                                                                                                                                                                            " & vbNewLine & _
     " select cat.Alias as ref, cat.name as Category, '' as processGroup, '' as macroProcess, '' as process                                                                                              " & vbNewLine & _
     " , cat.Alias as orderField, cat.ea_guid as ea_guid, cat.Object_ID                                                                                                                                  " & vbNewLine & _
     " from t_object cat                                                                                                                                                                                 " & vbNewLine & _
     " inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                                                           " & vbNewLine & _
     " inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                                                            " & vbNewLine & _
     " where cat.Stereotype = 'ArchiMate_ValueStream'                                                                                                                                                    " & vbNewLine & _
     " and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                                                         " & vbNewLine & _
     " and cat.ParentID = 0                                                                                                                                                                              " & vbNewLine & _
     " and not exists (                                                                                                                                                                                  " & vbNewLine & _
     "  select c.Connector_ID from t_connector c                                                                                                                                                        " & vbNewLine & _
     "  inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                                                        " & vbNewLine & _
     "       and o.Stereotype = 'ArchiMate_ValueStream'                                                                                                                                  " & vbNewLine & _
     "  where c.Stereotype = 'ArchiMate_Composition'                                                                                                                                                    " & vbNewLine & _
     "  and c.End_Object_ID = cat.Object_ID)                                                                                                                                                            " & vbNewLine & _
     " union all                                                                                                                                                                                         " & vbNewLine & _
     " select pg.alias as ref, cat.Name as Category, pg.name as processGroup, '' as macroProcess, '' as process                                                                                          " & vbNewLine & _
     " , pg.Alias as orderField, pg.ea_guid, pg.Object_ID                                                                                                                                                " & vbNewLine & _
     " from t_object pg                                                                                                                                                                                  " & vbNewLine & _
     " inner join t_connector c on c.End_Object_ID = pg.Object_ID                                                                                                                                        " & vbNewLine & _
     "        and c.Stereotype = 'ArchiMate_Composition'                                                                                                                              " & vbNewLine & _
     " inner join (                                                                                                                                                                                      " & vbNewLine & _
     "  select cat.Object_ID, cat.Alias, cat.Name                                                                                                                                                       " & vbNewLine & _
     "  from t_object cat                                                                                                                                                                               " & vbNewLine & _
     "  inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                                                         " & vbNewLine & _
     "  inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                                                          " & vbNewLine & _
     "  where cat.Stereotype = 'ArchiMate_ValueStream'                                                                                                                                                  " & vbNewLine & _
     "  and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                                                       " & vbNewLine & _
     "  and cat.ParentID = 0                                                                                                                                                                            " & vbNewLine & _
     "  and not exists (                                                                                                                                                                                " & vbNewLine & _
     "   select c.Connector_ID from t_connector c                                                                                                                                                    " & vbNewLine & _
     "   inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                                                    " & vbNewLine & _
     "        and o.Stereotype = 'ArchiMate_ValueStream'                                                                                                                              " & vbNewLine & _
     "   where c.Stereotype = 'ArchiMate_Composition'                                                                                                                                                " & vbNewLine & _
     "   and c.End_Object_ID = cat.Object_ID)) cat on cat.Object_ID = c.Start_Object_ID                                                                                                              " & vbNewLine & _
     " where                                                                                                                                                                                             " & vbNewLine & _
     " pg.Stereotype = 'ArchiMate_ValueStream'                                                                                                                                                           " & vbNewLine & _
     " union                                                                                                                                                                                             " & vbNewLine & _
     " select mp.Alias as ref, pg.Category as Category, pg.Name as processGroup, mp.name as macroProcess, '' as process                                                                                  " & vbNewLine & _
     " , COALESCE(CASE WHEN TRIM(mp.Alias) = '' THEN null ELSE mp.Alias END, pg.Alias + '-' + FORMAT(isnull(mpp.TPos,0)*10, '000')) as orderField, mp.ea_guid, mp.Object_ID                              " & vbNewLine & _
     " from t_object mp                                                                                                                                                                                  " & vbNewLine & _
     "         inner join t_package mpp on mp.Package_ID =  mpp.Package_ID                                                                                                                               " & vbNewLine & _
     " inner join t_connector c on c.End_Object_ID = mp.Object_ID                                                                                                                                        " & vbNewLine & _
     "        and c.Stereotype = 'ArchiMate_Composition'                                                                                                                              " & vbNewLine & _
     " inner join (                                                                                                                                                                                      " & vbNewLine & _
     "  select pg.Object_ID, pg.Alias, pg.Name, cat.Name as Category                                                                                                                                    " & vbNewLine & _
     "  from t_object pg                                                                                                                                                                                " & vbNewLine & _
     "  inner join t_connector c on c.End_Object_ID = pg.Object_ID                                                                                                                                      " & vbNewLine & _
     "         and c.Stereotype = 'ArchiMate_Composition'                                                                                                                          " & vbNewLine & _
     "  inner join (                                                                                                                                                                                    " & vbNewLine & _
     "   select cat.Object_ID, cat.Alias, cat.Name                                                                                                                                                   " & vbNewLine & _
     "   from t_object cat                                                                                                                                                                           " & vbNewLine & _
     "   inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                                                     " & vbNewLine & _
     "   inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                                                      " & vbNewLine & _
     "   where cat.Stereotype = 'ArchiMate_ValueStream'                                                                                                                                              " & vbNewLine & _
     "   and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                                                   " & vbNewLine & _
     "   and cat.ParentID = 0                                                                                                                                                                        " & vbNewLine & _
     "   and not exists (                                                                                                                                                                            " & vbNewLine & _
     "    select c.Connector_ID from t_connector c                                                                                                                                                " & vbNewLine & _
     "    inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                                                " & vbNewLine & _
     "         and o.Stereotype = 'ArchiMate_ValueStream'                                                                                                                          " & vbNewLine & _
     "    where c.Stereotype = 'ArchiMate_Composition'                                                                                                                                            " & vbNewLine & _
     "    and c.End_Object_ID = cat.Object_ID)                                                                                                                                                    " & vbNewLine & _
     "    ) cat on cat.Object_ID = c.Start_Object_ID                                                                                                                                              " & vbNewLine & _
     "   ) pg on pg.Object_ID = c.Start_Object_ID                                                                                                                                                    " & vbNewLine & _
     " where                                                                                                                                                                                             " & vbNewLine & _
     " mp.Stereotype = 'ArchiMate_ValueStream'                                                                                                                                                           " & vbNewLine & _
     " union                                                                                                                                                                                             " & vbNewLine & _
     " select pr.Alias as ref, mp.Category as Category, mp.ProcessGroup as processGroup, mp.Name as macroProcess, pr.Name as process                                                                     " & vbNewLine & _
     " , COALESCE( CASE WHEN TRIM(pr.Alias) = '' THEN null ELSE pr.Alias END, COALESCE( CASE WHEN TRIM(mp.mpAlias) = '' THEN null ELSE mp.mpAlias END,                                                   " & vbNewLine & _
     "   mp.pgAlias + '-' + FORMAT(isnull(mp.mppTPos,0)*10, '000')) + '-' + FORMAT(isnull(pr.TPos, 0)*10, '000')) as orderField                                                                          " & vbNewLine & _
     " , pr.ea_guid, pr.Object_ID                                                                                                                                                                        " & vbNewLine & _
     " from t_object pr                                                                                                                                                                                  " & vbNewLine & _
     " inner join t_connector c on c.End_Object_ID = pr.Object_ID                                                                                                                                        " & vbNewLine & _
     "        and c.Stereotype = 'ArchiMate_Composition'                                                                                                                              " & vbNewLine & _
     " inner join (                                                                                                                                                                                      " & vbNewLine & _
     "  select mp.Object_ID, mpp.TPos as mppTPos, mp.Alias as mpAlias, mp.Name, pg.Alias as pgAlias, pg.Category, pg.Name as ProcessGroup                                                               " & vbNewLine & _
     "  from t_object mp                                                                                                                                                                                " & vbNewLine & _
     "   inner join t_package mpp on mp.Package_ID =  mpp.Package_ID                                                                                                                                     " & vbNewLine & _
     "  inner join t_connector c on c.End_Object_ID = mp.Object_ID                                                                                                                                      " & vbNewLine & _
     "         and c.Stereotype = 'ArchiMate_Composition'                                                                                                                          " & vbNewLine & _
     "  inner join (                                                                                                                                                                                    " & vbNewLine & _
     "   select pg.Object_ID, pg.Alias, pg.Name, cat.Name as Category                                                                                                                                " & vbNewLine & _
     "   from t_object pg                                                                                                                                                                            " & vbNewLine & _
     "   inner join t_connector c on c.End_Object_ID = pg.Object_ID                                                                                                                                  " & vbNewLine & _
     "          and c.Stereotype = 'ArchiMate_Composition'                                                                                                                      " & vbNewLine & _
     "   inner join (                                                                                                                                                                                " & vbNewLine & _
     "    select cat.Object_ID, cat.Alias, cat.Name                                                                                                                                               " & vbNewLine & _
     "    from t_object cat                                                                                                                                                                       " & vbNewLine & _
     "    inner join t_package p on cat.Package_ID = p.Package_ID                                                                                                                                 " & vbNewLine & _
     "    inner join t_package pp on pp.Package_ID = p.Parent_ID                                                                                                                                  " & vbNewLine & _
     "    where cat.Stereotype = 'ArchiMate_ValueStream'                                                                                                                                          " & vbNewLine & _
     "    and pp.ea_guid = '" & PURpackageGUID & "'                                                                                                                                               " & vbNewLine & _
     "    and cat.ParentID = 0                                                                                                                                                                    " & vbNewLine & _
     "    and not exists (                                                                                                                                                                        " & vbNewLine & _
     "     select c.Connector_ID from t_connector c                                                                                                                                            " & vbNewLine & _
     "     inner join t_object o on c.Start_Object_ID = o.Object_ID                                                                                                                            " & vbNewLine & _
     "          and o.Stereotype = 'ArchiMate_ValueStream'                                                                                                                      " & vbNewLine & _
     "     where c.Stereotype = 'ArchiMate_Composition'                                                                                                                                        " & vbNewLine & _
     "     and c.End_Object_ID = cat.Object_ID)                                                                                                                                                " & vbNewLine & _
     "     ) cat on cat.Object_ID = c.Start_Object_ID                                                                                                                                          " & vbNewLine & _
     "    ) pg on pg.Object_ID = c.Start_Object_ID                                                                                                                                                " & vbNewLine & _
     "   ) mp on mp.Object_ID = c.Start_Object_ID                                                                                                                                                    " & vbNewLine & _
     " where                                                                                                                                                                                             " & vbNewLine & _
     " pr.Stereotype = 'ArchiMate_ValueStream' ) bpr                                                                                                                                                     " & vbNewLine & _
     " inner join t_object o on o.Object_ID = bpr.Object_ID                                                                                                                                              " & vbNewLine & _
     " left join t_objectproperties tvo on bpr.Object_ID = tvo.Object_ID                                                                                                                                 " & vbNewLine & _
     "           and tvo.Property = 'Business Owner'                                                                                                                            " & vbNewLine & _
     " left join t_objectproperties tvsp on bpr.Object_ID = tvsp.Object_ID                                                                                                                               " & vbNewLine & _
     "           and tvsp.Property = 'SPOC Proces'                                                                                                                       " & vbNewLine & _
     " left join t_diagram od on od.Diagram_ID = o.PDATA1                                                                                                                                                " & vbNewLine & _
     "       and od.ParentID = o.Object_ID                                                                                                                                               " & vbNewLine & _
     " left join t_diagram dd on dd.ParentID = o.Object_ID                                                                                                                                               " & vbNewLine & _
     "        and dd.Diagram_ID <> isnull(od.Diagram_ID, 0)                                                                                                                              " & vbNewLine & _
     " order by bpr.orderField, macroProcess, process                                                                                                                                                    "
end function 

Private Function getE2EHeaders()
 dim headers
 set headers = CreateObject("System.Collections.ArrayList")
 headers.add("Ref")
 headers.add("Categorie")
 headers.Add("Proces Groep")
 headers.Add("Macro Proces")
 headers.Add("Proces")
 headers.Add("Toelichting")
 headers.Add("Business Owner")
 headers.Add("SPOC proces")
 headers.Add("webEAUrl")
 headers.Add("webEAUrl Overzicht Diagram")
 headers.Add("webEAUrl Detail Diagram")
 set getE2EHeaders = headers
End Function

function correctCompositionDirection(relation)
 'return 0 by default
 correctCompositionDirection = 0
 if lcase(relation.Stereotype) = lcase("ArchiMate_Composition") then
  'check aggregationKind
  if relation.SupplierEnd.Aggregation <> 0 _
   and relation.ClientEnd.Aggregation = 0 then
   'switch source and target
   'switch ID's
   dim tempID
   tempID = relation.ClientID
   relation.ClientID = relation.SupplierID
   relation.SupplierID = tempID
   'switch Ends
   switchRelationEnds relation
   'save relation
   relation.Update
   'return 1 as indicator of success
   correctCompositionDirection = 1
  end if
  if relation.Direction <> "Unspecified" then
   'make sure there are no arrows
   relation.Direction = "Unspecified"
   'save relation
   relation.Update
   'return 1 as indicator of success
   correctCompositionDirection = 1
  end if
 end if
end function

function switchRelationEnds (relation)
 dim tempVar
 tempvar = relation.ClientEnd.Aggregation
 relation.ClientEnd.Aggregation = relation.SupplierEnd.Aggregation
 relation.SupplierEnd.Aggregation       = tempvar
 tempvar = relation.ClientEnd.Alias
 relation.ClientEnd.Alias = relation.SupplierEnd.Alias
 relation.SupplierEnd.Alias             = tempvar
 tempvar = relation.ClientEnd.AllowDuplicates
 relation.ClientEnd.AllowDuplicates = relation.SupplierEnd.AllowDuplicates
 relation.SupplierEnd.AllowDuplicates   = tempvar
 tempvar = relation.ClientEnd.Cardinality
 relation.ClientEnd.Cardinality = relation.SupplierEnd.Cardinality
 relation.SupplierEnd.Cardinality       = tempvar
 tempvar = relation.ClientEnd.Constraint
 relation.ClientEnd.Constraint = relation.SupplierEnd.Constraint
 relation.SupplierEnd.Constraint        = tempvar
 tempvar = relation.ClientEnd.Containment
 relation.ClientEnd.Containment = relation.SupplierEnd.Containment
 relation.SupplierEnd.Containment       = tempvar
 tempvar = relation.ClientEnd.Derived
 relation.ClientEnd.Derived = relation.SupplierEnd.Derived
 relation.SupplierEnd.Derived           = tempvar
 tempvar = relation.ClientEnd.DerivedUnion
 relation.ClientEnd.DerivedUnion = relation.SupplierEnd.DerivedUnion
 relation.SupplierEnd.DerivedUnion      = tempvar
 tempvar = relation.ClientEnd.IsChangeable
 relation.ClientEnd.IsChangeable = relation.SupplierEnd.IsChangeable
 relation.SupplierEnd.IsChangeable      = tempvar
 tempvar = relation.ClientEnd.IsNavigable
 relation.ClientEnd.IsNavigable = relation.SupplierEnd.IsNavigable
 relation.SupplierEnd.IsNavigable       = tempvar
 tempvar = relation.ClientEnd.Navigable
 relation.ClientEnd.Navigable = relation.SupplierEnd.Navigable
 relation.SupplierEnd.Navigable         = tempvar
 tempvar = relation.ClientEnd.Ordering
 relation.ClientEnd.Ordering = relation.SupplierEnd.Ordering
 relation.SupplierEnd.Ordering          = tempvar
 tempvar = relation.ClientEnd.OwnedByClassifier
 relation.ClientEnd.OwnedByClassifier = relation.SupplierEnd.OwnedByClassifier
 relation.SupplierEnd.OwnedByClassifier = tempvar
 tempvar = relation.ClientEnd.Qualifier
 relation.ClientEnd.Qualifier = relation.SupplierEnd.Qualifier
 relation.SupplierEnd.Qualifier         = tempvar
 tempvar = relation.ClientEnd.Role
 relation.ClientEnd.Role = relation.SupplierEnd.Role
 relation.SupplierEnd.Role              = tempvar
 tempvar = relation.ClientEnd.RoleNote
 relation.ClientEnd.RoleNote = relation.SupplierEnd.RoleNote
 relation.SupplierEnd.RoleNote          = tempvar
 tempvar = relation.ClientEnd.RoleType
 relation.ClientEnd.RoleType = relation.SupplierEnd.RoleType
 relation.SupplierEnd.RoleType          = tempvar
 tempvar = relation.ClientEnd.Stereotype
 relation.ClientEnd.Stereotype = relation.SupplierEnd.Stereotype
 relation.SupplierEnd.Stereotype        = tempvar
 tempvar = relation.ClientEnd.StereotypeEx
 relation.ClientEnd.StereotypeEx = relation.SupplierEnd.StereotypeEx
 relation.SupplierEnd.StereotypeEx      = tempvar
' tempvar = relation.ClientEnd.TaggedValues
' relation.ClientEnd.TaggedValues = relation.SupplierEnd.TaggedValues
' relation.SupplierEnd.TaggedValues      = tempvar
 tempvar = relation.ClientEnd.Visibility
 relation.ClientEnd.Visibility = relation.SupplierEnd.Visibility
 relation.SupplierEnd.Visibility        = tempvar
 relation.ClientEnd.Update
 relation.SupplierEnd.Update
end function

Main
