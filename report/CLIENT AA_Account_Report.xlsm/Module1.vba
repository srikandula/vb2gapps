Attribute VB_Name = "Module1"
Sub pivot()
Attribute pivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' pivot Macro
'

'
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet43!R1C1:R47C20", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Sheet44!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=xlPivotTableVersion14
    Sheets("Sheet44").Select
    Cells(3, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
End Sub
Sub selections()
Attribute selections.VB_ProcData.VB_Invoke_Func = " \n14"
'
' selections Macro
'

'
    Range("M1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
End Sub
