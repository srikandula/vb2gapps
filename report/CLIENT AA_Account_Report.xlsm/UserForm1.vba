Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{0ABC0F62-C9A8-4F2C-9E0A-087937E32312}{9BF30DEA-3FD6-4E3B-B093-C963D3D3DEFD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CommandButton1_Click()

End Sub

Private Sub Management_Entity_Click()

End Sub

Private Sub Run_Button_Click()

End Sub

Private Sub btn_Del_Rpts_Start_Over_Click()

Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
If ws.Name Like "*Sheet*" Then ws.Delete
Next ws
Application.DisplayAlerts = True
Worksheets("ME_LE_Report").Activate
ActiveSheet.Range("A1").Select

'Clear listbox selection after running query
    For i = 0 To List42.ListCount - 1
    List42.Selected(i) = False
    Next i

    For i = 0 To List82.ListCount - 1
    List82.Selected(i) = False
    Next i
    
    For i = 0 To List13.ListCount - 1
    List13.Selected(i) = False
    Next i
    
    For i = 0 To List91.ListCount - 1
    List91.Selected(i) = False
    Next i

End Sub

Private Sub btn_Report_Pivot_Click()
On Error GoTo Err_btn_Report_Pivot_Click
    Dim mws As Worksheet
    Dim ws As Worksheet
    Dim copyRange As Range
    Dim i As Integer
    Dim p As Integer
    Dim strWhere_ME As Variant
    Dim strIN_ME As String
    Dim strWhere_LE As Variant
    Dim strIN_LE As String
    Dim strWhere_AT As Variant
    Dim strIN_AT As String
    Dim strWhere_AN As Variant
    Dim strIN_AN As String
    Dim varItem As Variant
    Dim data_sht As Worksheet
    Dim pivot_sht As Worksheet
    Dim startpoint As Range
    Dim datarange As Variant
    Dim newrange As Range
    
    'Set master worksheet
    Set mws = Worksheets("ME_LE_Report")
    
    flgSelectAll = 0
    flgAll_ME = 0
    flgAll_LE = 0
    flgAll_AT = 0
    flgAll_AN = 0
    
    For i = 0 To List42.ListCount - 1
        If List42.Selected(i) And List42.Column(0, i) = "All" Then
                flgAll_ME = flgAll_ME + 1
                For p = 1 To List42.ListCount - 1
                strIN_ME = strIN_ME & Left(List42.Column(0, p), 6) & ","
                Next p
                Exit For
        ElseIf List42.Selected(i) Then strIN_ME = strIN_ME & Left(List42.Column(0, i), 6) & ","
        End If
    Next i
    
  For i = 0 To List82.ListCount - 1
        If List82.Selected(i) And List82.Column(0, i) = "All" Then
                flgAll_LE = flgAll_LE + 1
                For p = 1 To List82.ListCount - 1
                strIN_LE = strIN_LE & Left(List82.Column(0, p), 3) & ","
                Next p
                Exit For
        ElseIf List82.Selected(i) Then strIN_LE = strIN_LE & Left(List82.Column(0, i), 3) & ","
        End If
    Next i
    
   For i = 0 To List13.ListCount - 1
        If List13.Selected(i) And List13.Column(0, i) = "All" Then
                flgAll_AT = flgAll_AT + 1
                For p = 1 To List13.ListCount - 1
                strIN_AT = strIN_AT & List13.Column(0, p) & ","
                Next p
                Exit For
        ElseIf List13.Selected(i) Then strIN_AT = strIN_AT & List13.Column(0, i) & ","
        End If
    Next i
    
  For i = 0 To List91.ListCount - 1
        If List91.Selected(i) And List91.Column(0, i) = "All" Then
                flgAll_AN = flgAll_AN + 1
                For p = 1 To List91.ListCount - 1
                strIN_AN = strIN_AN & Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) & ","
                Next p
                Exit For
        ElseIf List91.Selected(i) Then strIN_AN = strIN_AN & Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) & ","
        End If
    Next i
    
    flgSelectAll = flgAll_ME + flgAll_LE + flgAll_AT + flgAll_AN
    If flgSelectAll = 4 Then
     GoTo Err_Select_All_Click
    End If
    
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_ME = Left(strIN_ME, Len(strIN_ME) - 1)
               
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_LE = Left(strIN_LE, Len(strIN_LE) - 1)
               
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AT = Left(strIN_AT, Len(strIN_AT) - 1)
               
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AN = Left(strIN_AN, Len(strIN_AN) - 1)
    
    With mws
    mws.Activate
    .AutoFilterMode = False
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=3, Criteria1:=Split(strWhere_ME, ","), Operator:=xlFilterValues
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=6, Criteria1:=Split(strWhere_AT, ","), Operator:=xlFilterValues
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=10, Criteria1:=Split(strWhere_AN, ","), Operator:=xlFilterValues
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=13, Criteria1:=Split(strWhere_LE, ","), Operator:=xlFilterValues
'    With .AutoFilter.Range
'        On Error Resume Next ' if none selected
'        .Offset(1).Resize(.Rows.Count - 1, Columns.Count).SpecialCells(xlCellTypeVisible).Select
'        On Error GoTo 0
'    End With
    .AutoFilterMode = False
    End With
    
    'Create new sheet
    Sheets.Add After:=Sheets("ME_LE_Report")
    
    'Copy and Paste filtered data to new sheet
    Set ws = Sheets(mws.Index + 1)
    mws.Activate
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set copyRange = Selection
    copyRange.Copy ws.Range("A1")
    ws.Columns("A:R").AutoFit
    
    'Clear filters from report
    mws.ListObjects(1).Range.AutoFilter Field:=3
    mws.ListObjects(1).Range.AutoFilter Field:=6
    mws.ListObjects(1).Range.AutoFilter Field:=10
    mws.ListObjects(1).Range.AutoFilter Field:=13

    'Retrieve data range
    ws.Activate
    ws.Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set datarange = Selection

    'Create Pivot Table
    Sheets.Add After:=ws
    Set pivot_sht = Sheets(ws.Index + 1)
    Application.WindowState = xlMaximized
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=datarange, _
    Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=pivot_sht.Range("A3"), TableName:="Report_Pivot", DefaultVersion _
        :=xlPivotTableVersion14
        
    pivot_sht.Select
    Cells(3, 1).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("LEGAL_ENTITY")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("AT_Sort")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("Acct_Sort")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT")
        .Orientation = xlRowField
        .Position = 5
    End With
    ActiveSheet.PivotTables("Report_Pivot").AddDataField ActiveSheet.PivotTables( _
        "Report_Pivot").PivotFields("MARS_AMOUNT_IN"), "Sum of MARS_AMOUNT_IN", xlSum
    ActiveSheet.PivotTables("Report_Pivot").AddDataField ActiveSheet.PivotTables( _
        "Report_Pivot").PivotFields("MARS_AMOUNT_IN"), "Sum of MARS_AMOUNT_IN2", xlSum
        ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN")
        .NumberFormat = "$#,##0.00"
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN2")
        .Calculation = xlPercentOfRow
        .NumberFormat = "0.00%"
    End With
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN"). _
        Caption = "Sum of LE In"
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN2"). _
        Caption = "% of LE In"
    Application.WindowState = xlMaximized
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("AT_Sort").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("AT_Sort").LayoutForm _
        = xlTabular
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE").LayoutForm _
        = xlTabular
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("Acct_Sort").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("Acct_Sort").LayoutForm _
        = xlTabular
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC"). _
        LayoutForm = xlTabular
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT").LayoutForm = _
        xlTabular
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE"). _
        Caption = "Account Type"
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC"). _
        Caption = "Financials Description"
    ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT"). _
        Caption = "Account"
    ActiveWorkbook.ShowPivotTableFieldList = False
    ActiveSheet.PivotTables("Report_Pivot").TableStyle2 = ""
    'Column headings bold
    Range("A7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Selection.End(xlToRight).Select
    'Create filters
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("ME")
       .Orientation = xlPageField
       .Position = 1
       .Caption = "Management Entity"
       .EnableMultiplePageItems = True
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("PERIOD_NAME")
        .Orientation = xlPageField
        .Position = 1
        .Caption = "Period"
        .EnableMultiplePageItems = True
    End With
    With ActiveSheet.PivotTables("Report_Pivot").PivotFields("ME_COUNTRY_PER_MARS")
        .Orientation = xlPageField
        .Position = 1
        .Caption = "Country"
        .EnableMultiplePageItems = True
    End With
    'Format column headings
    Range("XFD7").Select
    Selection.End(xlToLeft).Select
    Range("B7", Selection).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    'Format column headings
    Range("XFD6").Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    'Format grand total
    Range("A10000").Select
    Selection.End(xlUp).Select
    Selection.Font.Bold = True
    'Replace underscores in column headings
    Range("E7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Replace What:="_", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Don't show gridlines
    ActiveWindow.DisplayGridlines = False
    'Clean up error values
    With ActiveSheet.PivotTables("Report_Pivot")
        .DisplayErrorString = True
        .ErrorString = "Sums to Zero"
    End With
    'Clean up visuals
    ActiveWindow.Zoom = 90
    ActiveSheet.Columns.AutoFit
    Cells.Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    'Hide sort columns
    Columns("A").Hidden = True
    Columns("C").Hidden = True
    'Assign and format filter labels
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Range("D1:D3").Select
    Selection.FillDown
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    'Hide form when complete
    UserForm1.Hide
    Range("A1").Select
    
    
Exit_btn_Report_Pivot_Click:
    Exit Sub

Err_Select_All_Click:
MsgBox "You may not select ALL in every list. The full report is in tab ""ME_LE_Report.""" _
               , , "Select a value from each box."
        GoTo Exit_btn_Report_Pivot_Click

Err_btn_Report_Pivot_Click:

    If Err.Number = 5 Then
        MsgBox "You must make a selection(s) from each list" _
               , , "Selection Required !"
        Resume Exit_btn_Report_Pivot_Click
    ElseIf Err.Number = 1004 Then
        MsgBox "Pleaes re-select as cetain selection(s) don't exist" _
               , , "Re-Selection Required !"
        Resume Exit_btn_Report_Pivot_Click
       'Write out the error and exit the sub
    Else: MsgBox Err.Description
        Resume Exit_btn_Report_Pivot_Click
    End If

End Sub

Private Sub btn_Reset_Form_Click()

Worksheets("ME_LE_Report").Activate

'Clear filters from report
ActiveSheet.ListObjects(1).Range.AutoFilter Field:=3
ActiveSheet.ListObjects(1).Range.AutoFilter Field:=6
ActiveSheet.ListObjects(1).Range.AutoFilter Field:=10
ActiveSheet.ListObjects(1).Range.AutoFilter Field:=13

'Clear listbox selection after running query
    For i = 0 To List42.ListCount - 1
    List42.Selected(i) = False
    Next i

    For i = 0 To List82.ListCount - 1
    List82.Selected(i) = False
    Next i
    
    For i = 0 To List13.ListCount - 1
    List13.Selected(i) = False
    Next i
    
    For i = 0 To List91.ListCount - 1
    List91.Selected(i) = False
    Next i
    
ActiveSheet.Range("A1").Select

End Sub

Private Sub btn_Run_Report_Click()
On Error GoTo Err_btn_Run_Report_Click
    Dim mws As Worksheet
    Dim ws As Worksheet
    Dim copyRange As Range
    Dim i As Integer
    Dim p As Integer
    Dim strWhere_ME As Variant
    Dim strIN_ME As String
    Dim strWhere_LE As Variant
    Dim strIN_LE As String
    Dim strWhere_AT As Variant
    Dim strIN_AT As String
    Dim strWhere_AN As Variant
    Dim strIN_AN As String
    Dim varItem As Variant
    
    'Set master worksheet
    Set mws = Worksheets("ME_LE_Report")
    
    flgSelectAll = 0
    flgAll_ME = 0
    flgAll_LE = 0
    flgAll_AT = 0
    flgAll_AN = 0
    
For i = 0 To List42.ListCount - 1
        If List42.Selected(i) And List42.Column(0, i) = "All" Then
                flgAll_ME = flgAll_ME + 1
                For p = 1 To List42.ListCount - 1
                strIN_ME = strIN_ME & Left(List42.Column(0, p), 6) & ","
                Next p
                Exit For
        ElseIf List42.Selected(i) Then strIN_ME = strIN_ME & Left(List42.Column(0, i), 6) & ","
        End If
    Next i
    
  For i = 0 To List82.ListCount - 1
        If List82.Selected(i) And List82.Column(0, i) = "All" Then
                flgAll_LE = flgAll_LE + 1
                For p = 1 To List82.ListCount - 1
                strIN_LE = strIN_LE & Left(List82.Column(0, p), 3) & ","
                Next p
                Exit For
        ElseIf List82.Selected(i) Then strIN_LE = strIN_LE & Left(List82.Column(0, i), 3) & ","
        End If
    Next i
    
   For i = 0 To List13.ListCount - 1
        If List13.Selected(i) And List13.Column(0, i) = "All" Then
                flgAll_AT = flgAll_AT + 1
                For p = 1 To List13.ListCount - 1
                strIN_AT = strIN_AT & List13.Column(0, p) & ","
                Next p
                Exit For
        ElseIf List13.Selected(i) Then strIN_AT = strIN_AT & List13.Column(0, i) & ","
        End If
    Next i
    
  For i = 0 To List91.ListCount - 1
        If List91.Selected(i) And List91.Column(0, i) = "All" Then
                flgAll_AN = flgAll_AN + 1
                For p = 1 To List91.ListCount - 1
                strIN_AN = strIN_AN & Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) & ","
                Next p
                Exit For
        ElseIf List91.Selected(i) Then strIN_AN = strIN_AN & Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) & ","
        End If
    Next i
    
    flgSelectAll = flgAll_ME + flgAll_LE + flgAll_AT + flgAll_AN
    If flgSelectAll = 4 Then
     mws.Activate
     UserForm1.Hide
     GoTo Exit_btn_Run_Report_Click
    End If
    
    
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_ME = Left(strIN_ME, Len(strIN_ME) - 1)
               
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_LE = Left(strIN_LE, Len(strIN_LE) - 1)
               
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AT = Left(strIN_AT, Len(strIN_AT) - 1)
               
    'Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AN = Left(strIN_AN, Len(strIN_AN) - 1)
    
    With mws
    mws.Activate
    .AutoFilterMode = False
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=3, Criteria1:=Split(strWhere_ME, ","), Operator:=xlFilterValues
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=6, Criteria1:=Split(strWhere_AT, ","), Operator:=xlFilterValues
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=10, Criteria1:=Split(strWhere_AN, ","), Operator:=xlFilterValues
    ActiveSheet.ListObjects(1).Range. _
        AutoFilter Field:=13, Criteria1:=Split(strWhere_LE, ","), Operator:=xlFilterValues
'    With .AutoFilter.Range
'        On Error Resume Next ' if none selected
'        .Offset(1).Resize(.Rows.Count - 1, Columns.Count).SpecialCells(xlCellTypeVisible).Select
'        On Error GoTo 0
'    End With
    .AutoFilterMode = False
    End With
    
    'Create new sheet
    Sheets.Add After:=Sheets("ME_LE_Report")
    
    'Copy and Paste filtered data to new sheet
    Set ws = Sheets(mws.Index + 1)
    mws.Activate
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set copyRange = Selection
    copyRange.Copy ws.Range("A1")
    ws.Columns("A:R").AutoFit
    
    'Clear filters from report
    mws.ListObjects(1).Range.AutoFilter Field:=3
    mws.ListObjects(1).Range.AutoFilter Field:=6
    mws.ListObjects(1).Range.AutoFilter Field:=10
    mws.ListObjects(1).Range.AutoFilter Field:=13
    
    mws.Range("A1").Select
    
    ws.Activate
    ws.Range("A1").Select
    
    UserForm1.Hide
    
Exit_btn_Run_Report_Click:
    Exit Sub

Err_btn_Run_Report_Click:

    If Err.Number = 5 Then
        MsgBox "You must make a selection(s) from each list" _
               , , "Selection Required !"
        Resume Exit_btn_Run_Report_Click
    Else
        'Write out the error and exit the sub
        MsgBox Err.Description
        Resume Exit_btn_Run_Report_Click
    End If
    
End Sub

Private Sub UserForm_Click()

End Sub
