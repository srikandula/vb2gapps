Attribute VB_Name( = "UserForm1")
Attribute VB_Base( = "0{0ABC0F62-C9A8-4F2C-9E0A-087937E32312}{9BF30DEA-3FD6-4E3B-B093-C963D3D3DEFD}")
Attribute VB_GlobalNameSpace = false
Attribute VB_Creatable = false
Attribute VB_PredeclaredId = true
Attribute VB_Exposed = false
Attribute VB_TemplateDerived = false
Attribute VB_Customizable = false

 function CommandButton1_Click(){

 }

 function Management_Entity_Click(){

 }

 function Run_Button_Click(){

 }

 function btn_Del_Rpts_Start_Over_Click(){

Application.DisplayAlerts = false
var ws  Worksheet
for ( var ws in ThisWorkbook.Sheets) {
if( ws.Name Like "*Sheet*" ) { ws.Delete
} ws
Application.DisplayAlerts = true
Worksheets("ME_LE_Report"().Activate)
ActiveSheet.Range(("A1").Select)

//Clear listbox selection after running query
    for( i = 0 To List42.ListCount - 1) {
    List42.Selected(i) = false
    } i

    for( i = 0 To List82.ListCount - 1) {
    List82.Selected(i) = false
    } i
    
    for( i = 0 To List13.ListCount - 1) {
    List13.Selected(i) = false
    } i
    
    for( i = 0 To List91.ListCount - 1) {
    List91.Selected(i) = false
    } i

 }

 function btn_Report_Pivot_Click(){
On Error( GoTo Err_btn_Report_Pivot_Click)
    var mws  Worksheet
    var ws  Worksheet
    var copyRange  Range
    var i  Integer
    var p  Integer
    var strWhere_ME  Variant
    var strIN_ME  
    var strWhere_LE  Variant
    var strIN_LE  
    var strWhere_AT  Variant
    var strIN_AT  
    var strWhere_AN  Variant
    var strIN_AN  
    var varItem  Variant
    var data_sht  Worksheet
    var pivot_sht  Worksheet
    var startpoint  Range
    var datarange  Variant
    var newrange  Range
    
    // master worksheet
     mws = Worksheets("ME_LE_Report")
    
    flgSelectAll( = 0)
    flgAll_ME( = 0)
    flgAll_LE( = 0)
    flgAll_AT( = 0)
    flgAll_AN( = 0)
    
    for( i = 0 To List42.ListCount - 1) {
        if( List42.Selected(i) And List42.Column(0, i) == "All" ) {
                flgAll_ME( = flgAll_ME + 1)
                for( p = 1 To List42.ListCount - 1) {
                strIN_ME = strIN_ME + Left(List42.Column(0, p), 6) + ","
                } p
                 break
        else if( List42.Selected(i) ) { strIN_ME = strIN_ME + Left(List42.Column(0, i), 6) + ","
         }
    } i
    
  for( i = 0 To List82.ListCount - 1) {
        if( List82.Selected(i) And List82.Column(0, i) == "All" ) {
                flgAll_LE( = flgAll_LE + 1)
                for( p = 1 To List82.ListCount - 1) {
                strIN_LE = strIN_LE + Left(List82.Column(0, p), 3) + ","
                } p
                 break
        else if( List82.Selected(i) ) { strIN_LE = strIN_LE + Left(List82.Column(0, i), 3) + ","
         }
    } i
    
   for( i = 0 To List13.ListCount - 1) {
        if( List13.Selected(i) And List13.Column(0, i) == "All" ) {
                flgAll_AT( = flgAll_AT + 1)
                for( p = 1 To List13.ListCount - 1) {
                strIN_AT = strIN_AT + List13.Column(0, p) + ","
                } p
                 break
        else if( List13.Selected(i) ) { strIN_AT = strIN_AT + List13.Column(0, i) + ","
         }
    } i
    
  for( i = 0 To List91.ListCount - 1) {
        if( List91.Selected(i) And List91.Column(0, i) == "All" ) {
                flgAll_AN( = flgAll_AN + 1)
                for( p = 1 To List91.ListCount - 1) {
                strIN_AN = strIN_AN + Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) + ","
                } p
                 break
        else if( List91.Selected(i) ) { strIN_AN = strIN_AN + Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) + ","
         }
    } i
    
    flgSelectAll( = flgAll_ME + flgAll_LE + flgAll_AT + flgAll_AN)
    if( flgSelectAll == 4 ) {
     GoTo( Err_Select_All_Click)
     }
    
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_ME( = Left(strIN_ME, Len(strIN_ME) - 1))
               
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_LE( = Left(strIN_LE, Len(strIN_LE) - 1))
               
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AT( = Left(strIN_AT, Len(strIN_AT) - 1))
               
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AN( = Left(strIN_AN, Len(strIN_AN) - 1))
    
    with( mws) {
    mws(.Activate)
    .AutoFilterMode = false
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=3, Criteria1=Split(strWhere_ME, ","), Operator=xlFilterValues)
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=6, Criteria1=Split(strWhere_AT, ","), Operator=xlFilterValues)
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=10, Criteria1=Split(strWhere_AN, ","), Operator=xlFilterValues)
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=13, Criteria1=Split(strWhere_LE, ","), Operator=xlFilterValues)
//    with( .AutoFilter.Range) {
//        On Error Resume } // if none selected
//        .Offset(1).Resize(.Rows.Count - 1, Columns.Count).SpecialCells(xlCellTypeVisible).Select
//        On Error GoTo 0
//     }
    .AutoFilterMode = false
     }
    
    //Create new sheet
    Sheets(.Add After=Sheets("ME_LE_Report"))
    
    //Copy and Paste filtered data to new sheet
     ws = Sheets(mws.Index + 1)
    mws(.Activate)
    Range(Selection, Selection.(xlToRight)).Select
    Range(Selection, Selection.(xlDown)).Select
     copyRange = Selection
    copyRange(.Copy ws.Range("A1"))
    ws(.Columns("AR").AutoFit)
    
    //Clear filters from report
    mws(.ListObjects(1).Range.AutoFilter Field=3)
    mws(.ListObjects(1).Range.AutoFilter Field=6)
    mws(.ListObjects(1).Range.AutoFilter Field=10)
    mws(.ListObjects(1).Range.AutoFilter Field=13)

    //Retrieve data range
    ws(.Activate)
    ws(.Range("A1").Select)
    Range(Selection, Selection.(xlDown)).Select
    Range(Selection, Selection.(xlToRight)).Select
     datarange = Selection

    //Create Pivot Table
    Sheets(.Add After=ws)
     pivot_sht = Sheets(ws.Index + 1)
    Application(.WindowState = xlMaximized)
    ActiveWorkbook(.PivotCaches.Create(SourceType=xlDatabase, SourceData=datarange, _)
    Version(=xlPivotTableVersion14).CreatePivotTable _)
        TableDestination(=pivot_sht.Range("A3"), TableName="Report_Pivot", DefaultVersion _)
        (=xlPivotTableVersion14)
        
    pivot_sht(.Select)
    Cells((3, 1).Select)
    ActiveWorkbook.ShowPivotTableFieldList = true
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("LEGAL_ENTITY")) {
        .Orientation( = xlColumnField)
        .Position( = 1)
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("AT_Sort")) {
        .Orientation( = xlRowField)
        .Position( = 1)
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE")) {
        .Orientation( = xlRowField)
        .Position( = 2)
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("Acct_Sort")) {
        .Orientation( = xlRowField)
        .Position( = 3)
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC")) {
        .Orientation( = xlRowField)
        .Position( = 4)
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("ACCOUNT")) {
        .Orientation( = xlRowField)
        .Position( = 5)
     }
    ActiveSheet(.PivotTables("Report_Pivot").AddDataField ActiveSheet.PivotTables( _)
        "Report_Pivot"().PivotFields("MARS_AMOUNT_IN"), "Sum of MARS_AMOUNT_IN", xlSum)
    ActiveSheet(.PivotTables("Report_Pivot").AddDataField ActiveSheet.PivotTables( _)
        "Report_Pivot"().PivotFields("MARS_AMOUNT_IN"), "Sum of MARS_AMOUNT_IN2", xlSum)
        ActiveWorkbook.ShowPivotTableFieldList = true
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN")) {
        .NumberFormat( = "$#,##0.00")
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN2")) {
        .Calculation( = xlPercentOfRow)
        .NumberFormat( = "0.00%")
     }
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN"). _)
        Caption( = "Sum of LE In")
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("Sum of MARS_AMOUNT_IN2"). _)
        Caption( = "% of LE In")
    Application(.WindowState = xlMaximized)
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("AT_Sort").Subtotals = _)
        Array(false, false, false, false, false, false, false, false, false, false, false, false)
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("AT_Sort").LayoutForm _)
        = xlTabular()
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE").Subtotals = _)
        Array(false, false, false, false, false, false, false, false, false, false, false, false)
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE").LayoutForm _)
        = xlTabular()
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("Acct_Sort").Subtotals = _)
        Array(false, false, false, false, false, false, false, false, false, false, false, false)
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("Acct_Sort").LayoutForm _)
        = xlTabular()
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC"). _)
        Subtotals = Array(false, false, false, false, false, false, false, false, false, false, _
        false, false)
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC"). _)
        LayoutForm( = xlTabular)
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("ACCOUNT").Subtotals = Array _)
        (false, false, false, false, false, false, false, false, false, false, false, false)
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("ACCOUNT").LayoutForm = _)
        xlTabular()
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("ACCOUNT_TYPE"). _)
        Caption( = "Account Type")
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("FINANCIALS_DESC"). _)
        Caption( = "Financials Description")
    ActiveSheet(.PivotTables("Report_Pivot").PivotFields("ACCOUNT"). _)
        Caption( = "Account")
    ActiveWorkbook.ShowPivotTableFieldList = false
    ActiveSheet(.PivotTables("Report_Pivot").TableStyle2 = "")
    //Column headings bold
    Range(("A7").Select)
    Range(Selection, Selection.(xlToRight)).Select
    Selection.Font.Bold = true
    Selection.(xlToRight).Select
    //Create filters
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("ME")) {
       .Orientation( = xlPageField)
       .Position( = 1)
       .Caption( = "Management Entity")
       .EnableMultiplePageItems = true
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("PERIOD_NAME")) {
        .Orientation( = xlPageField)
        .Position( = 1)
        .Caption( = "Period")
        .EnableMultiplePageItems = true
     }
    with( ActiveSheet.PivotTables("Report_Pivot").PivotFields("ME_COUNTRY_PER_MARS")) {
        .Orientation( = xlPageField)
        .Position( = 1)
        .Caption( = "Country")
        .EnableMultiplePageItems = true
     }
    //Format column headings
    Range(("XFD7").Select)
    Selection.(xlToLeft).Select
    Range(("B7", Selection).Select)
    Selection.Font.Bold = true
    with( Selection.Interior) {
        .Pattern( = xlSolid)
        .PatternColorIndex( = xlAutomatic)
        .ThemeColor( = xlThemeColorAccent1)
        .TintAndShade( = 0.599993896298105)
        .PatternTintAndShade( = 0)
     }
    //Format column headings
    Range(("XFD6").Select)
    Selection.(xlToLeft).Select
    Range(Selection, Selection.(xlToLeft)).Select
    Selection.Font.Bold = true
    with( Selection.Interior) {
        .Pattern( = xlSolid)
        .PatternColorIndex( = xlAutomatic)
        .ThemeColor( = xlThemeColorAccent1)
        .TintAndShade( = 0.599993896298105)
        .PatternTintAndShade( = 0)
     }
    //Format grand total
    Range(("A10000").Select)
    Selection.(xlUp).Select
    Selection.Font.Bold = true
    //Replace underscores in column headings
    Range(("E7").Select)
    Range(Selection, Selection.(xlToRight)).Select
    Selection(.Replace What="_", Replacement=" ", LookAt=xlPart, _)
        SearchOrder=xlByRows, MatchCase=false, SearchFormat=false, _
        ReplaceFormat=false
    //Don//t show gridlines
    ActiveWindow.DisplayGridlines = false
    //Clean up error values
    with( ActiveSheet.PivotTables("Report_Pivot")) {
        .DisplayErrorString = true
        .ErrorString( = "Sums to Zero")
     }
    //Clean up visuals
    ActiveWindow(.Zoom = 90)
    ActiveSheet(.Columns.AutoFit)
    Cells(.Select)
    with( Selection) {
        .HorizontalAlignment( = xlLeft)
        .VerticalAlignment( = xlBottom)
     }
    //Hide sort columns
    Columns("A").Hidden = true
    Columns("C").Hidden = true
    //Assign and format filter labels
    Range(("D1").Select)
    ActiveCell(.FormulaR1C1 = "=RC[-3]")
    Range(("D1D3").Select)
    Selection(.FillDown)
    Selection(.Borders(xlDiagonalDown).LineStyle = xlNone)
    Selection(.Borders(xlDiagonalUp).LineStyle = xlNone)
    with( Selection.Borders(xlEdgeLeft)) {
        .LineStyle( = xlContinuous)
        .ColorIndex( = 0)
        .TintAndShade( = 0)
        .Weight( = xlThin)
     }
    with( Selection.Borders(xlEdgeTop)) {
        .LineStyle( = xlContinuous)
        .ColorIndex( = 0)
        .TintAndShade( = 0)
        .Weight( = xlThin)
     }
    with( Selection.Borders(xlEdgeBottom)) {
        .LineStyle( = xlContinuous)
        .ColorIndex( = 0)
        .TintAndShade( = 0)
        .Weight( = xlThin)
     }
    with( Selection.Borders(xlEdgeRight)) {
        .LineStyle( = xlContinuous)
        .ColorIndex( = 0)
        .TintAndShade( = 0)
        .Weight( = xlThin)
     }
    with( Selection.Borders(xlInsideVertical)) {
        .LineStyle( = xlContinuous)
        .ColorIndex( = 0)
        .TintAndShade( = 0)
        .Weight( = xlThin)
     }
    with( Selection.Borders(xlInsideHorizontal)) {
        .LineStyle( = xlContinuous)
        .ColorIndex( = 0)
        .TintAndShade( = 0)
        .Weight( = xlThin)
     }
    //Hide form when complete
    UserForm1(.Hide)
    Range(("A1").Select)
    
    
Exit_btn_Report_Pivot_Click()
     break}

Err_Select_All_Click()
MsgBox "You( may not select ALL in every list. The full report is in tab ""ME_LE_Report.""" _)
               , , "Select( a value from each box.")
        GoTo( Exit_btn_Report_Pivot_Click)

Err_btn_Report_Pivot_Click()

    if( Err.Number == 5 ) {
        MsgBox( "You must make a selection(s) from each list" _)
               , , "Selection( Required !")
        Resume( Exit_btn_Report_Pivot_Click)
    else if( Err.Number = 1004 ) {
        MsgBox "Pleaes re-select as cetain selection(s) don//t exist" _
               , , "Re-Selection( Required !")
        Resume( Exit_btn_Report_Pivot_Click)
       //Write out the error and exit the sub
    }else{ MsgBox Err.Description
        Resume( Exit_btn_Report_Pivot_Click)
     }

 }

 function btn_Reset_Form_Click(){

Worksheets("ME_LE_Report"().Activate)

//Clear filters from report
ActiveSheet.ListObjects((1).Range.AutoFilter Field=3)
ActiveSheet.ListObjects((1).Range.AutoFilter Field=6)
ActiveSheet.ListObjects((1).Range.AutoFilter Field=10)
ActiveSheet.ListObjects((1).Range.AutoFilter Field=13)

//Clear listbox selection after running query
    for( i = 0 To List42.ListCount - 1) {
    List42.Selected(i) = false
    } i

    for( i = 0 To List82.ListCount - 1) {
    List82.Selected(i) = false
    } i
    
    for( i = 0 To List13.ListCount - 1) {
    List13.Selected(i) = false
    } i
    
    for( i = 0 To List91.ListCount - 1) {
    List91.Selected(i) = false
    } i
    
ActiveSheet.Range(("A1").Select)

 }

 function btn_Run_Report_Click(){
On Error( GoTo Err_btn_Run_Report_Click)
    var mws  Worksheet
    var ws  Worksheet
    var copyRange  Range
    var i  Integer
    var p  Integer
    var strWhere_ME  Variant
    var strIN_ME  
    var strWhere_LE  Variant
    var strIN_LE  
    var strWhere_AT  Variant
    var strIN_AT  
    var strWhere_AN  Variant
    var strIN_AN  
    var varItem  Variant
    
    // master worksheet
     mws = Worksheets("ME_LE_Report")
    
    flgSelectAll( = 0)
    flgAll_ME( = 0)
    flgAll_LE( = 0)
    flgAll_AT( = 0)
    flgAll_AN( = 0)
    
for( i = 0 To List42.ListCount - 1) {
        if( List42.Selected(i) And List42.Column(0, i) == "All" ) {
                flgAll_ME( = flgAll_ME + 1)
                for( p = 1 To List42.ListCount - 1) {
                strIN_ME = strIN_ME + Left(List42.Column(0, p), 6) + ","
                } p
                 break
        else if( List42.Selected(i) ) { strIN_ME = strIN_ME + Left(List42.Column(0, i), 6) + ","
         }
    } i
    
  for( i = 0 To List82.ListCount - 1) {
        if( List82.Selected(i) And List82.Column(0, i) == "All" ) {
                flgAll_LE( = flgAll_LE + 1)
                for( p = 1 To List82.ListCount - 1) {
                strIN_LE = strIN_LE + Left(List82.Column(0, p), 3) + ","
                } p
                 break
        else if( List82.Selected(i) ) { strIN_LE = strIN_LE + Left(List82.Column(0, i), 3) + ","
         }
    } i
    
   for( i = 0 To List13.ListCount - 1) {
        if( List13.Selected(i) And List13.Column(0, i) == "All" ) {
                flgAll_AT( = flgAll_AT + 1)
                for( p = 1 To List13.ListCount - 1) {
                strIN_AT = strIN_AT + List13.Column(0, p) + ","
                } p
                 break
        else if( List13.Selected(i) ) { strIN_AT = strIN_AT + List13.Column(0, i) + ","
         }
    } i
    
  for( i = 0 To List91.ListCount - 1) {
        if( List91.Selected(i) And List91.Column(0, i) == "All" ) {
                flgAll_AN( = flgAll_AN + 1)
                for( p = 1 To List91.ListCount - 1) {
                strIN_AN = strIN_AN + Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) + ","
                } p
                 break
        else if( List91.Selected(i) ) { strIN_AN = strIN_AN + Left(List91.Column(0, p), InStr(List91.Column(0, p), "-") - 2) + ","
         }
    } i
    
    flgSelectAll( = flgAll_ME + flgAll_LE + flgAll_AT + flgAll_AN)
    if( flgSelectAll == 4 ) {
     mws(.Activate)
     UserForm1(.Hide)
     GoTo( Exit_btn_Run_Report_Click)
     }
    
    
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_ME( = Left(strIN_ME, Len(strIN_ME) - 1))
               
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_LE( = Left(strIN_LE, Len(strIN_LE) - 1))
               
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AT( = Left(strIN_AT, Len(strIN_AT) - 1))
               
    //Create the WHERE string, and strip off the last comma of the IN string
    strWhere_AN( = Left(strIN_AN, Len(strIN_AN) - 1))
    
    with( mws) {
    mws(.Activate)
    .AutoFilterMode = false
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=3, Criteria1=Split(strWhere_ME, ","), Operator=xlFilterValues)
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=6, Criteria1=Split(strWhere_AT, ","), Operator=xlFilterValues)
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=10, Criteria1=Split(strWhere_AN, ","), Operator=xlFilterValues)
    ActiveSheet(.ListObjects(1).Range. _)
        AutoFilter( Field=13, Criteria1=Split(strWhere_LE, ","), Operator=xlFilterValues)
//    with( .AutoFilter.Range) {
//        On Error Resume } // if none selected
//        .Offset(1).Resize(.Rows.Count - 1, Columns.Count).SpecialCells(xlCellTypeVisible).Select
//        On Error GoTo 0
//     }
    .AutoFilterMode = false
     }
    
    //Create new sheet
    Sheets(.Add After=Sheets("ME_LE_Report"))
    
    //Copy and Paste filtered data to new sheet
     ws = Sheets(mws.Index + 1)
    mws(.Activate)
    Range(Selection, Selection.(xlToRight)).Select
    Range(Selection, Selection.(xlDown)).Select
     copyRange = Selection
    copyRange(.Copy ws.Range("A1"))
    ws(.Columns("AR").AutoFit)
    
    //Clear filters from report
    mws(.ListObjects(1).Range.AutoFilter Field=3)
    mws(.ListObjects(1).Range.AutoFilter Field=6)
    mws(.ListObjects(1).Range.AutoFilter Field=10)
    mws(.ListObjects(1).Range.AutoFilter Field=13)
    
    mws(.Range("A1").Select)
    
    ws(.Activate)
    ws(.Range("A1").Select)
    
    UserForm1(.Hide)
    
Exit_btn_Run_Report_Click()
     break}

Err_btn_Run_Report_Click()

    if( Err.Number == 5 ) {
        MsgBox( "You must make a selection(s) from each list" _)
               , , "Selection( Required !")
        Resume( Exit_btn_Run_Report_Click)
    }else{
        //Write out the error and exit the sub
        MsgBox( Err.Description)
        Resume( Exit_btn_Run_Report_Click)
     }
    
 }

 function UserForm_Click(){

 }
