Attribute VB_Name( = "OCF_Template")
// TODO  Group constant initialization more sensibly

// Force declaration of variables
Option Explicit()

// Declare global variables and constants
 ACCOUNTING_NUMBER_FORMAT  

 OCF_ACTIVITY_HEADER  
 TRADING_PARTNER_HEADER  
 TE_BNS_PC_HEADER  

// Declare constants for field names for LOOKUP field insertions
 HFM_ACCOUNT_HEADER  
 GL_ICP_HEADER  
 GL_PC_HEADER  

// TODO  Rename GL_Other to GL_Account in pivot?
 GL_ACCOUNT_HEADER  
 BALANCE_HEADER  
 GL_ACCOUNT_DESCRIPTION_HEADER  
 TOTAL_ACCOUNT_BALANCE_HEADER  
 TE_BALANCE_HEADER  
 BNS_BALANCE_HEADER  
 CONTRA_BNS_PC_HEADER  
 IS_CLEARING_ACCOUNT_HEADER  
 HAS_CLEARING_ACCOUNT_HEADER  
 CLEARING_ACCOUNT_BALANCE_HEADER  
// ADJUSTMENT_TO_BNS_HEADER  
 BNS_YTD_BALANCE_HEADER  
 BNS_AMOUNT_TO_CLEAR_HEADER  

 PRE_CLOSE_AMF_PIVOT_SHEET_NAME  
 PRE_CLOSE_AMF_PIVOT_TABLE_NAME  
// Break out pre-close BNS balance for TE- / CS-owned
 BNS_PRE_CLOSE_INCOME_TE_HEADER  
 BNS_POST_CLOSE_INCOME_CS_HEADER  
 BNS_PRE_CLOSE_INCOME_TABLE_NAME  
 PRE_CLOSE_BNS_INCOME_SHEET_NAME  
 PRE_CLOSE_BNS_INCOME_TABLE_NAME  
 AMF_PIVOT_TABLE_NAME  
 ALL_ACCOUNTS_PIVOT_SHEET_NAME  
 ALL_ACCOUNTS_PIVOT_TABLE_NAME  
 ALL_ACCOUNTS_SHEET_NAME  
 AMF_PIVOT_CACHE_INDEX  Integer
 MODDED_AMF_PIVOT_ACCOUNTS_SHEET_NAME  
 MODDED_AMF_PIVOT_ACCOUNTS_TABLE_NAME  
 CLEARING_ENTRY_HEADERS(1 To 9)  


// Table (ListOjbect) names
 OCF_ACTIVITY_TEMPLATE_TABLE_NAME  
 CLEARING_JE_TABLE_NAME  
 TABLE_STYLE_FORMAT  

 CONTRA_BNS_PROFIT_CENTER  

 CLEARING_ENTRY_TEXT  
 GL_DEBIT_CODE  
 GL_CREDIT_CODE  

 MODIFIED_HEADER_BACKGROUND_COLOR  
 OCF_ACTIVITY_TEMPLATE_TAB_COLOR  
 JE_VOUCHER_TAB_COLOR  

// AMF_RAW_DATA_SHEET_NAME is set in ImportExternalAMF()
 AMF_RAW_DATA_SHEET_NAME  
 MODIFIED_AMF_SHEET_NAME  
 MODIFIED_AMF_TABLE_NAME  
 MOD_PRE_CLOSE_AMF_TABLE_NAME  

 PERSISTENT_STORAGE_SHEET  
 AMF_PIVOT_SHEET_NAME  
 OCF_ACTIVITY_TEMPLATE_SHEET_NAME  
 JE_VOUCHER_SHEET_NAME  

 AMF_GL_ACCT_DESC_HEADER  

// Define constants for OCF Activity field indicators of "Excluded"
 OCF_ACTIVITY_EXCLUDE1  
// OCF_ACTIVITY_EXCLUDE2  

// Define constant for ICP / trading partner field search
 ICP_FIELD_SEARCH_VALUE  

// Define TE profit center flag
 TE_PROFIT_CENTER_FLAG  

 CONTRA_BNS_PC_PIVOT_SHEET_NAME  
 CONTRA_BNS_PC_BAL_SHEET_NAME  
 CONTRA_BNS_PC_BAL_TABLE_NAME  
 BNS_PROFIT_CENTERS_SHEET_NAME  
 BNS_PROFIT_CENTERS_TABLE_NAME  
 HFM_ACCT_PL_EXCLUDE_RANGE  

//var CLEARING_ENTRY_HEADER_BACKGROUND_RGB = "79,129,129"
//var WHITE_RGB = "255,255,255"
 ENTRY_FIELD_BKG_COLOR  
 CALC_MAPPING_TABLE_TAB_COLOR  

// Sheet names not to delete with RESET button
 DO_NOT_DELETE_SHEETS(1 To 9)  

 ZBV_OCFAT_TABLE_NAME  
 ZBV_POST_OCFAT_SHEET_NAME  
 WORKSHEET_PASSWORD  
 HIDDEN_WORKSHEETS(1 To 3)  



function InitializeTemplateGlobals(){
    // 5/27/2015 -- revised format for regional / localization issues
    ACCOUNTING_NUMBER_FORMAT( = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)")
    OCF_ACTIVITY_HEADER( = "OCF_Activity")
    TRADING_PARTNER_HEADER( = "Trading_Partner")
    TE_BNS_PC_HEADER( = "TE / BNS PC")
    
    HFM_ACCOUNT_HEADER( = "HFM_Account")
    GL_ICP_HEADER( = "GL_ICP")
    GL_PC_HEADER( = "GL_PC")
    GL_ACCOUNT_HEADER( = "GL_Other")
    BALANCE_HEADER( = "Amount")
    GL_ACCOUNT_DESCRIPTION_HEADER( = "GL_Account_Description")

    TOTAL_ACCOUNT_BALANCE_HEADER = "TE + BNS Balance"
    TE_BALANCE_HEADER( = "TE")
    BNS_BALANCE_HEADER( = "BNS")
    CONTRA_BNS_PC_HEADER( = "Contra-BNS PC")
    IS_CLEARING_ACCOUNT_HEADER( = "Is Clearing Acct?")
    HAS_CLEARING_ACCOUNT_HEADER( = "Has Clearing Acct?")
    CLEARING_ACCOUNT_BALANCE_HEADER( = "Clearing Acct Balance")
    //TODO  refactor ADJUSTMENT_TO_BNS_HEADER and "Adjustment to BNS Balance"
    //ADJUSTMENT_TO_BNS_HEADER = "Adjustment to BNS Balance"
    BNS_YTD_BALANCE_HEADER( = "Total BNS Balance")
    BNS_AMOUNT_TO_CLEAR_HEADER( = "BNS Amount to Clear")
    
    PRE_CLOSE_AMF_PIVOT_SHEET_NAME( = "Pre-Close AMF Pivot")
    PRE_CLOSE_AMF_PIVOT_TABLE_NAME( = "Pre_Close_AMF_Pivot")
    BNS_PRE_CLOSE_INCOME_TE_HEADER( = "BNS PRE-Close Income - TE")
    BNS_POST_CLOSE_INCOME_CS_HEADER( = "CS-Owned BNS Balance")
    BNS_PRE_CLOSE_INCOME_TABLE_NAME( = "Pre_Close_Income_Table")
    PRE_CLOSE_BNS_INCOME_SHEET_NAME( = "Pre-Close BNS Income")
    PRE_CLOSE_BNS_INCOME_TABLE_NAME( = "Pre_Close_BNS_Income")
    AMF_PIVOT_TABLE_NAME( = "AMF_Pivot_Table")
    ALL_ACCOUNTS_PIVOT_SHEET_NAME( = "All_Accounts_Pivot")
    ALL_ACCOUNTS_PIVOT_TABLE_NAME( = "All_Accounts_Pivot")
    AMF_PIVOT_CACHE_INDEX( = 2)
    ALL_ACCOUNTS_SHEET_NAME( = "All_Accounts")
    MODDED_AMF_PIVOT_ACCOUNTS_SHEET_NAME( = "Modded_AMF_Pivot_Accounts")
    MODDED_AMF_PIVOT_ACCOUNTS_TABLE_NAME( = "Modded_AMF_Pivot_Accounts")
    
    // 7/28/2015 -- moved this array here from ModifyActivityTemplateData
    CLEARING_ENTRY_HEADERS((1) = "Doc Date")
    CLEARING_ENTRY_HEADERS((2) = "Posting Date")
    CLEARING_ENTRY_HEADERS((3) = "PK")
    CLEARING_ENTRY_HEADERS((4) = "Account")
    CLEARING_ENTRY_HEADERS((5) = "PC")
    CLEARING_ENTRY_HEADERS((6) = "CC")
    CLEARING_ENTRY_HEADERS((7) = "Trad Ptr")
    CLEARING_ENTRY_HEADERS((8) = "Amount")
    CLEARING_ENTRY_HEADERS((9) = "Text")
    
    OCF_ACTIVITY_TEMPLATE_TABLE_NAME( = "OCF_Activity_Template_Table")
    CLEARING_JE_TABLE_NAME( = "Clearing_JE_Table")
    TABLE_STYLE_FORMAT( = "TableStyleLight9")
    
    CONTRA_BNS_PROFIT_CENTER( = ActiveWorkbook.Names("input_Contra_BNS_PC").RefersToRange.Value)
    CLEARING_ENTRY_TEXT( = """Clear BNS activity from """)
    
    Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
        Case( "SAP PR2")
            GL_DEBIT_CODE( = "40")
            GL_CREDIT_CODE( = "50")
        Case( "C1")
            GL_DEBIT_CODE( = "DR")
            GL_CREDIT_CODE( = "CR")
        Case( "AMPICS")
            GL_DEBIT_CODE( = "DR")
            GL_CREDIT_CODE( = "CR")
     Select
    
    MODIFIED_HEADER_BACKGROUND_COLOR = 6 //Yellow
    //MODIFIED_HEADER_BACKGROUND_COLOR = vbYellow
    //OCF_ACTIVITY_TEMPLATE_TAB_COLOR = 3 //Red
    OCF_ACTIVITY_TEMPLATE_TAB_COLOR( = vbRed)
    //OCF_ACTIVITY_TEMPLATE_TAB_COLOR = vbRed
    JE_VOUCHER_TAB_COLOR = 10 //Green
    CALC_MAPPING_TABLE_TAB_COLOR( = vbMagenta)
    
    MODIFIED_AMF_SHEET_NAME( = "AMF Table - Modded")
    PERSISTENT_STORAGE_SHEET( = "PersistentStorage")
    AMF_PIVOT_SHEET_NAME( = "Modded AMF Pivot")
    OCF_ACTIVITY_TEMPLATE_SHEET_NAME( = "OCF Activity Template")
    JE_VOUCHER_SHEET_NAME( = "Journal Entries")
    MODIFIED_AMF_TABLE_NAME( = "Modded_AMF_Table")
    MOD_PRE_CLOSE_AMF_TABLE_NAME( = "Pre_Close_AMF_Table")
    
    
    AMF_GL_ACCT_DESC_HEADER( = "GL_AcctDesc")
    
    // 7/30/2015 -- revised for single "Exclude" classifier
    OCF_ACTIVITY_EXCLUDE1( = "Exclude")
    //OCF_ACTIVITY_EXCLUDE2 = "Exclude - N/A"
    
    ICP_FIELD_SEARCH_VALUE( = "NoICP")
    
    TE_PROFIT_CENTER_FLAG( = "TE")
    
    CONTRA_BNS_PC_PIVOT_SHEET_NAME( = "Contra-BNS PC Pivot")
    CONTRA_BNS_PC_BAL_SHEET_NAME( = "Contra-BNS PC")
    CONTRA_BNS_PC_BAL_TABLE_NAME( = "Contra_BNS_PC")
    BNS_PROFIT_CENTERS_SHEET_NAME( = "BNS Profit Centers")
    BNS_PROFIT_CENTERS_TABLE_NAME( = "BNS_Profit_Centers")
    
    HFM_ACCT_PL_EXCLUDE_RANGE( = "123")
    
    ENTRY_FIELD_BKG_COLOR( = RGB(197, 217, 241))
    
    // Populate DO_NOT_DELETE_SHEETS array
    DO_NOT_DELETE_SHEETS((1) = "TODO")
    DO_NOT_DELETE_SHEETS((2) = "ADMIN")
    DO_NOT_DELETE_SHEETS(3) = "Input + Assumptions"
    DO_NOT_DELETE_SHEETS((4) = "Data Dictionary")
    DO_NOT_DELETE_SHEETS((5) = "HFM Acct - OCF Activity")
    DO_NOT_DELETE_SHEETS((6) = BNS_PROFIT_CENTERS_SHEET_NAME)
    DO_NOT_DELETE_SHEETS((7) = "GL Acct - Clearing")
    DO_NOT_DELETE_SHEETS((8) = "Cost Centers")
    DO_NOT_DELETE_SHEETS((9) = PERSISTENT_STORAGE_SHEET)
    
    ZBV_OCFAT_TABLE_NAME( = "ZBV_OCFAT_Table")
    ZBV_POST_OCFAT_SHEET_NAME( = "ZBV")
    
    WORKSHEET_PASSWORD( = "Ni!")
    
    HIDDEN_WORKSHEETS((1) = "ADMIN")
    HIDDEN_WORKSHEETS((2) = "Data Dictionary")
    HIDDEN_WORKSHEETS((3) = "PersistentStorage")
    
 }

function IsInArray(varFind  Variant, varArray  Variant)  Boolean{
    var intI  Integer
    
    for( intI = 1 To UBound(varArray)) {
        if( varFind == varArray(intI) ) {
            IsInArray = true
             function{
         }
    } intI
    
    IsInArray = false
 }

// function prompts user to select a worksheet and returns the name of the selected worksheet{
function SelectWorksheetName(){
    var rngWorksheetCell  Range
    
    // Pop up InputBox which selects a range to identify worksheet to operate on
     rngWorksheetCell = Application.InputBox(prompt="Select the target worksheet by clicking anywhere in it", Type=8)
    
    // Get name of selected worksheet
    SelectWorksheetName( = rngWorksheetCell.Worksheet.Name)
 }

function ModifyAllMappingFile(){
        
    var rngFirstCellOfHeaders, rngHeaders  Range
    var lngLastRow  
    
    // 5/26/2015 -- modified for inclusion of pre-close data
    AMF_RAW_DATA_SHEET_NAME( = ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange.Value)
    
    //  first cell of headers per Greg Stanko//s assurance that this can be relied on from FDM
     rngFirstCellOfHeaders = Worksheets(AMF_RAW_DATA_SHEET_NAME).Range("a5")
    
    Application(.StatusBar = "Modifying All Mapping File preparatory for OCF analysis ...")
    // Turn off screen updating to speed function execution
    Application.ScreenUpdating = false
    
    // Copy worksheet and perform modifications in copied table
    rngFirstCellOfHeaders(.Worksheet.Copy before=rngFirstCellOfHeaders.Worksheet)
    ActiveSheet(.Name = MODIFIED_AMF_SHEET_NAME)
    ActiveSheet(.Tab.Color = vbYellow)
    ActiveSheet(.Unprotect Password=WORKSHEET_PASSWORD)
    
    // Update pointer to first cell of headers on copied sheet
     rngFirstCellOfHeaders = Worksheets(MODIFIED_AMF_SHEET_NAME).Range(rngFirstCellOfHeaders.Address)
    // Identify last row
    lngLastRow = rngFirstCellOfHeaders.(xlDown).Row
    
    // Identify range of table headers
    // This will need to be updated after columns are inserted to keep it current
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    // Tidy-up raw All Mapping File
    // Insert row above headers to enable clean CurrentRegion selection
    rngFirstCellOfHeaders(.EntireRow.Insert)
    // Delete two incongruous rows below header row of AMF
    rngFirstCellOfHeaders(.Offset(1, 0).EntireRow.Delete)
    rngFirstCellOfHeaders(.Offset(1, 0).EntireRow.Delete)
    
    
    // Create data table (ListObject) from copy of All Mapping File
    rngFirstCellOfHeaders(.CurrentRegion.Select)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = MODIFIED_AMF_TABLE_NAME)
    
    // Insert OCF Activity lookup from HFM Account value
    rngHeaders(.Find(what=HFM_ACCOUNT_HEADER, lookat=xlWhole).Offset(0, 1).Activate)
    ActiveCell(.EntireColumn.Insert)
    // Due to insert, ActiveCell now points to cell in newly-inserted column
    with( ActiveCell) {
        .Value( = OCF_ACTIVITY_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font(.Color = vbBlack)
     }
    
    //  formula for lookup of OCF Activity
    ActiveCell(.Offset(1, 0).Formula = "=VLOOKUP([@[HFM_Account]],HFM_Acct_OCF_Activity,3,FALSE)")

    // Insert Trading Partner lookup from GL_ICP
    rngHeaders(.Find(what=GL_ICP_HEADER, lookat=xlWhole).Offset(0, 1).Activate)
    ActiveCell(.EntireColumn.Insert)
    with( ActiveCell) {
        .Value( = TRADING_PARTNER_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font(.Color = vbBlack)
     }
    //  formula for VLOOKUP of Trading Partner
    //ActiveCell.Formula = "=IF(NOT(ISERROR(FIND(""NoICP"",$G7))),"""",$G7)"
    // [C1 TAILORING] 6/9/2015 -- leave this field blank
    Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
        Case( "SAP PR2")
            ActiveCell.Offset(1, 0).Formula = "=IF(NOT(ISERROR(FIND(""" + ICP_FIELD_SEARCH_VALUE + """,[@[GL_ICP]]))),"""",[@[GL_ICP]])"
        Case( "C1")
            ActiveCell(.Offset(1, 0).Value = " ")
            ActiveCell(.Offset(1, 0).AutoFill Destination=Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(lngLastRow - 1, 0)))
        Case( "AMPICS")
            ActiveCell(.Offset(1, 0).Value = " ")
            ActiveCell(.Offset(1, 0).AutoFill Destination=Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(lngLastRow - 1, 0)))
        Case }else{
            MsgBox( prompt="Else condition in Select on ERP hit!")
     Select
    
    // Insert TE / BNS Profit Center lookup from GL_PC
    rngHeaders(.Find(GL_PC_HEADER).Offset(0, 1).Activate)
    ActiveCell(.EntireColumn.Insert)
    with( ActiveCell) {
        .Value( = TE_BNS_PC_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font(.Color = vbBlack)
     }
    //  formula for VLOOKUP of TE / BNS Profit Center
    Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
        Case( "SAP PR2")
            ActiveCell.Offset(1, 0).Formula = "=IFERROR(VLOOKUP(VALUE([@[GL_PC]]),BNS_Profit_Centers,2,FALSE),""" + TE_PROFIT_CENTER_FLAG + """)"
        Case( "C1")
            ActiveCell.Offset(1, 0).Formula = "=IFERROR(VLOOKUP([@[GL_PC]],BNS_Profit_Centers,2,FALSE),""" + TE_PROFIT_CENTER_FLAG + """)"
        Case( "AMPICS")
            ActiveCell.Offset(1, 0).Formula = "=IFERROR(VLOOKUP([@[GL_PC]],BNS_Profit_Centers,2,FALSE),""" + TE_PROFIT_CENTER_FLAG + """)"
     Select
    
    // Identify range of table headers
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    // Store header range persistently
    SetPersistentVariable( strVariable="rngHeaders", strValue=rngHeaders.Address, strDescription="Address of the header row of the modified AMF table.")
    
    // Tidy up
    Worksheets((MODIFIED_AMF_SHEET_NAME).Activate)
    
    //ActivateFilters rngHeaders
    Worksheets((MODIFIED_AMF_SHEET_NAME).ListObjects(MODIFIED_AMF_TABLE_NAME).HeaderRowRange.CurrentRegion.Columns.AutoFit)
    Worksheets((MODIFIED_AMF_SHEET_NAME).ListObjects(MODIFIED_AMF_TABLE_NAME).HeaderRowRange.Item(1).Activate)
    ActiveWindow(.Zoom = 85)
    
    Application.StatusBar = false
    Application.ScreenUpdating = true
 }

// Store persistent variable
function SetPersistentVariable(strVariable  , strValue  ,  strDescription  ){
    Call( InitializeTemplateGlobals)
    
    Worksheets((PERSISTENT_STORAGE_SHEET).Activate)
    if( IsEmpty(Range("a2")) ) {
        Range(("a2").Value = strVariable)
        Range(("b2").Value = strValue)
        Range(("c2").Value = strDescription)
    }else{
        Range("a1").(xlDown).Offset(1, 0).Activate
        ActiveCell(.Value = strVariable)
        ActiveCell(.Offset(0, 1).Value = strValue)
        ActiveCell(.Offset(0, 2).Value = strDescription)
     }
 }

// Retrieve persistent variable
function GetPersistentVariable(strVarName  ){
    var rngVariables  Range
    var lngLastRow, lngMatchRow  
    var strMatchFound  
    
    Call( InitializeTemplateGlobals)
    
    with( Worksheets(PERSISTENT_STORAGE_SHEET)) {
        lngLastRow = .Range("A1").(xlDown).Row
        
        // Get search range
        if( ! IsEmpty(.Range("A2")) ) {
             rngVariables = .Range("A2", "B" + lngLastRow)
        }else{
            //TODO  proper error handling here and when calling function
            MsgBox( ("No variables stored persistently!"))
         }
     }
    
    GetPersistentVariable( = rngVariables.Item(Application.WorksheetFunction.Match(strVarName, rngVariables.Columns(1), 0), 2))
 }

function ActivateFilters(rngFilterRange  Range){
    ActiveSheet.AutoFilterMode = false
    ActiveSheet(.Range(rngFilterRange.Address).AutoFilter)
 }


function PivotAMFData(){
    var rngHeaders, rngPivotRange  Range
    var pvtCache  PivotCache
    var pvtTable  PivotTable
    var pvtField  PivotField
    var pvtPivotItem  PivotItem
    var wsNewSheet  Worksheet
    var strPivotItemName  
    
    // 5/26/2015 -- updated for pre-close data inclusion
    //AMF_RAW_DATA_SHEET_NAME = GetPersistentVariable(strVarName="AMF_RAW_DATA_SHEET_NAME")
    AMF_RAW_DATA_SHEET_NAME( = ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange.Value)
    
    Application(.StatusBar = "Pivoting Modified All Mapping File data ...")
    Application.ScreenUpdating = false
    Application.DisplayStatusBar = true
    
    with( Worksheets(MODIFIED_AMF_SHEET_NAME)) {
        .Activate()
        //6/11/2015 -- testing elimination of GetPersistentVariable call
        // rngHeaders = Range(GetPersistentVariable("rngHeaders"))
         rngHeaders = Worksheets(MODIFIED_AMF_SHEET_NAME).ListObjects(MODIFIED_AMF_TABLE_NAME).HeaderRowRange
        rngHeaders(.Item(1).Activate)
         rngPivotRange = Range(ActiveCell.Address, ActiveCell.(xlToRight).(xlDown).Address)
     }
    
  
    // Create pivot cache
     pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType=xlDatabase, SourceData=rngPivotRange)
        
    // Create pivot table
    Worksheets(.Add.Name = AMF_PIVOT_SHEET_NAME)
    //AMF_PIVOT_TABLE_NAME
     pvtTable = pvtCache.CreatePivotTable( _
        tabledestination(=Worksheets(AMF_PIVOT_SHEET_NAME).Range("a3"), _)
        tablename(=AMF_PIVOT_TABLE_NAME))
    Worksheets((AMF_PIVOT_SHEET_NAME).Tab.Color = vbBlue)
    
    // Disable automatic calculations for code execution speed
    pvtTable.ManualUpdate = true
       
    // Add fields to pivot table
    with( pvtTable) {
        .PivotFields((OCF_ACTIVITY_HEADER).Orientation = xlRowField)
        .PivotFields((HFM_ACCOUNT_HEADER).Orientation = xlRowField)
        .PivotFields((GL_ACCOUNT_HEADER).Orientation = xlRowField)
        .PivotFields((TRADING_PARTNER_HEADER).Orientation = xlRowField)
        .PivotFields((TE_BNS_PC_HEADER).Orientation = xlColumnField)
        // Add profit center as page filter in pivot table
        .PivotFields((GL_PC_HEADER).Orientation = xlPageField)
     }
    
    if( hasActivityContraBNSPC ) {
        // Exclude Contra-BNS profit center from base pivot
        with( pvtTable.PivotFields(GL_PC_HEADER)) {
            .CurrentPage( = "(All)")
            .EnableMultiplePageItems = true
            .PivotItems(CONTRA_BNS_PROFIT_CENTER).Visible = false
         }
     }
        
    // Format pivot table
    Application(.StatusBar = "Formatting pivot table ...")
    with( pvtTable) {
        .RowAxisLayout( xlTabularRow)
        .RepeatAllLabels( xlRepeatLabels)
        .RowGrand = true
        .ColumnGrand = false
        for ( var pvtField in .PivotFields) {
            pvtField.Subtotals(1) = false
            Application(.StatusBar = "Still working on formatting ...")
        } pvtField
     }
    
    // Add balance to pivot table + calculate
    Application(.StatusBar = "Calculating pivot table ...")
    pvtTable(.AddDataField pvtTable.PivotFields(BALANCE_HEADER), "Balance", xlSum)
    
    // Format Balance as Accounting number format
    pvtTable(.PivotFields("Balance").NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    
    // Move "TE" balance to first column
    pvtTable(.PivotFields("TE / BNS PC").PivotItems("TE").Position = 1)
    
    // Calculate pivot table
    pvtTable.ManualUpdate = false
    
    
    // Tidy up view
    pvtTable(.TableRange1.Columns.AutoFit)
    ActiveWindow(.Zoom = 85)
    
    // if( there has been activity in the Contra-BNS PC, pivot this and break it out
    if( hasActivityContraBNSPC ) {
        // Copy pivot table out for calculation of Contra-BNS profit center balances
         wsNewSheet = Worksheets.Add(before=Worksheets(MODIFIED_AMF_SHEET_NAME))
        wsNewSheet(.Name = CONTRA_BNS_PC_PIVOT_SHEET_NAME)
        wsNewSheet(.Tab.Color = vbBlue)
        pvtTable(.TableRange2.Copy Destination=wsNewSheet.Range("A1"))
         pvtTable = ActiveSheet.PivotTables(1)
    
        // Pivot for _only_ Contra-BNS profit center
        with( pvtTable.PivotFields(GL_PC_HEADER)) {
            .ClearAllFilters()
            .CurrentPage( = CONTRA_BNS_PROFIT_CENTER)
         }
        
        // 5/27/2015 -- Rename //Grand Total// to resolve localization + confusion issues
        Worksheets(CONTRA_BNS_PC_PIVOT_SHEET_NAME).Range("A4").(xlToRight).Value = TOTAL_ACCOUNT_BALANCE_HEADER
    
        // Calculate pivot table
        pvtTable.ManualUpdate = false
        
        // Tidy up view
        pvtTable(.TableRange1.Columns.AutoFit)
        ActiveWindow(.Zoom = 85)
     }
    
    // 5/27/2015 -- Rename //Grand Total// field to resolve localization + confusion issues
    Worksheets(AMF_PIVOT_SHEET_NAME).Range("A4").(xlToRight).Value = TOTAL_ACCOUNT_BALANCE_HEADER
    
    // Hide pivot table field list
    ActiveWorkbook.ShowPivotTableFieldList = false
    
    Application.StatusBar = false
    Application.ScreenUpdating = true
        
 }

function PivotAllAccounts(){
    var rngHeaders, rngPivotRange  Range
    var pvtCache  PivotCache
    var pvtTable  PivotTable
    var pvtField  PivotField
    var pvtPivotItem  PivotItem
    var wsNewSheet  Worksheet
    var strPivotItemName  
    
    // 5/26/2015 -- updated for pre-close data inclusion
    //AMF_RAW_DATA_SHEET_NAME = GetPersistentVariable(strVarName="AMF_RAW_DATA_SHEET_NAME")
    AMF_RAW_DATA_SHEET_NAME( = ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange.Value)
    
    Application(.StatusBar = "Pivoting all accounts ...")
    Application.ScreenUpdating = false
    Application.DisplayStatusBar = true
    
//    with( Worksheets(MODIFIED_AMF_SHEET_NAME)) {
//        .Activate
//        //6/11/2015 -- testing elimination of GetPersistentVariable call
//        // rngHeaders = Range(GetPersistentVariable("rngHeaders"))
//         rngHeaders = Worksheets(MODIFIED_AMF_SHEET_NAME).ListObjects(MODIFIED_AMF_TABLE_NAME).HeaderRowRange
//        rngHeaders.Item(1).Activate
//         rngPivotRange = Range(ActiveCell.Address, ActiveCell.(xlToRight).(xlDown).Address)
//     }
    
  
//    // Create pivot cache
//     pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType=xlDatabase, SourceData=rngPivotRange)
    // Lift existing PivotCache
     pvtCache = ActiveWorkbook.PivotCaches(AMF_PIVOT_CACHE_INDEX)
        
    // Create pivot table
     wsNewSheet = Worksheets.Add(before=Worksheets(PRE_CLOSE_BNS_INCOME_SHEET_NAME))
    wsNewSheet(.Name = ALL_ACCOUNTS_PIVOT_SHEET_NAME)
    //Worksheets.Add.Name = ALL_ACCOUNTS_PIVOT_SHEET_NAME
     pvtTable = pvtCache.CreatePivotTable( _
        tabledestination(=Worksheets(ALL_ACCOUNTS_PIVOT_SHEET_NAME).Range("a3"), _)
        tablename(=ALL_ACCOUNTS_PIVOT_TABLE_NAME))
    Worksheets((ALL_ACCOUNTS_PIVOT_SHEET_NAME).Tab.Color = vbBlue)
    
    // Disable automatic calculations for code execution speed
    pvtTable.ManualUpdate = true
       
    // Add fields to pivot table
    with( pvtTable) {
        .PivotFields((OCF_ACTIVITY_HEADER).Orientation = xlRowField)
        .PivotFields((HFM_ACCOUNT_HEADER).Orientation = xlRowField)
        .PivotFields((GL_ACCOUNT_HEADER).Orientation = xlRowField)
        .PivotFields((TRADING_PARTNER_HEADER).Orientation = xlRowField)
        .PivotFields((TE_BNS_PC_HEADER).Orientation = xlColumnField)
        // Add profit center as page filter in pivot table
        //.PivotFields(GL_PC_HEADER).Orientation = xlPageField
     }
    
//    if( hasActivityContraBNSPC ) {
//        // Exclude Contra-BNS profit center from base pivot
//        with( pvtTable.PivotFields(GL_PC_HEADER)) {
//            .CurrentPage = "(All)"
//            .EnableMultiplePageItems = true
//            .PivotItems(CONTRA_BNS_PROFIT_CENTER).Visible = false
//         }
//     }
        
    // Format pivot table
    Application(.StatusBar = "Formatting pivot table ...")
    with( pvtTable) {
        .RowAxisLayout( xlTabularRow)
        .RepeatAllLabels( xlRepeatLabels)
        .RowGrand = true
        .ColumnGrand = false
        for ( var pvtField in .PivotFields) {
            pvtField.Subtotals(1) = false
            Application(.StatusBar = "Still working on formatting ...")
        } pvtField
     }
    
    // Add balance to pivot table + calculate
    Application(.StatusBar = "Calculating pivot table ...")
    pvtTable(.AddDataField pvtTable.PivotFields(BALANCE_HEADER), "Balance", xlSum)
    
    // Format Balance as Accounting number format
    pvtTable(.PivotFields("Balance").NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    
    // Move "TE" balance to first column
    pvtTable(.PivotFields("TE / BNS PC").PivotItems("TE").Position = 1)
    
    // Calculate pivot table
    pvtTable.ManualUpdate = false
    
    
    // Tidy up view
    pvtTable(.TableRange1.Columns.AutoFit)
    ActiveWindow(.Zoom = 85)
    
//    // if( there has been activity in the Contra-BNS PC, pivot this and break it out
//    if( hasActivityContraBNSPC ) {
//        // Copy pivot table out for calculation of Contra-BNS profit center balances
//         wsNewSheet = Worksheets.Add(before=Worksheets(MODIFIED_AMF_SHEET_NAME))
//        wsNewSheet.Name = CONTRA_BNS_PC_PIVOT_SHEET_NAME
//        wsNewSheet.Tab.Color = vbBlue
//        pvtTable.TableRange2.Copy Destination=wsNewSheet.Range("A1")
//         pvtTable = ActiveSheet.PivotTables(1)
//
//        // Pivot for _only_ Contra-BNS profit center
//        with( pvtTable.PivotFields(GL_PC_HEADER)) {
//            .ClearAllFilters
//            .CurrentPage = CONTRA_BNS_PROFIT_CENTER
//         }
//
//        // 5/27/2015 -- Rename //Grand Total// to resolve localization + confusion issues
//        Worksheets(CONTRA_BNS_PC_PIVOT_SHEET_NAME).Range("A4").(xlToRight).Value = TOTAL_ACCOUNT_BALANCE_HEADER
//
//        // Calculate pivot table
//        pvtTable.ManualUpdate = false
//
//        // Tidy up view
//        pvtTable.TableRange1.Columns.AutoFit
//        ActiveWindow.Zoom = 85
//     }
    
    // 5/27/2015 -- Rename //Grand Total// field to resolve localization + confusion issues
    Worksheets(ALL_ACCOUNTS_PIVOT_SHEET_NAME).Range("A4").(xlToRight).Value = TOTAL_ACCOUNT_BALANCE_HEADER
    
    // Hide pivot table field list
    ActiveWorkbook.ShowPivotTableFieldList = false
    
    Application.StatusBar = false
    Application.ScreenUpdating = true
        
 }

function CopyAllAccountsPivot(){
    var rngCopyRange  Range
    var lngLastRow  
    
    //Worksheets(AMF_PIVOT_SHEET_NAME).Activate
    Worksheets(.Add.Name = ALL_ACCOUNTS_SHEET_NAME)
    ActiveSheet(.Tab.Color = CALC_MAPPING_TABLE_TAB_COLOR)
    
    Worksheets((ALL_ACCOUNTS_PIVOT_SHEET_NAME).Activate)
    // TEMPLATED  Dynamic range copy of pivot table -- headers-only
     rngCopyRange = Range(Range("a4", Range("a4").(xlToRight).(xlDown)).Address)
    rngCopyRange(.Select)
    rngCopyRange(.Copy Destination=Worksheets(ALL_ACCOUNTS_SHEET_NAME).Range("A1"))
    
    Worksheets((ALL_ACCOUNTS_SHEET_NAME).Activate)
    ActiveCell(.CurrentRegion.Activate)
    // Tidy up view
    Selection(.Columns.AutoFit)
    ActiveWindow(.Zoom = 85)
    
    lngLastRow = Range("a1").(xlDown).Row
    
    Worksheets(.Add.Name = OCF_ACTIVITY_TEMPLATE_SHEET_NAME)
    Worksheets((ALL_ACCOUNTS_SHEET_NAME).Activate)
    //Worksheets(ALL_ACCOUNTS_SHEET_NAME).Range("a1d1").(xlDown).Copy Destination=Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Range("a1")
    //Worksheets(ALL_ACCOUNTS_SHEET_NAME).Range(Range("a1"), Range("d" + lngLastRow).Address).Copy Destination=Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Range("a1")
    Worksheets(ALL_ACCOUNTS_SHEET_NAME).Range(Range("a1"), Range("d" + lngLastRow).Address).Copy
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Paste)
    
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Tab.Color = OCF_ACTIVITY_TEMPLATE_TAB_COLOR)
    
//    with( ActiveSheet.ListObjects(CONTRA_BNS_PC_BAL_TABLE_NAME).HeaderRowRange) {
//        .Font.Color = vbBlack
//     }
 }

function hasActivityContraBNSPC()  Boolean{
    var pvtPivotTable  PivotTable
    var pvtPivotField  PivotField
    var strCheckName  
    
    // Check AMF pivot table for existence of Contra-BNS profit center
     pvtPivotTable = Worksheets(AMF_PIVOT_SHEET_NAME).PivotTables(1)
    On( Error GoTo HandleError)
    // if( strCheckName can be set, Contra-BNS PC exists in the AMF pivot (and, therefore, the AMF data)
    strCheckName( = pvtPivotTable.PivotFields(GL_PC_HEADER).PivotItems(CONTRA_BNS_PROFIT_CENTER).Name)

HandleError()
    Select( Case Err)
        Case( 0)
            //MsgBox prompt="Activity found in Contra-BNS profit center!"
            // TODO Reset error handling here? on error goto 0
            hasActivityContraBNSPC = true
        Case }else{
            //MsgBox prompt="NO ACTIVITY FOUND in Contra-BNS profit center!"
            // TODO Reset error handling here? on error goto 0
            hasActivityContraBNSPC = false
     Select
 }

function PivotPreCloseAMFData(){
    var rngHeaders, rngPivotRange  Range
    var pvtCache  PivotCache
    var pvtTable  PivotTable
    var pvtField  PivotField
    var pvtPivotItem  PivotItem
    var wsNewSheet  Worksheet
    
    MODIFIED_AMF_SHEET_NAME = "Pre-Close " + ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange.Value
    
    Application(.StatusBar = "Pivoting Pre-Close All Mapping File data ...")
    Application.ScreenUpdating = false
    Application.DisplayStatusBar = true
    
    // TODO  Update rngHeaders recall to use RangeRowHeaders of AMF raw data table
    with( Worksheets(MODIFIED_AMF_SHEET_NAME)) {
        .Activate()
        //6/11/2015 -- testing elimination of GetPersistentVariable call
        // rngHeaders = Range(GetPersistentVariable("rngHeaders"))
         rngHeaders = Worksheets(MODIFIED_AMF_SHEET_NAME).ListObjects(1).HeaderRowRange
        MODIFIED_AMF_SHEET_NAME = "Pre-Close " + ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange.Value
        rngHeaders(.Item(1).Activate)
         rngPivotRange = Range(ActiveCell.Address, ActiveCell.(xlToRight).(xlDown).Address)
     }
    
  
    // Create pivot cache
     pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType=xlDatabase, SourceData=rngPivotRange)
    
    // Create pivot table
    Worksheets(.Add.Name = PRE_CLOSE_AMF_PIVOT_SHEET_NAME)
     pvtTable = pvtCache.CreatePivotTable( _
        tabledestination(=Worksheets(PRE_CLOSE_AMF_PIVOT_SHEET_NAME).Range("a3"), _)
        tablename(=PRE_CLOSE_AMF_PIVOT_TABLE_NAME))
    Worksheets((PRE_CLOSE_AMF_PIVOT_SHEET_NAME).Tab.Color = vbBlue)
    
    // Disable automatic calculations for code execution speed
    pvtTable.ManualUpdate = true
       
    // Add fields to pivot table
    with( pvtTable) {
        .PivotFields((OCF_ACTIVITY_HEADER).Orientation = xlRowField)
        .PivotFields((HFM_ACCOUNT_HEADER).Orientation = xlRowField)
        .PivotFields((GL_ACCOUNT_HEADER).Orientation = xlRowField)
        .PivotFields((TRADING_PARTNER_HEADER).Orientation = xlRowField)
        .PivotFields((TE_BNS_PC_HEADER).Orientation = xlColumnField)
        // Add profit center as page filter in pivot table
        .PivotFields((GL_PC_HEADER).Orientation = xlPageField)
     }

    //Exclude non P&L HFM accounts from table
    for ( var pvtPivotItem in pvtTable.PivotFields(HFM_ACCOUNT_HEADER).PivotItems) {
        // if( HFM account begins with 1, 2 or 3 ...
        if( InStr(HFM_ACCT_PL_EXCLUDE_RANGE, Left(pvtPivotItem.Name, 1)) ) {
            // Exclude HFM account from pivot table
            pvtPivotItem.Visible = false
        }else{
            // Otherwise, include HFM account -- i.e. HFM account is a P&L account
            pvtPivotItem.Visible = true
         }
    } pvtPivotItem
    
    // Format pivot table
    Application(.StatusBar = "Formatting pivot table ...")
    with( pvtTable) {
        .RowAxisLayout( xlTabularRow)
        .RepeatAllLabels( xlRepeatLabels)
        .RowGrand = true
        .ColumnGrand = false
        for ( var pvtField in .PivotFields) {
            pvtField.Subtotals(1) = false
            Application(.StatusBar = "Still working on formatting ...")
        } pvtField
     }
    
    // Add balance to pivot table + calculate
    Application(.StatusBar = "Calculating pivot table ...")
    pvtTable(.AddDataField pvtTable.PivotFields(BALANCE_HEADER), "Balance", xlSum)
    
    // Format Balance as Accounting number format
    pvtTable(.PivotFields("Balance").NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    
    // Move "TE" balance to first column
    pvtTable(.PivotFields("TE / BNS PC").PivotItems("TE").Position = 1)
    
    // Calculate pivot table
    pvtTable.ManualUpdate = false
    
    // 5/27/2015 -- Rename //Grand Total// field to resolve localization + confusion issues
    Worksheets(PRE_CLOSE_AMF_PIVOT_SHEET_NAME).Range("A4").(xlToRight).Value = TOTAL_ACCOUNT_BALANCE_HEADER
    
    // Tidy up view
    pvtTable(.TableRange1.Columns.AutoFit)
    ActiveWorkbook.ShowPivotTableFieldList = false
    Application.StatusBar = false
    ActiveWindow(.Zoom = 85)
 }

function ModifyPreCloseAMF(){
    var rngFirstCellOfHeaders, rngHeaders  Range
    var lngLastRow  
    
    Call( InitializeTemplateGlobals)
    // TEMP modify subroutine for pre-close AMF specifics
    AMF_RAW_DATA_SHEET_NAME( = ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange.Value)
    MODIFIED_AMF_SHEET_NAME = "Pre-Close " + AMF_RAW_DATA_SHEET_NAME
    
    //  first cell of headers per Greg Stanko//s assurance that this can be relied on from FDM
     rngFirstCellOfHeaders = Worksheets(AMF_RAW_DATA_SHEET_NAME).Range("a5")
    
    Application(.StatusBar = "Modifying Pre-Close All Mapping File preparatory for OCF analysis ...")
    // Turn off screen updating to speed function execution
    Application.ScreenUpdating = false
    
    // Copy worksheet and perform modifications in copied table
    rngFirstCellOfHeaders(.Worksheet.Copy before=rngFirstCellOfHeaders.Worksheet)
    ActiveSheet(.Name = MODIFIED_AMF_SHEET_NAME)
    ActiveSheet(.Tab.Color = vbYellow)
    ActiveSheet(.Unprotect Password=WORKSHEET_PASSWORD)
    
    // Update pointer to first cell of headers on copied sheet
     rngFirstCellOfHeaders = Worksheets(MODIFIED_AMF_SHEET_NAME).Range(rngFirstCellOfHeaders.Address)
    // Identify last row
    lngLastRow = rngFirstCellOfHeaders.(xlDown).Row
    
    // Identify range of table headers
    // This will need to be updated after columns are inserted to keep it current
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    // Tidy-up raw All Mapping File
    // Insert row above headers to enable clean CurrentRegion selection
    rngFirstCellOfHeaders(.EntireRow.Insert)
    // Delete two incongruous rows below header row of AMF
    rngFirstCellOfHeaders(.Offset(1, 0).EntireRow.Delete)
    rngFirstCellOfHeaders(.Offset(1, 0).EntireRow.Delete)
    
    
    // Create data table (ListObject) from copy of All Mapping File
    rngFirstCellOfHeaders(.CurrentRegion.Select)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = MOD_PRE_CLOSE_AMF_TABLE_NAME)
    
    // Insert OCF Activity lookup from HFM Account value
    rngHeaders(.Find(what=HFM_ACCOUNT_HEADER, lookat=xlWhole).Offset(0, 1).Activate)
    ActiveCell(.EntireColumn.Insert)
    // Due to insert, ActiveCell now points to cell in newly-inserted column
    with( ActiveCell) {
        .Value( = OCF_ACTIVITY_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font(.Color = vbBlack)
     }
    //  formula for lookup of OCF Activity
    ActiveCell(.Offset(1, 0).Formula = "=VLOOKUP([@[HFM_Account]],HFM_Acct_OCF_Activity,3,FALSE)")

    // Insert Trading Partner lookup from GL_ICP
    rngHeaders(.Find(what=GL_ICP_HEADER, lookat=xlWhole).Offset(0, 1).Activate)
    ActiveCell(.EntireColumn.Insert)
    with( ActiveCell) {
        .Value( = TRADING_PARTNER_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font(.Color = vbBlack)
     }
    //  formula for VLOOKUP of Trading Partner
    Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
        Case( "SAP PR2")
            ActiveCell.Offset(1, 0).Formula = "=IF(NOT(ISERROR(FIND(""" + ICP_FIELD_SEARCH_VALUE + """,[@[GL_ICP]]))),"""",[@[GL_ICP]])"
        Case( "C1")
            ActiveCell(.Offset(1, 0).Value = " ")
            ActiveCell(.Offset(1, 0).AutoFill Destination=Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(lngLastRow - 1, 0)))
        Case( "AMPICS")
            ActiveCell(.Offset(1, 0).Value = " ")
            ActiveCell(.Offset(1, 0).AutoFill Destination=Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(lngLastRow - 1, 0)))
     Select
    
    // Insert TE / BNS Profit Center lookup from GL_PC
    rngHeaders(.Find(GL_PC_HEADER).Offset(0, 1).Activate)
    ActiveCell(.EntireColumn.Insert)
    with( ActiveCell) {
        .Value( = TE_BNS_PC_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font(.Color = vbBlack)
     }
    //  formula for VLOOKUP of TE / BNS Profit Center
    //ActiveCell.Offset(1, 0).Formula = "=IFERROR(VLOOKUP(VALUE([@[GL_PC]]),BNS_Profit_Centers,2,FALSE),""" + TE_PROFIT_CENTER_FLAG + """)"
    Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
        Case( "SAP PR2")
            ActiveCell.Offset(1, 0).Formula = "=IFERROR(VLOOKUP(VALUE([@[GL_PC]]),BNS_Profit_Centers,2,FALSE),""" + TE_PROFIT_CENTER_FLAG + """)"
        Case( "C1")
            ActiveCell.Offset(1, 0).Formula = "=IFERROR(VLOOKUP([@[GL_PC]],BNS_Profit_Centers,2,FALSE),""" + TE_PROFIT_CENTER_FLAG + """)"
        Case( "AMPICS")
            ActiveCell.Offset(1, 0).Formula = "=IFERROR(VLOOKUP([@[GL_PC]],BNS_Profit_Centers,2,FALSE),""" + TE_PROFIT_CENTER_FLAG + """)"
     Select
    
    // Identify range of table headers
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    // Tidy up
    Worksheets((MODIFIED_AMF_SHEET_NAME).Activate)
    
    //ActivateFilters rngHeaders
    Worksheets((MODIFIED_AMF_SHEET_NAME).ListObjects(MOD_PRE_CLOSE_AMF_TABLE_NAME).HeaderRowRange.CurrentRegion.Columns.AutoFit)
    Worksheets((MODIFIED_AMF_SHEET_NAME).ListObjects(MOD_PRE_CLOSE_AMF_TABLE_NAME).HeaderRowRange.Item(1).Activate)
    ActiveWindow(.Zoom = 85)
    
    Application.StatusBar = false
    Application.ScreenUpdating = true
 }

// Prepare pre-close AMF data
function aa_PreparePreCloseAMFData(){
    Call( ModifyPreCloseAMF)
    Call( PivotPreCloseAMFData)
    Call( CopyPreCloseAMFPivotData)
    // Focus on input tab
    Worksheets("Input + Assumptions").Activate
    ActiveWorkbook(.Names("input_Contra_BNS_PC").RefersToRange.Activate)
 }

// Copy pivoted data out into Activity Template
function CopyAMFPivotData(){
    var rngCopyRange  Range
    var lngLastRow  
    
    Worksheets((AMF_PIVOT_SHEET_NAME).Activate)
    //Worksheets.Add.Name = OCF_ACTIVITY_TEMPLATE_SHEET_NAME
    Worksheets(.Add.Name = MODDED_AMF_PIVOT_ACCOUNTS_SHEET_NAME)
    ActiveSheet(.Tab.Color = CALC_MAPPING_TABLE_TAB_COLOR)
    
    Worksheets((AMF_PIVOT_SHEET_NAME).Activate)
    lngLastRow = Range("a4").(xlDown).Row
    // TEMPLATED  Dynamic range copy of pivot table -- headers-only
     rngCopyRange = Range("a4", Range("a4").(xlToRight).(xlDown).Address)
    rngCopyRange(.Select)
    rngCopyRange(.Copy Destination=Worksheets(MODDED_AMF_PIVOT_ACCOUNTS_SHEET_NAME).Range("A1"))
    
    Worksheets((MODDED_AMF_PIVOT_ACCOUNTS_SHEET_NAME).Activate)
    Worksheets((MODDED_AMF_PIVOT_ACCOUNTS_SHEET_NAME).Range("a1").Activate)
    ActiveCell(.CurrentRegion.Select)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = MODDED_AMF_PIVOT_ACCOUNTS_TABLE_NAME)
    
    // Tidy up view
    Selection(.Columns.AutoFit)
    ActiveWindow(.Zoom = 85)
//    with( ActiveSheet.ListObjects(CONTRA_BNS_PC_BAL_TABLE_NAME).HeaderRowRange) {
//        .Font.Color = vbBlack
//     }
 }

// Copy pivoted data out into Activity Template
function CopyContraBNSPCPivotData(){
    var rngCopyRange  Range

    // Copy header-down from Contra-BNS PC balance pivot table
    Worksheets((CONTRA_BNS_PC_PIVOT_SHEET_NAME).Activate)
    
     rngCopyRange = Range(Worksheets(CONTRA_BNS_PC_PIVOT_SHEET_NAME).Range("a4", Range("a4").(xlToRight).(xlDown)).Address)
    
    // Add Contra-BNS PC worksheet + copy table to it
    Worksheets(.Add after=Worksheets(BNS_PROFIT_CENTERS_SHEET_NAME))
    ActiveSheet(.Name = CONTRA_BNS_PC_BAL_SHEET_NAME)
    Worksheets((CONTRA_BNS_PC_BAL_SHEET_NAME).Tab.Color = vbMagenta)
    rngCopyRange(.Copy Destination=Worksheets(CONTRA_BNS_PC_BAL_SHEET_NAME).Range("a1"))
    
    // Create data table in Contra-BNS PC sheet
    Worksheets((CONTRA_BNS_PC_BAL_SHEET_NAME).Range("a1").Activate)
    ActiveCell(.CurrentRegion.Select)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = CONTRA_BNS_PC_BAL_TABLE_NAME)
    
    // Tidy up view
    Selection(.Columns.AutoFit)
    ActiveWindow(.Zoom = 85)
//    with( ActiveSheet.ListObjects(CONTRA_BNS_PC_BAL_TABLE_NAME).HeaderRowRange) {
//        .Font.Color = vbBlack
//     }
 }

// Copy pivoted pre-close income data out into data table
function CopyPreCloseAMFPivotData(){
    var rngCopyRange  Range
    
    // Initialize globals
    Call( InitializeTemplateGlobals)
    
    Worksheets((PRE_CLOSE_AMF_PIVOT_SHEET_NAME).Activate)
    //lngLastRow = Range("A4").(xlDown).Row
    Worksheets(.Add.Name = PRE_CLOSE_BNS_INCOME_SHEET_NAME)
    Worksheets((PRE_CLOSE_BNS_INCOME_SHEET_NAME).Tab.Color = vbMagenta)
    
    Worksheets((PRE_CLOSE_AMF_PIVOT_SHEET_NAME).Activate)
     rngCopyRange = Range(Range("a4", Range("a4").(xlToRight).(xlDown)).Address)
    rngCopyRange(.Select)
    rngCopyRange(.Copy Destination=Worksheets(PRE_CLOSE_BNS_INCOME_SHEET_NAME).Range("A1"))
    
    // Create data table in Pre-Close BNS Income sheet
    Worksheets((PRE_CLOSE_BNS_INCOME_SHEET_NAME).Activate)
    Worksheets((PRE_CLOSE_BNS_INCOME_SHEET_NAME).Range("a1").Activate)
    ActiveCell(.CurrentRegion.Select)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = PRE_CLOSE_BNS_INCOME_TABLE_NAME)

    // Tidy up view
    Selection(.Columns.AutoFit)
    ActiveWindow(.Zoom = 85)
    with( ActiveSheet.ListObjects(PRE_CLOSE_BNS_INCOME_TABLE_NAME).HeaderRowRange) {
        .Font(.Color = vbBlack)
     }
    
 }

// Manipulate Activity Template data
function ModifyActivityTemplateData()  Integer{
    var rngFirstCellOfHeaders, rngHeaders, rngModifiedAMFData, rngFirstCellModifiedAMFData, rngLastCellModifiedAMFData, _
        rngTmpRange, rngClearingHeaders, rngIsClearingAcct  Range
    var lngLastRow, lngAMFLastRow  
    var i  Integer
    var strTempString, strTmpFormula, strTmpFormula2, strBNSPC  
    
    // 7/28/2015 -- moved this array into global variables section
    // Declare + initialize array with clearing entry headers
//    var CLEARING_ENTRY_HEADERS(1 To 9)  
//    CLEARING_ENTRY_HEADERS(1) = "Doc Date"
//    CLEARING_ENTRY_HEADERS(2) = "Posting Date"
//    CLEARING_ENTRY_HEADERS(3) = "PK"
//    CLEARING_ENTRY_HEADERS(4) = "Account"
//    CLEARING_ENTRY_HEADERS(5) = "PC"
//    CLEARING_ENTRY_HEADERS(6) = "CC"
//    CLEARING_ENTRY_HEADERS(7) = "Trad Ptr"
//    CLEARING_ENTRY_HEADERS(8) = "Amount"
//    CLEARING_ENTRY_HEADERS(9) = "Text"
    
    Application.ScreenUpdating = false
    Application.DisplayStatusBar = true
    
    Application(.StatusBar = "Building OCF activity template ...")
    
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Activate)
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Range("A1").Activate)
    lngLastRow = ActiveCell.(xlDown).Row
    
     rngFirstCellOfHeaders = ActiveSheet.Range("A1")
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    Worksheets((MODIFIED_AMF_SHEET_NAME).Activate)
    //6/11/2015 -- remove call to GetPersistentVariable
    // rngFirstCellModifiedAMFData = Worksheets(MODIFIED_AMF_SHEET_NAME).Range(GetPersistentVariable("rngHeaders")).Item(1)
     rngFirstCellModifiedAMFData = Worksheets(MODIFIED_AMF_SHEET_NAME).ListObjects(1).HeaderRowRange.Item(1)
    //MsgBox (rngFirstCellModifiedAMFData.Address)
    //MsgBox ("Last cell = " + rngFirstCellModifiedAMFData.(xlToRight).(xlDown).Address)
     rngLastCellModifiedAMFData = rngFirstCellModifiedAMFData.(xlToRight).(xlDown)
    //MsgBox (rngLastCellModifiedAMFData.Address)
     rngModifiedAMFData = Range(rngFirstCellModifiedAMFData.Address, rngLastCellModifiedAMFData.Address)
    //MsgBox (rngModifiedAMFData.Address)
    //MsgBox (rngModifiedAMFData.Parent.Name + " - " + rngModifiedAMFData.Address)
    
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Activate)
    
    Application(.StatusBar = "Adding GL Account Description ...")
    // Insert column for GL Account Description
    rngHeaders(.Find(GL_ACCOUNT_HEADER).Offset(0, 1).Activate)
    ActiveCell(.EntireColumn.Insert)
    // Insert GL Account Description lookup field
    ActiveCell(.Value = GL_ACCOUNT_DESCRIPTION_HEADER)
    ActiveCell(.Interior.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
    
    lngAMFLastRow = rngModifiedAMFData.Item(1).(xlDown).Row

    //  formula for VLOOKUP of GL Account Description
    // 6/3/2015 -- revised to use GL_Other in VLOOKUP
    ActiveCell.Offset(1, 0).Formula = "=INDEX(Modded_AMF_Table[#All], MATCH($C2,Modded_AMF_Table[[#All],[GL_Other]],0), MATCH(""" + AMF_GL_ACCT_DESC_HEADER + """,Modded_AMF_Table[#Headers],0))"
    // Fill-down formula
    ActiveCell(.Offset(1, 0).AutoFill Destination=Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(lngLastRow - 1, 0)))
    
    //6/30/2015 -- v1.33 -- removed for all-account mapping
//    // Cut + paste total account balance before TE + BNS balances (for usability)
//    Range(rngHeaders.Find(TOTAL_ACCOUNT_BALANCE_HEADER).Address, rngHeaders.Find(TOTAL_ACCOUNT_BALANCE_HEADER).Offset(lngLastRow, 0)).Cut //Destination=rngHeaders.Find(TE_BALANCE_HEADER)
//    rngHeaders.Find(TE_BALANCE_HEADER).Insert shift=xlToRight
    
        // Create data table (ListObject) from activity template bare-bones data
    // NB -- CANNOT CREATE TABLE BEFORE MOVING GRAND TOTAL COLUMN!  CAN//T CUT + PASTE COLUMNS IN A TABLE!
    // 6/30/2015 -- This can be done if you cut + paste the entire column
    rngHeaders(.Item(1).Activate)
    ActiveCell(.CurrentRegion.Select)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = OCF_ACTIVITY_TEMPLATE_TABLE_NAME)
    with( ActiveSheet.ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME).HeaderRowRange) {
        .Font(.Color = vbBlack)
     }
    
    // 6/30/2015 -- v1.33 -- added mapping for [TE + BNS Balance]
    rngHeaders(.Find(TRADING_PARTNER_HEADER).Offset(0, 1).Activate)
    ActiveCell(.Value = TOTAL_ACCOUNT_BALANCE_HEADER)
    //=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[TE + BNS Balance]),2,FALSE),0)
    //ActiveCell.Offset(1, 0).FormulaArray = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[TE + BNS Balance]),2,FALSE),0)"
    ActiveCell(.Offset(1, 0).Activate)
    with( ActiveCell) {
        strTmpFormula( = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],""XXX"",2,FALSE),0)")
        strTmpFormula2 = "CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[" + TOTAL_ACCOUNT_BALANCE_HEADER + "])"
        .FormulaArray( = strTmpFormula)
        strTempString( = """XXX""")
        .Replace( what=strTempString, replacement=strTmpFormula2, lookat=xlPart)
     }
    //ActiveCell.Offset(1, 0).Activate
    ActiveCell(.AutoFill Destination=Range(ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address))
    Range((ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address).NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    
    // 6/30/2015 -- v1.33 -- added mapping for [TE] balance
    //ActiveCell.Offset(1, 0).Activate
    rngHeaders(.Find(TRADING_PARTNER_HEADER).Offset(0, 2).Activate)
    ActiveCell(.Value = TE_PROFIT_CENTER_FLAG)
    //ActiveCell.Offset(1, 0).FormulaArray = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[""" + TE_PROFIT_CENTER_FLAG + """]),2,FALSE),0)"
    ActiveCell(.Offset(1, 0).Activate)
    with( ActiveCell) {
        strTmpFormula( = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],""XXX"",2,FALSE),0)")
        strTmpFormula2 = "CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[" + TE_PROFIT_CENTER_FLAG + "])"
        .FormulaArray( = strTmpFormula)
        strTempString( = """XXX""")
        .Replace( what=strTempString, replacement=strTmpFormula2, lookat=xlPart)
     }
    //ActiveCell.Offset(1, 0).Activate
    ActiveCell(.AutoFill Destination=Range(ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address))
    Range((ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address).NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    
    // 6/30/2015 -- v1.33 -- added mapping for [BNS] balance
    //ActiveCell.Offset(1, 0).Activate
    rngHeaders(.Find(TRADING_PARTNER_HEADER).Offset(0, 3).Activate)
    ActiveCell(.Value = ActiveWorkbook.Names("admin_BNS_profit_center_flag").RefersToRange.Value)
    //ActiveCell.Offset(1, 0).FormulaArray = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[""" + ActiveWorkbook.Names("admin_BNS_profit_center_flag").RefersToRange.Value + """]),2,FALSE),0)"
    ActiveCell(.Offset(1, 0).Activate)
    with( ActiveCell) {
        strTmpFormula( = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],""XXX"",2,FALSE),0)")
        strTmpFormula2 = "CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[" + ActiveWorkbook.Names("admin_BNS_profit_center_flag").RefersToRange.Value + "])"
        //strBNCPC = ActiveWorkbook.Names("admin_BNS_profit_center_flag").RefersToRange.Value
        //strTmpFormula2 = "CHOOSE({1,2,3},Modded_AMF_Pivot_Accounts[HFM_Account]&Modded_AMF_Pivot_Accounts[GL_Other]&Modded_AMF_Pivot_Accounts[Trading_Partner],Modded_AMF_Pivot_Accounts[""" + strBNSPC + """])"
        .FormulaArray( = strTmpFormula)
        strTempString( = """XXX""")
        .Replace( what=strTempString, replacement=strTmpFormula2, lookat=xlPart)
     }
    //ActiveCell.Offset(1, 0).Activate
    ActiveCell(.AutoFill Destination=Range(ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address))
    Range((ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address).NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    rngHeaders(.Find(TRADING_PARTNER_HEADER).Offset(1, 0).Activate)
    Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
        Case( "C1")
            Range((ActiveCell.Address, ActiveCell.Offset(lngLastRow - 1, 0)).Value = " ")
        Case( "AMPICS")
            Range((ActiveCell.Address, ActiveCell.Offset(lngLastRow - 1, 0)).Value = " ")
     Select
    
    // Add Pre-Close Income
    // Pre-close income mapping formula = =IFERROR(VLOOKUP([@[GL_Other]]&[@[Trading_Partner]],CHOOSE({1,2},Pre_Close_BNS_Income[GL_Other]&Pre_Close_BNS_Income[Trading_Partner],Pre_Close_BNS_Income[BNS]),2,FALSE),0)
    Application(.StatusBar = "Adding pre-close income fields")
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    ActiveCell(.Value = BNS_PRE_CLOSE_INCOME_TE_HEADER)
    ActiveCell.Font.Bold = true
    ActiveCell(.Interior.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
    ActiveCell(.Offset(1, 0).Activate)
    
    // Check for likely profit center configuration error
    On( Error GoTo HANDLE_PC_ERROR)
    ActiveCell(.FormulaArray = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],CHOOSE({1,2,3},Pre_Close_BNS_Income[HFM_Account]&Pre_Close_BNS_Income[GL_Other]&Pre_Close_BNS_Income[Trading_Partner],Pre_Close_BNS_Income[BNS]),2,FALSE),0)")
    // if( above is successfulsuccessful, reset error handler
    On( Error GoTo 0)
    ActiveCell(.NumberFormat = ACCOUNTING_NUMBER_FORMAT)
        
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    ActiveCell(.Value = BNS_POST_CLOSE_INCOME_CS_HEADER)
    ActiveCell.Font.Bold = true
    ActiveCell(.Interior.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
    ActiveCell(.Offset(1, 0).Activate)
    ActiveCell(.Formula = "=[@BNS]-[@[BNS PRE-Close Income - TE]]")
    ActiveCell(.NumberFormat = ACCOUNTING_NUMBER_FORMAT)
         
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    Application(.StatusBar = "Adding Contra-BNS Profit Center balance ...")
    // Add Contra-BNS Profit Center field + lookup
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    ActiveCell(.Value = CONTRA_BNS_PC_HEADER)
    ActiveCell.Font.Bold = true
    ActiveCell(.Interior.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
    ActiveCell(.Offset(1, 0).Activate)
    if( hasActivityContraBNSPC ) {
        ActiveCell.FormulaArray = "=IFERROR(VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],CHOOSE({1,2,3},Contra_BNS_PC[HFM_Account]&Contra_BNS_PC[GL_Other]&Contra_BNS_PC[Trading_Partner],Contra_BNS_PC[TE + BNS Balance]),2,FALSE),0)"
    }else{
        ActiveCell(.Value = 0)
     }
    ActiveCell(.NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    ActiveCell(.AutoFill Destination=Range(ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address))
    
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    Application(.StatusBar = "Adding IsClearingAccount field ...")
    // Add IsClearingAccount field
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    ActiveCell(.Value = IS_CLEARING_ACCOUNT_HEADER)
    ActiveCell.Font.Bold = true
    ActiveCell(.Interior.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
    ActiveCell(.Offset(1, 0).Activate)
    // 6/30/2015 -- v1.33 -- Updated to lookup account directly as clearing accounts are now in the OCF Activity Table
    //ActiveCell.Formula = "=IFERROR(VLOOKUP([@[GL_Other]],GL_Acct_to_Clearing_Acct,2,FALSE)=[@[GL_Other]],FALSE)"
    ActiveCell(.Formula = "=IFERROR(VLOOKUP([@[GL_Other]],GL_Acct_to_Clearing_Acct[Clearing Acct],1,FALSE)=[@[GL_Other]],FALSE)")
    
    ActiveCell(.AutoFill Destination=Range(ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address))
       
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    Application(.StatusBar = "Adding HasClearingAccount field ...")
    // add HasClearingAcct field
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    with( ActiveCell) {
        .Value( = HAS_CLEARING_ACCOUNT_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font.Bold = true
     }
    with( ActiveCell.Offset(1, 0)) {
        .Activate()
        .Formula( = "=NOT(ISERROR(VLOOKUP([@[GL_Other]],GL_Acct_to_Clearing_Acct[#All],2,FALSE)))")
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    Application(.StatusBar = "Adding Clearing Account Balance ...")
    // Add Clearing Account Balance field
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    with( ActiveCell) {
        .Value( = CLEARING_ACCOUNT_BALANCE_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font.Bold = true
     }
    with( ActiveCell.Offset(1, 0)) {
        .Activate()
        .NumberFormat( = ACCOUNTING_NUMBER_FORMAT)
        if( hasActivityContraBNSPC ) {
            // Break-out formula to deal with 255-character limit of FormulaArray property
            strTmpFormula( = "=IF([@[Is Clearing Acct?]],0,IFERROR(VLOOKUP([@[HFM_Account]]&""XXX""&[@[Trading_Partner]],CHOOSE({1,2,3},Contra_BNS_PC[HFM_Account]&Contra_BNS_PC[GL_Other]&Contra_BNS_PC[Trading_Partner],Contra_BNS_PC[TE]),2,FALSE),0))")
            strTmpFormula2( = "VLOOKUP([@[GL_Other]],GL_Acct_to_Clearing_Acct,2,FALSE)")
            .FormulaArray( = strTmpFormula)
            strTempString( = """XXX""")
            .Replace( what=strTempString, replacement=strTmpFormula2, lookat=xlPart)
        }else{
            .Value( = 0)
         }
            .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    // 6/30/2015 -- v1.33 -- removed Adjustment to BNS field per requirements
//    Application.StatusBar = "Adding Adjustment to BNS field ..."
//    // Add Adjustment to BNS header
//    rngHeaders.(xlToRight).Offset(0, 1).Activate
//    with( ActiveCell) {
//        .Value = ADJUSTMENT_TO_BNS_HEADER
//        .Interior.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR
//        .Font.Bold = true
//     }
//    // Format entire column as Accounting
//    with( Range(ActiveCell.Offset(1, 0).Address, ActiveCell.Offset(lngLastRow - 1, 0).Address)) {
//        .NumberFormat = ACCOUNTING_NUMBER_FORMAT
//     }
//
//    // Update header range variable
//     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)

//    // Add conditional formatting to Adjustment to BNS field
//    with( Range(rngHeaders.Find(ADJUSTMENT_TO_BNS_HEADER).Offset(1, 0).Address, rngHeaders.Find(ADJUSTMENT_TO_BNS_HEADER).Offset(lngLastRow - 1).Address)) {
//         rngTmpRange = Range(rngHeaders.Find(OCF_ACTIVITY_HEADER).Offset(1, 0).Address)
//        .FormatConditions.Add Type=xlExpression, Formula1="=" + rngTmpRange.Address(rowabsolute=false, columnabsolute=true) + "<>""Direct"""
//        .FormatConditions(1).Interior.Pattern = xlPatternGray50
//        .Interior.Color = ENTRY_FIELD_BKG_COLOR
//        //.AutoFill Destination=Range(.Address, .Offset(lngLastRow - 1, 0).Address)
//     }
//
//    // Update header range variable
//     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    // Add BNS YTD Balance header
    Application(.StatusBar = "Adding BNS YTD Balance field ...")
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    with( ActiveCell) {
        .Value( = BNS_YTD_BALANCE_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font.Bold = true
     }
    // Format entire field as Accounting
    with( Range(ActiveCell.Offset(1, 0).Address, ActiveCell.Offset(lngLastRow - 1, 0).Address)) {
        .NumberFormat( = ACCOUNTING_NUMBER_FORMAT)
     }
    
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
     rngIsClearingAcct = rngHeaders.Find(what=IS_CLEARING_ACCOUNT_HEADER, lookat=xlWhole)
    
    // Add conditional formatting to BNS YTD Balance field
    with( Range(rngHeaders.Find(BNS_YTD_BALANCE_HEADER).Offset(1, 0).Address, rngHeaders.Find(BNS_YTD_BALANCE_HEADER).Offset(lngLastRow - 1).Address)) {
         rngTmpRange = Range(rngHeaders.Find(OCF_ACTIVITY_HEADER).Offset(1, 0).Address)
        // 7/30/2015 -- revised for single "Exclude" classifier
        //ActiveWorkbook.Names("admin_tmp_formula_local").RefersToRange.Formula = "=or(" + rngTmpRange.Address(rowabsolute=false, columnabsolute=true) + "=""Direct""," _
            + rngTmpRange.Address(rowabsolute=false, columnabsolute=true) + "=""" + OCF_ACTIVITY_EXCLUDE1 + """," _
            + rngTmpRange.Address(rowabsolute=false, columnabsolute=true) + "=""" + OCF_ACTIVITY_EXCLUDE2 + """)"
        ActiveWorkbook.Names("admin_tmp_formula_local").RefersToRange.Formula = "=or(" + rngTmpRange.Address(rowabsolute=false, columnabsolute=true) + "=""Direct""," _
            + rngTmpRange.Address(rowabsolute=false, columnabsolute=true) + "=""" + OCF_ACTIVITY_EXCLUDE1 + """)"
        .FormatConditions(.Add Type=xlExpression, Formula1=ActiveWorkbook.Names("admin_tmp_formula_local").RefersToRange.FormulaLocal)
        //BOOKMARK -- TODO  Get this conditional formatting right for clearing accounts
        .FormatConditions.Add Type=xlExpression, Formula1=rngIsClearingAcct.Address(rowabsolute=false, columnabsolute=true)
        .FormatConditions((1).Interior.Pattern = xlPatternGray50)
        .FormatConditions((2).Interior.Pattern = xlPatternGray50)
        .Interior(.Color = ENTRY_FIELD_BKG_COLOR)
        //.AutoFill Destination=Range(.Address, .Offset(lngLastRow - 1, 0).Address)
     }
    
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    Application(.StatusBar = "Adding BNS Amount to Clear ...")
    // Add BNS Amount to Clear field
    rngHeaders.(xlToRight).Offset(0, 1).Activate
    with( ActiveCell) {
        .Value( = BNS_AMOUNT_TO_CLEAR_HEADER)
        .Interior(.ColorIndex = MODIFIED_HEADER_BACKGROUND_COLOR)
        .Font.Bold = true
     }
    with( ActiveCell.Offset(1, 0)) {
        .Activate()
        .NumberFormat( = ACCOUNTING_NUMBER_FORMAT)
        // 6/30/2015 -- v1.33 -- removed Adjustment to BNS field per requirements
        //.Formula = "=IF(OR([@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE1 + """, [@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE2 + """,[@[Is Clearing Acct?]]),0,IF([@[OCF_Activity]]=""Direct"",-[@[CS-Owned BNS Balance]]-[@[Adjustment to BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]],-[@[CS-Owned BNS Balance]]-[@[Total BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]]))"
        // 7/30/2015 -- revised for single "Exclude" classifier
        //.Formula = "=IF(OR([@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE1 + """, [@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE2 + """,[@[Is Clearing Acct?]]),0,IF([@[OCF_Activity]]=""Direct"",-[@[CS-Owned BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]],-[@[CS-Owned BNS Balance]]-[@[Total BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]]))"
        .Formula = "=IF(OR([@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE1 + """,[@[Is Clearing Acct?]]),0,IF([@[OCF_Activity]]=""Direct"",-[@[CS-Owned BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]],-[@[CS-Owned BNS Balance]]-[@[Total BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]]))"
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Update header range variable
     rngHeaders = Range(rngFirstCellOfHeaders.Address, rngFirstCellOfHeaders.(xlToRight).Address)
    
    Application(.StatusBar = "Adding GL clearing entry headers ...")
    // Add clearing entry headers
    rngHeaders.(xlToRight).Offset(0, 3).Activate
    for( i = 0 To UBound(CLEARING_ENTRY_HEADERS, 1) - 1) {
        with( ActiveCell.Offset(0, i)) {
            .Value( = CLEARING_ENTRY_HEADERS(i + 1))
            .Interior(.Color = RGB(79, 129, 189))
            .Font(.Color = RGB(255, 255, 255))
            .Font.Bold = true
         }
    } i
    
    // Initialize rngClearingHeaders variable
     rngClearingHeaders = Range(ActiveCell.Address, ActiveCell.(xlToRight).Address)
    
    // Create clearing entries table before populating fields (ListObject)
    rngClearingHeaders(.Item(1).Activate)
    ActiveCell(.CurrentRegion.Select)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = CLEARING_JE_TABLE_NAME)
    ActiveSheet.ListObjects(CLEARING_JE_TABLE_NAME).Resize Range(rngClearingHeaders.Item(1).Address, Cells(lngLastRow, rngClearingHeaders.Item(1).(xlToRight).Column))
    
    Application(.StatusBar = "Calculating clearing entries (for Direct accounts only) ...")
    
    // Add "Doc Date" to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(1), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
        .Value( = ActiveWorkbook.Names("input_SAP_Document_Date").Value)
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Add "Posting Date" to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(2), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
        .Value( = ActiveWorkbook.Names("input_SAP_Posting_Date").Value)
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Add Posting Key ("PK") to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(3), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
        .Formula = "=IF(OCF_Activity_Template_Table[@[BNS Amount to Clear]]>=0,""" + GL_DEBIT_CODE + """,""" + GL_CREDIT_CODE + """)"
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
        
    // Add Account to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(4), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
        .Formula( = "=IF(NOT(ISERROR(VLOOKUP(OCF_Activity_Template_Table[@[GL_Other]],GL_Acct_to_Clearing_Acct,2,FALSE))),VLOOKUP(OCF_Activity_Template_Table[@[GL_Other]],GL_Acct_to_Clearing_Acct,2,FALSE),OCF_Activity_Template_Table[@[GL_Other]])")
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Add Contra-BNS profit center to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(5), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
        .Value( = CONTRA_BNS_PROFIT_CENTER)
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Add Cost Center to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(6), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
        .Formula( = "=IFERROR(VLOOKUP(OCF_Activity_Template_Table[@[HFM_Account]],HFM_Acct_to_CC[#All],2,FALSE),"""")")
     }
    
    // Add Trading Partner to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(7), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
         rngTmpRange = Range(rngHeaders.Find(TRADING_PARTNER_HEADER).Offset(1, 0).Address)
        .Formula( = "=IF(OCF_Activity_Template_Table[@[Trading_Partner]]="""","""",OCF_Activity_Template_Table[@[Trading_Partner]])")
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    // Add Amount to clearing entries -- must be ABS(BNS_AMOUNT_TO_CLEAR) for SAP
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(8), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
         rngTmpRange = Range(rngHeaders.Find(BNS_AMOUNT_TO_CLEAR_HEADER).Offset(1, 0).Address)
        Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
            Case( "AMPICS")
                .Formula( = "=ROUND(OCF_Activity_Template_Table[@[BNS Amount to Clear]],2)")
            Case }else{
                .Formula( = "=ROUND(ABS(OCF_Activity_Template_Table[@[BNS Amount to Clear]]),2)")
                .NumberFormat( = ACCOUNTING_NUMBER_FORMAT)
                .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
         Select
     }
    
    // Add Text to clearing entries
    with( rngClearingHeaders.Find(what=CLEARING_ENTRY_HEADERS(9), lookat=xlWhole).Offset(1, 0)) {
        .Activate()
         rngTmpRange = Range(rngHeaders.Find(GL_ACCOUNT_HEADER).Address)
        .Formula = "=" + CLEARING_ENTRY_TEXT + "&OCF_Activity_Template_Table[@[GL_Other]]"
        .AutoFill( Destination=Range(.Address, .Offset(lngLastRow - 2, 0).Address))
     }
    
    Application(.StatusBar = "Adjusting column widths and zoom ...")
    // Autofit column width
    rngHeaders(.CurrentRegion.Columns.AutoFit)
    rngClearingHeaders(.CurrentRegion.Columns.AutoFit)
    
    
    
    
    // Zoom out
    ActiveWindow(.Zoom = 85)
    
    Application.StatusBar = false
    Application.ScreenUpdating = true
    Application(.Calculation = xlCalculationAutomatic)
    
    ModifyActivityTemplateData( = 0)
     function{
    
HANDLE_PC_ERROR()
    MsgBox prompt="AN ERROR HAS OCCURRED!" + vbCrLf + vbCrLf + _
        "Did you remember to input profit centers for your legal entity?  Have you selected the correct ERP on the //Input + Assumptions// tab?"
    ModifyActivityTemplateData( = 1)
     function{
 }

//TODO  remove when done testing
function aa_tmpTestAllAccounts(){
    Call( InitializeTemplateGlobals)
    Call( ModifyAllMappingFile)
    Call( PivotAMFData)
    Call( CopyAMFPivotData)
    Call( PivotAllAccounts)
    Call( CopyAllAccountsPivot)
 }

function aa_PrepareOCFActivityTemplate(){
    Call( InitializeTemplateGlobals)
    if( CheckWorksheetExists(OCF_ACTIVITY_TEMPLATE_SHEET_NAME) ) {
        MsgBox( prompt="Please RESET the OCF Template before running the PREPARE functionality.")
         break}
    else if( ConfirmPrepOCFTemplate ) {
        //Call InitializeTemplateGlobals
        Call( ModifyAllMappingFile)
        Call( PivotAMFData)
        Call( CopyAMFPivotData)
        // 6/30/2015 -- v1.33 -- added workflow for all-accounts issue
        Call( PivotAllAccounts)
        Call( CopyAllAccountsPivot)
        if( hasActivityContraBNSPC ) {
            Call( CopyContraBNSPCPivotData)
         }
        
        if( ModifyActivityTemplateData == 0 ) {
            Call( LockOCFActivityTemplate)
            MsgBox( ("OCF Activity Template ready for analysis!"))
        }else{
            MsgBox( prompt="An error has occurred!  Please reset the OCF Template and begin again.")
             break}
         }
    }else{
        MsgBox prompt="Please ensure that the following two files have been imported" + vbCrLf + _
            "1.)  Pre-close All Mapping file for month-end " + _
                Format(Month(ActiveWorkbook.Names("admin_Pre_Close_AMF_Date").RefersToRange.Value), "mmmm") + " " + _
                Year(ActiveWorkbook.Names("admin_Pre_Close_AMF_Date").RefersToRange.Value) + vbCrLf + _
            "2(.)  Current month All Mapping file")
     }
 }

function CheckWorksheetExists(strWorksheetName  )  Boolean{
    On Error Resume }
    CheckWorksheetExists = Worksheets(strWorksheetName).Name != 0
 }

function ConfirmPrepOCFTemplate()  Boolean{
    if( (IsEmpty(ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange) || IsEmpty(ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange)) ) {
        ConfirmPrepOCFTemplate = false
    }else{
        ConfirmPrepOCFTemplate = true
     }
 }

function zz_ResetOCFActivityTemplate(){
    Call( InitializeTemplateGlobals)
    var wsWorksheetCheck  Worksheet
    if( MsgBox("This will completely reset (DELETE) the OCF Template + generated data." + vbCrLf + _
        "ARE YOU SURE YOU WANT TO DELETE ALL OCF TEMPLATE DATA?", vbYesNo) = vbNo ) {
         break}
    }else{
        Application.DisplayAlerts = false
        for ( var wsWorksheetCheck in Worksheets) {
            if( ! IsInArray(wsWorksheetCheck.Name, DO_NOT_DELETE_SHEETS) ) {
                Worksheets((wsWorksheetCheck.Name).Delete)
             }
        } wsWorksheetCheck
        
        // Reset pre-close and current-month admin values
        ActiveWorkbook(.Names("admin_Pre_Close_AMF_Tab").RefersToRange.ClearContents)
        ActiveWorkbook(.Names("admin_Current_Month_AMF_Tab").RefersToRange.ClearContents)
        
        Application.DisplayAlerts = true
        MsgBox( ("OCF Template reset!"))
     }
    
    Application.StatusBar = false
 }
function RunExtractJournalEntries(){
    Call( InitializeTemplateGlobals)
    if( CheckWorksheetExists(JE_VOUCHER_SHEET_NAME) ) {
       MsgBox( prompt="Please RESET the OCF Template or delete the Journal Entries tab before running EXTRACT Journal Entries!")
        break}
    else if( ! CheckWorksheetExists(OCF_ACTIVITY_TEMPLATE_SHEET_NAME) ) {
        MsgBox( prompt="Please IMPORT the required files and run the OCF analysis before running EXTRACT Journal Entries!")
         break}
    }else{
        Call( ExtractJournalEntries)
     }
 }

function ExtractJournalEntries(){
    var tblOCFTable, tblJETable  ListObject
    var lstTableField  ListColumn
    var wsWorksheet  Worksheet
    var curDTDFAmount  Currency
    var rngBNSAdjRange  Range
    var strDTDFAcct, strDTDFPC  
    var intCounter  Integer
    
    Call( InitializeTemplateGlobals)
    Call( UnlockOCFActivityTemplate)
    
    // Add Journal Entries worksheet
     wsWorksheet = Worksheets.Add(before=Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME))
    wsWorksheet(.Name = JE_VOUCHER_SHEET_NAME)
    
    // Activate journal entries table in OCF activity template
     tblJETable = Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(CLEARING_JE_TABLE_NAME)
    
    // Filter to exclude zero-dollar journal entries
    tblJETable(.Range.AutoFilter field=Application.WorksheetFunction.Match("Amount", tblJETable.HeaderRowRange, 0), Criteria1="<>0")
    
    // Copy filtered records to newly-created Journal Entries worksheets
    tblJETable(.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy)
    with( Worksheets(JE_VOUCHER_SHEET_NAME).Range("A1")) {
        .PasteSpecial( xlPasteFormats)
        .PasteSpecial( xlPasteValues)
     }
    
     tblOCFTable = Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME)
    
     rngBNSAdjRange = tblOCFTable.ListColumns(WorksheetFunction.Match(BNS_AMOUNT_TO_CLEAR_HEADER, tblOCFTable.HeaderRowRange, 0)).DataBodyRange
    curDTDFAmount( = WorksheetFunction.Sum(rngBNSAdjRange) * -1)
    strDTDFAcct( = ActiveWorkbook.Names("input_DTDF_Acct").RefersToRange.Value)
    strDTDFPC( = ActiveWorkbook.Names("input_DTDF_PC").RefersToRange.Value)
    

    
    // Add Due-To / Due-From account entry
    // HARDCODED LOGIC -- MUST BE UPDATED FOR CHANGES TO JE VOUCHER
    Worksheets((JE_VOUCHER_SHEET_NAME).Activate)
    Worksheets(JE_VOUCHER_SHEET_NAME).Range("A1").(xlDown).Offset(1, 0).Activate
    // Copy Posting Date
    ActiveCell(.Offset(-1, 0).Copy Destination=ActiveCell)
    ActiveCell.Offset(0, 1).Activate //Select cell in next column
    // Copy Document Date
    ActiveCell(.Offset(-1, 0).Copy Destination=ActiveCell)
    ActiveCell.Offset(0, 1).Activate //Select cell in next column
    //  Posting Key
    if( curDTDFAmount < 0 ) {
        ActiveCell.Value = "//" + GL_CREDIT_CODE
    }else{
        ActiveCell.Value = "//" + GL_DEBIT_CODE
     }
    
    ActiveCell.Offset(0, 1).Activate //Select cell in next column
    //  DTDF Account
    ActiveCell.Value = "//" + strDTDFAcct
    ActiveCell.Offset(0, 1).Activate //Select cell in next column
    //  DTDF PC
    ActiveCell(.Value = strDTDFPC)
    ActiveCell.Offset(0, 3).Activate //Select Amount field
    //  DTDF posting amount
    Select( Case ActiveWorkbook.Names("input_ERP_name").RefersToRange.Value)
        Case( "AMPICS")
            ActiveCell(.Value = curDTDFAmount)
        Case }else{
            ActiveCell(.Value = Abs(curDTDFAmount))
     Select
    ActiveCell(.NumberFormat = ACCOUNTING_NUMBER_FORMAT)
    ActiveCell.Offset(0, 1).Activate //Select cell in next column
    ActiveCell(.Value = "Clear BNS Activity")
    
    //ActiveCell.(xlToLeft).Offset(-1, 0).Activate
    Range("A" + ActiveCell.Row).Offset(-1, 0).Activate
    Range(ActiveCell.Address, ActiveCell.(xlToRight).Address).Select
    Selection(.Copy)
    Selection(.Offset(1, 0).Select)
    Selection(.PasteSpecial xlPasteFormats)
    
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(CLEARING_JE_TABLE_NAME).AutoFilter.ShowAllData)
    
    Call( LockOCFActivityTemplate)
    
    Worksheets((JE_VOUCHER_SHEET_NAME).Activate)
    
    // Auto-fit columns
    Worksheets((JE_VOUCHER_SHEET_NAME).Range("A1").CurrentRegion.Columns.AutoFit)
    Worksheets((JE_VOUCHER_SHEET_NAME).Range("A1").Activate)
    
    Worksheets(JE_VOUCHER_SHEET_NAME).Protect AllowFiltering=true, Password=WORKSHEET_PASSWORD
    
    Worksheets((JE_VOUCHER_SHEET_NAME).Tab.ColorIndex = JE_VOUCHER_TAB_COLOR)
    
    // Zoom out
    ActiveWindow(.Zoom = 85)
 }

function LockOCFActivityTemplate(){
    var rngHeaders  Range
    var lngLastRow, lngCurrentRow, lngOCFActivityCol, lngIsClearingAcctCol  
    
    Application.ScreenUpdating = false
    Application.DisplayStatusBar = true
    
    Call( InitializeTemplateGlobals)
    
     rngHeaders = Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME).HeaderRowRange
    rngHeaders.Find(BNS_YTD_BALANCE_HEADER).EntireColumn.Locked = false
    
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Activate)
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Range("a1").Activate)
    
    lngLastRow = ActiveCell.(xlDown).Row
    lngOCFActivityCol( = rngHeaders.Find(OCF_ACTIVITY_HEADER).Column)
    lngIsClearingAcctCol( = rngHeaders.Find(IS_CLEARING_ACCOUNT_HEADER).Column)
    
    rngHeaders(.Find(BNS_YTD_BALANCE_HEADER).Activate)
    ActiveCell.Locked = true
    ActiveCell(.Offset(1, 0).Activate)
    
    Application(.StatusBar = "Locking-down OCF Activity Template ...")
    
    //Lock-down "YTD BNS Balance" field for excluded, Direct and clearing accounts
    for( lngCurrentRow = ActiveCell.Row To lngLastRow) {
        if( (Cells(lngCurrentRow, lngOCFActivityCol).Value == "Direct" || _
            Cells(lngCurrentRow, lngOCFActivityCol).Value = OCF_ACTIVITY_EXCLUDE1 || _
            Cells(lngCurrentRow, lngIsClearingAcctCol).Value = true) ) {
            ActiveCell.Locked = true
         }
        ActiveCell(.Offset(1, 0).Activate)
    } lngCurrentRow
    
    Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Protect AllowFiltering=true, Password=WORKSHEET_PASSWORD
    
    Application.ScreenUpdating = true
    Application.StatusBar = false
 }


function UnlockOCFActivityTemplate(){
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Unprotect Password=WORKSHEET_PASSWORD)
 }

function DisplaySelectAMFPeriodForm(){
    SelectAMFPeriodForm(.Show)
 }

function OverwriteAMFData( bPreCloseAMF  Boolean = false)  Boolean{
    var wsWorksheet  Worksheet
    var strName  
    
    if( bPreCloseAMF ) { //trying to import pre-close AMF
        if( ! IsEmpty(ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange) ) {
            //pre-close tab exists
            //prompt that they//ll overwrite
            if( MsgBox("Importing another pre-close All Mapping File will completely overwrite (DELETE) the current data.  Are you sure you want to do this?", vbYesNo) == vbNo ) {
                MsgBox( ("Pre-close All Mapping File import CANCELLED!"))
                OverwriteAMFData = false
                 function{
            }else{
                OverwriteAMFData = true
                 function{
            //import the file + process it
            //exit sub
             }
        }else{ //No pre-close AMF presently imported
            OverwriteAMFData = false
             function{
         }
    }else{ //trying to import current AMF -- i.e. bPreCloseAMF == false
        if( ! IsEmpty(ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange) ) {
            //current tab exists
            //prompt that they//ll overwrite
            if( MsgBox("Importing another current month All Mapping File will completely overwrite (DELETE) the current data.  Are you sure you want to do this?", vbYesNo) == vbNo ) {
                MsgBox( ("All Mapping File import CANCELLED!"))
                OverwriteAMFData = false
                 function{
            }else{
            //delete all current month tabs (...)
                OverwriteAMFData = true
                 function{
            //import the file
            //exit sub
             }
         }
     }
    OverwriteAMFData = false
 }

function WriteAdminRange(rngRange  Range, strValue  ){
    //TODO  stuff ...
 }

function ResetPreCloseData(){
    var PRE_DO_NOT_DELETE_SHEETS(1 To 10)  
    var wsWorksheetCheck  Worksheet
    
    Call( InitializeTemplateGlobals)
    // Populate PRE_DO_NOT_DELETE_SHEETS array
    PRE_DO_NOT_DELETE_SHEETS((1) = "TODO")
    PRE_DO_NOT_DELETE_SHEETS((2) = "ADMIN")
    PRE_DO_NOT_DELETE_SHEETS(3) = "Input + Assumptions"
    PRE_DO_NOT_DELETE_SHEETS((4) = "Data Dictionary")
    PRE_DO_NOT_DELETE_SHEETS((5) = "HFM Acct - OCF Activity")
    PRE_DO_NOT_DELETE_SHEETS((6) = BNS_PROFIT_CENTERS_SHEET_NAME)
    PRE_DO_NOT_DELETE_SHEETS((7) = "GL Acct - Clearing")
    PRE_DO_NOT_DELETE_SHEETS((8) = "Cost Centers")
    PRE_DO_NOT_DELETE_SHEETS((9) = PERSISTENT_STORAGE_SHEET)
    PRE_DO_NOT_DELETE_SHEETS((10) = ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange.Value)
    
    //Keep all internals and the current month AMF, if exists
    Application.DisplayAlerts = false
    for ( var wsWorksheetCheck in Worksheets) {
        if( ! IsInArray(wsWorksheetCheck.Name, PRE_DO_NOT_DELETE_SHEETS) ) {
            Worksheets((wsWorksheetCheck.Name).Delete)
         }
    } wsWorksheetCheck
    ActiveWorkbook(.Names("admin_Pre_Close_AMF_Tab").RefersToRange.ClearContents)
    Application.DisplayAlerts = true
 }

//function ResetCurrentMonthData(){
//    Call InitializeTemplateGlobals
//
//    On Error Resume }
//    Application.DisplayAlerts = false
//    Worksheets(ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange.Value).Delete
//    Worksheets(MODIFIED_AMF_SHEET_NAME).Delete
//    Worksheets(AMF_PIVOT_SHEET_NAME).Delete
//    Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Delete
//    Worksheets(CONTRA_BNS_PC_PIVOT_SHEET_NAME).Delete
//    Worksheets(CONTRA_BNS_PC_BAL_SHEET_NAME).Delete
//    Worksheets(JE_VOUCHER_SHEET_NAME).Delete
//    ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange.ClearContents
//    Application.DisplayAlerts = true
// }

function ResetCurrentMonthData(){
    var CUR_DO_NOT_DELETE_SHEETS(1 To 13)  
    var wsWorksheetCheck  Worksheet
    
    Call( InitializeTemplateGlobals)
    
    CUR_DO_NOT_DELETE_SHEETS((1) = "TODO")
    CUR_DO_NOT_DELETE_SHEETS((2) = "ADMIN")
    CUR_DO_NOT_DELETE_SHEETS(3) = "Input + Assumptions"
    CUR_DO_NOT_DELETE_SHEETS((4) = "Data Dictionary")
    CUR_DO_NOT_DELETE_SHEETS((5) = "HFM Acct - OCF Activity")
    CUR_DO_NOT_DELETE_SHEETS((6) = BNS_PROFIT_CENTERS_SHEET_NAME)
    CUR_DO_NOT_DELETE_SHEETS((7) = "GL Acct - Clearing")
    CUR_DO_NOT_DELETE_SHEETS((8) = "Cost Centers")
    CUR_DO_NOT_DELETE_SHEETS((9) = PERSISTENT_STORAGE_SHEET)
    // do not delete pre-close sheets
    // pre-close amf
    CUR_DO_NOT_DELETE_SHEETS((10) = ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange.Value)
    // modded pre-close amf
    CUR_DO_NOT_DELETE_SHEETS(11) = "Pre-Close " + ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange.Value
    // pre-close amf pivot
    CUR_DO_NOT_DELETE_SHEETS((12) = PRE_CLOSE_AMF_PIVOT_SHEET_NAME)
    // pre-close bns income
    CUR_DO_NOT_DELETE_SHEETS((13) = PRE_CLOSE_BNS_INCOME_SHEET_NAME)
    
    Application.DisplayAlerts = false
    for ( var wsWorksheetCheck in Worksheets) {
        if( ! IsInArray(wsWorksheetCheck.Name, CUR_DO_NOT_DELETE_SHEETS) ) {
            Worksheets((wsWorksheetCheck.Name).Delete)
         }
    } wsWorksheetCheck
    ActiveWorkbook(.Names("admin_Current_Month_AMF_Tab").RefersToRange.ClearContents)
    Application.DisplayAlerts = true
 }

// function required to deal with regionalization and parsing English file dates{
function MonthNameToNumber(ByVal strMonthName  )  {
    var strMonthNum  
    Select( Case strMonthName)
        Case( "Jan")
            strMonthNum( = "01")
        Case( "Feb")
            strMonthNum( = "02")
        Case( "Mar")
            strMonthNum( = "03")
        Case( "Apr")
            strMonthNum( = "04")
        Case( "May")
            strMonthNum( = "05")
        Case( "Jun")
            strMonthNum( = "06")
        Case( "Jul")
            strMonthNum( = "07")
        Case( "Aug")
            strMonthNum( = "08")
        Case( "Sep")
            strMonthNum( = "09")
        Case( "Oct")
            strMonthNum( = "10")
        Case( "Nov")
            strMonthNum( = "11")
        Case( "Dec")
            strMonthNum( = "12")
     Select
    MonthNameToNumber( = strMonthNum)
 }

function ImportExternalAMF( bPreCloseAMF  Boolean = false){
    var fdDialog  FileDialog
    var wbAMFWorkbook, wbOCFWorkbook  Workbook
    var strEntityNumber, strMonth, strYear, strAMFDate, strAdminDate, strAdminCheck  
    var AMF_SHEET_NAME  
    
    // Initialize sheet name for AMF data
    AMF_SHEET_NAME( = "Sheet1")
     wbOCFWorkbook = Application.ThisWorkbook
    
    // File dialog for user to select All Mapping File
     fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.AllowMultiSelect = false
    fdDialog(.Title = "Please select the All Mapping File for your legal entity")
    fdDialog(.Filters.Clear)
    fdDialog(.Filters.Add Description="Excel files", Extensions="*.xls*")
    fdDialog(.InitialFileName = ActiveWorkbook.Path)
    
    // fdDialog.Show = true when a file is selected in the dialog box, otherwise false if cancelled
    if( fdDialog.Show ) {
        // Instantiate workbook variable + open AMF for data extraction
         wbAMFWorkbook = Workbooks.Open(fdDialog.SelectedItems(1))
    }else{ //Cancel was pressed
         break}
     }
    
    wbAMFWorkbook(.Activate)
    
    // Extract legal entity number from string in All Mapping File
    strEntityNumber( = Right(Left(wbAMFWorkbook.Worksheets(AMF_SHEET_NAME).Range("A1").Value, 5), 4))
    
    // Extract date from string in AMF
    strMonth( = Mid(wbAMFWorkbook.Worksheets(AMF_SHEET_NAME).Range("A1").Value, _)
        Application(.WorksheetFunction.Find("_", wbAMFWorkbook.Worksheets(AMF_SHEET_NAME).Range("A1").Value) + 1, 3))
        
    strYear( = Right(Mid(wbAMFWorkbook.Worksheets(AMF_SHEET_NAME).Range("A1").Value, _)
        Application(.WorksheetFunction.Find("-", wbAMFWorkbook.Worksheets(AMF_SHEET_NAME).Range("A1").Value) + 1, 4), 2))
    
    // Map English month abbreviation to month number
    strMonth( = MonthNameToNumber(strMonthName=strMonth))
    strYear = "20" + strYear
    strAMFDate = strYear + "/" + strMonth
    
    // Pull admin-input pre-close date
    strAdminDate( = wbOCFWorkbook.Names("admin_Pre_Close_AMF_Date").RefersToRange.Value)

    if( bPreCloseAMF ) {
        // Check AMF date against admin input for pre-close date
        // if( consistent, continue
        // }else{, clean-up and call ImportAMF again and pass parm
        //  sub
        strAdminCheck( = strAdminDate)
        strAdminCheck = Year(strAdminDate) + "/" + Format(Month(strAdminDate), "00")
        
        if( ! (strAMFDate == strAdminCheck) ) {
            MsgBox prompt="Pre-close All Mapping File must be for month ending " + _
                Format(strAdminDate, "mmmm") + " " + Year(strAdminDate) + _
                "(.  Please import the correct pre-close All Mapping File.")
            wbAMFWorkbook(.Close)
             break}
         }
    }else{
        strAdminCheck( = CDate(strAMFDate))
        if( ! CDate(strAdminCheck) > CDate(strAdminDate) ) {
            MsgBox prompt="Current month All Mapping File must be dated after month ending " + _
                Format(strAdminDate, "mmmm") + " " + Year(strAdminDate) + _
                "(.  Please import a valid current month All Mapping File.")
            wbAMFWorkbook(.Close)
             break}
         }
     }
    
    AMF_RAW_DATA_SHEET_NAME = strEntityNumber + " - " + "AMF - " + Year(strAMFDate) + Format(Month(strAMFDate), "00")
    
    wbAMFWorkbook(.Worksheets(AMF_SHEET_NAME).Copy before=wbOCFWorkbook.Worksheets("PersistentStorage"))
    ActiveSheet(.Name = AMF_RAW_DATA_SHEET_NAME)
    ActiveSheet(.Tab.Color = vbBlack)
    
    SetPersistentVariable( strVariable="AMF_RAW_DATA_SHEET_NAME", strValue=AMF_RAW_DATA_SHEET_NAME, strDescription="All Mapping File raw data worksheet name")
    
    // Close AMF file
    wbAMFWorkbook(.Close)
    
    // Protect raw data tab
    Worksheets((AMF_RAW_DATA_SHEET_NAME).Activate)
    Worksheets((AMF_RAW_DATA_SHEET_NAME).Range("A1").Activate)
    Worksheets((AMF_RAW_DATA_SHEET_NAME).Protect Password=WORKSHEET_PASSWORD)
    
    if( bPreCloseAMF ) {
        ActiveWorkbook(.Names("admin_Pre_Close_AMF_Tab").RefersToRange.Value = AMF_RAW_DATA_SHEET_NAME)
        Call( aa_PreparePreCloseAMFData)
    }else{
        ActiveWorkbook(.Names("admin_Current_Month_AMF_Tab").RefersToRange.Value = AMF_RAW_DATA_SHEET_NAME)
     }
    
    // Focus on inputs tab
    Worksheets("Input + Assumptions").Activate
    ActiveWorkbook(.Names("input_Contra_BNS_PC").RefersToRange.Activate)
    
    MsgBox( prompt="All Mapping File succesfully imported!")
        
 }

function zzz_OLD_ValidatePostOCFAMF(){
    var wbOCFWorkbook, wbValWorkbook  Workbook
    var wsWorksheet  Worksheet
    var strValidationSheetName  
    var strCopySheets(1 To 5)  
    var strVBIDERef, strTmpExport, strMappingFormula  
    // Classes below are part of VBA Extensibility object library
    var refReference  Reference
    var comVBAModule  VBComponent
    
    strCopySheets((1) = "ADMIN")
    strCopySheets((2) = "HFM Acct - OCF Activity")
    strCopySheets((3) = "BNS Profit Centers")
    strCopySheets((4) = "GL Acct - Clearing")
    strCopySheets((5) = "Cost Centers")
    
    strVBIDERef( = "Microsoft Visual Basic for Applications Extensibility 5.3")
    
    strMappingFormula = "VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]]," + _
        "CHOOSE({1,2,3},OCF_Activity_Template_Table[HFM_Account]&OCF_Activity_Template_Table[GL_Other]&OCF_Activity_Template_Table[Trading_Partner]," + _
        "OCF_Activity_Template_Table[Total( BNS Balance]),2,FALSE)")
        
    strReplaceMapping( = wbOCFWorkbook)
    
     wbOCFWorkbook = ThisWorkbook
     wbValWorkbook = Workbooks.Add
    wbValWorkbook(.Title = "OCF Template Validation - Temporary Workbook")
    
    //wbWorkbook.VBProject.VBComponents.
    // Add VBIDE object library to project for import / export of OCF Template codebase
    // "Microsoft Visual Basic for Applications Extensibility 5.3"
    On Error Resume }
    ThisWorkbook(.VBProject.References.AddFromGuid GUID="{0002E157-0000-0000-C000-000000000046}", Major=5, Minor=3)
    On( Error GoTo 0)
    MsgBox( prompt="Reference added!")
    
    // Copy OCF_Template module from OCF Template to temporary validation workbook
    //wbValWorkbook.VBProject.VBComponents.Import wbOCFWorkbook.VBProject.VBComponents("OCF_Template").Export
    strTmpExport = Environ("Temp") + "\tmp_OCF_Template.bas"
    wbOCFWorkbook(.VBProject.VBComponents("OCF_Template").Export strTmpExport)
    wbValWorkbook(.VBProject.VBComponents.Import strTmpExport)
    
    // Copy SelectAMFPeriodForm from OCF Template to temporary validation workbook
    //wbValWorkbook.VBProject.VBComponents.Import wbOCFWorkbook.VBProject.VBComponents("OCF_Template").Export
    strTmpExport = Environ("Temp") + "\tmp_OCF_Template.bas"
    wbOCFWorkbook(.VBProject.VBComponents("SelectAMFPeriodForm").Export strTmpExport)
    wbValWorkbook(.VBProject.VBComponents.Import strTmpExport)
    
    //TODO  revise subroutines below for error handling / function calls
    Application.Run wbValWorkbook.Name + "!DisplaySelectAMFPeriodForm"
    Application.Run wbValWorkbook.Name + "!DisplaySelectAMFPeriodForm"
    Application.Run wbValWorkbook.Name + "!aa_PrepareOCFActivityTemplate"
       
    //=//[06162015 NEW OCF Activity Template - SAP Entities v1.32.xlsm]OCF Activity Template//!OCF_Activity_Template_Table[@[Total BNS Balance]]
    
        
    
    for ( var refReference in ThisWorkbook.VBProject.References) {
        // TODO consider removing description reference and use GUID
        if( refReference.Description == strVBIDERef ) {
            ThisWorkbook(.VBProject.References.Remove refReference)
            MsgBox( prompt="Reference removed!")
         }
    } refReference
    
//   ? Create new workbook
//   ? Copy the codebase out into a new workbook
//   ? Copy the requisite tabs into the New Workbook
//       § ADMIN
//       § HFM Acct - OCF Activity
//       § BNS Profit Centers
//       § GL Acct - Clearing
//       § Cost Centers
//   ? if( pre-close tab exists, copy it into the new workbook
//       § Update new workbook//s "admin_Pre_Close_AMF_Tab"
//   ? Prompt user to input post-AMF
//   ? Prepare OCF Activity Template
//   ? Rename OCFAT -> VALIDATION
//   ? Copy Tab back into Template
//   ? Map-in entries from OCFAT
//   ? Close and delete new workbook
 }

function ZeroesBasedValidation(){
    var wbOCFWorkbook, wbValWorkbook  Workbook
    var wsOCFATSheet, wsValSheet, wsCheckSheet  Worksheet
    var strOCFATIDString, strMappingFormula, strTmpFormula, strTmpFormula2, strTempString  
    var fdDialog  FileDialog
    var rngHeaders  Range
    var lngLastRow  
    
    Call( InitializeTemplateGlobals)
    
     wbValWorkbook = ThisWorkbook
    
    if( ! CheckWorksheetExists(OCF_ACTIVITY_TEMPLATE_SHEET_NAME) ) {
        MsgBox prompt="Please import the pre-close and current month //post// All Mapping Files and click //PREPARE// before running zeroes-based validation analysis."
         break}
     }
    
    // File dialog for user to select All Mapping File
     fdDialog = Application.FileDialog(msoFileDialogFilePicker)
    fdDialog.AllowMultiSelect = false
    fdDialog(.Title = "Please select the OCF Template you have completed and wish to validate")
    fdDialog(.Filters.Clear)
    fdDialog(.Filters.Add Description="Excel files", Extensions="*.xls*")
    fdDialog(.InitialFileName = ActiveWorkbook.Path)
    
    // fdDialog.Show = true when a file is selected in the dialog box, otherwise false if cancelled
    if( fdDialog.Show ) {
        // Instantiate workbook variable + open AMF for data extraction
         wbOCFWorkbook = Workbooks.Open(fdDialog.SelectedItems(1))
    }else{ //Cancel was pressed
         break}
     }
    
    Application.ScreenUpdating = false
    Application.DisplayAlerts = false
    
    On Error Resume }
     wsCheckSheet = wbOCFWorkbook.Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME)
    On( Error GoTo 0)
    
    if( wsCheckSheet Is Nothing ) {
        MsgBox( prompt="Please select a valid, completed OCF Template file!")
         break}
     }
    
    strOCFATIDString( = wbOCFWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange.Value)
    strOCFATIDString = Left(strOCFATIDString, 4) + " - " + Right(strOCFATIDString, 6)
    
     wsOCFATSheet = wbValWorkbook.Worksheets.Add(after=Worksheets("Cost Centers"))
    wsOCFATSheet.Name = strOCFATIDString + " - OCFAT"
    wsOCFATSheet(.Tab.Color = vbBlack)
     wsValSheet = wbValWorkbook.Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME)
    wsValSheet(.Unprotect Password=WORKSHEET_PASSWORD)
    
    // Copy + paste the OCF Activity Template sheet from the completed Template
//    wbOCFWorkbook.Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Copy
//    wsValSheet.PasteSpecial xlPasteFormats
//    wsValSheet.PasteSpecial xlPasteValues
//    // Take 2
//    wbOCFWorkbook.Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Cells.Copy
//    wsOCFATSheet.Cells.PasteSpecial xlPasteFormats
//    wsOCFATSheet.Cells.PasteSpecial xlPasteValues
    // Take 3
    wbOCFWorkbook(.Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(CLEARING_JE_TABLE_NAME).AutoFilter.ShowAllData)
    //Range(wbOCFWorkbook.Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME).Range.Address).Activate
    wbOCFWorkbook(.Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME).Range.Copy)
    
    wsOCFATSheet(.Range("a1").PasteSpecial xlPasteFormats)
    wsOCFATSheet(.Range("a1").PasteSpecial xlPasteValues)
    
    // Add Excel table for completed OCFAT
    wsOCFATSheet(.Activate)
    wsOCFATSheet(.Range("A1").CurrentRegion.Activate)
    ActiveSheet(.ListObjects.Add(SourceType=xlSrcRange, Source=Selection, XlListObjectHasHeaders=xlYes, TableStyleName=TABLE_STYLE_FORMAT).Name = ZBV_OCFAT_TABLE_NAME)
    
    // Rename OCF Activity Template in validation workbook
    wsValSheet(.Name = ZBV_POST_OCFAT_SHEET_NAME)
    
     rngHeaders = wsValSheet.ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME).HeaderRowRange
    
    
    wsValSheet(.Activate)
    lngLastRow = wsValSheet.Range(rngHeaders.Item(1).Address).(xlDown).Row
    wsValSheet(.ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME).HeaderRowRange.Item(1).Activate)
    // Map-in balances from OCFAT
    rngHeaders(.Find(what=BNS_YTD_BALANCE_HEADER, lookat=xlWhole).Offset(1, 0).Activate)
    //MsgBox prompt=ActiveCell.Address
    
    with( ActiveCell) {
        strTmpFormula( = "=VLOOKUP([@[HFM_Account]]&[@[GL_Other]]&[@[Trading_Partner]],""XXX"",2,FALSE)")
        strTmpFormula2 = "CHOOSE({1,2,3}," + ZBV_OCFAT_TABLE_NAME + "[HFM_Account]&" + ZBV_OCFAT_TABLE_NAME + "[GL_Other]&" + ZBV_OCFAT_TABLE_NAME + "[Trading_Partner]," _
            + ZBV_OCFAT_TABLE_NAME + "[" + BNS_YTD_BALANCE_HEADER + "])"
        .FormulaArray( = strTmpFormula)
        strTempString( = """XXX""")
        .Replace( what=strTempString, replacement=strTmpFormula2, lookat=xlPart)
     }
    ActiveCell(.AutoFill Destination=Range(ActiveCell.Address, ActiveCell.Offset(lngLastRow - 2, 0).Address))
    
    // Close OCF Template
    wbOCFWorkbook(.Close)
    
    Application.ScreenUpdating = true
    Application.DisplayAlerts = true
    
    MsgBox( prompt="OCF Template ready for zeroes-based validation analysis!")
    
 }

// Subroutine for modifying balance in a clearing account directly
function ZeroClearingAccountBal(){
    // TODO  Add error handling for button presses before OCF Template prepared
    // TODO  Hide this until the OCF template is prepared.
    
    var rngClearingAcct, rngHeaders, rngJEHeaders, rngBNSAmountToClear, rngGLAcct  Range
    var bIsClearingAcct  Boolean
    var strPrompt, strGLAcct, strClearingAcct, strBal, strGLAcctDesc, strFormat  

    Call( InitializeTemplateGlobals)
    
    if( ! CheckWorksheetExists(OCF_ACTIVITY_TEMPLATE_SHEET_NAME) ) {
        MsgBox( prompt="You cannot flatten clearing account balances before preparing the OCF Activity Template!" _)
            + " Please prepare the OCF Activity Template and try again!"
         break}
     }

    strFormat( = "#,##0.00")
    
//    • User clicks button
//    • Prompted to click on row / cell they wish to change
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Activate)
    strPrompt( = "Please select the row containing the clearing account balance you wish to modify")
    
    On( Error GoTo NO_RANGE_SELECTED)
     rngClearingAcct = Application.InputBox(prompt=strPrompt, Type=8)
    On( Error GoTo 0)
    
//    • Template checks to confirm that the account in that row is a clearing account
//    • Template prompts them to confirm that they wish to modify the $XXX,XXX balance for GL Account ##### which clears balances from GL Account ###### <GL Acct Desc>
    // Get header row of OCF Activity Template
     rngHeaders = Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(OCF_ACTIVITY_TEMPLATE_TABLE_NAME).HeaderRowRange
     rngJEHeaders = Worksheets(OCF_ACTIVITY_TEMPLATE_SHEET_NAME).ListObjects(CLEARING_JE_TABLE_NAME).HeaderRowRange
    
    bIsClearingAcct( = rngHeaders.Find(what=IS_CLEARING_ACCOUNT_HEADER, lookat=xlWhole).Offset(rngClearingAcct.Row - 1, 0).Value)
    if( ! bIsClearingAcct ) {
        MsgBox( prompt="You have NOT selected a clearing account.  Please review and reselect.")
         break}
     }
    
    //strClearingAcct = rngJEHeaders.Find(what="Account", lookat=xlWhole).Offset(rngClearingAcct.Row - 1, 0).Value
    strClearingAcct( = rngHeaders.Find(what=GL_ACCOUNT_HEADER, lookat=xlWhole).Offset(rngClearingAcct.Row - 1, 0).Value)
    strGLAcct( = Worksheets("GL Acct - Clearing").ListObjects(1).DataBodyRange.Find(what=strClearingAcct, lookat=xlWhole).Offset(0, -1).Value)
    strBal( = Format(rngHeaders.Find(what=CONTRA_BNS_PC_HEADER, lookat=xlWhole).Offset(rngClearingAcct.Row - 1, 0).Value, strFormat))
    //strGLAcctDesc = rngHeaders.Find(what=GL_ACCOUNT_DESCRIPTION_HEADER, lookat=xlWhole).Offset(rngClearingAcct.Row - 1, 0).Value
    
    //strPrompt = "Are you sure you want to modify the " + strBal + " balance in clearing account # " + strClearingAcct + _
        " which clears activity from GL account # " + strGLAcct + " - //" + strGLAcctDesc + "//?"
    strPrompt = "Are you sure you want to zero-out the " + strBal + " balance in clearing account # " + strClearingAcct + _
        " which clears activity from GL account # " + strGLAcct + "?"
            
    if( MsgBox(prompt==strPrompt, Buttons==vbYesNo) == vbNo ) {
         break}
     }
    
    Application.AutoCorrect.AutoFillFormulasInLists = false
    Application.DisplayAlerts = false
    On Error Resume }
    Worksheets((JE_VOUCHER_SHEET_NAME).Delete)
    On( Error GoTo 0)
    Application.DisplayAlerts = true
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Unprotect Password=WORKSHEET_PASSWORD)
    
//    • Template modifies formula for Col. P of this row to incorporate YTD balance reflected in Col. P
//    • Template prompts user that they should enter what the balance _should be_ for the clearing account in Col. O
        
     rngBNSAmountToClear = rngHeaders.Find(what=BNS_AMOUNT_TO_CLEAR_HEADER, lookat=xlWhole).Offset(rngClearingAcct.Row - 1, 0)
    // 7/30/2015 -- revised for single "Exclude" classifier
    //rngBNSAmountToClear.Formula = "=IF(OR([@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE1 + """, [@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE2 + """),0,IF([@[OCF_Activity]]=""Direct"",-[@[CS-Owned BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]],-[@[CS-Owned BNS Balance]]-[@[Total BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]]))"
    rngBNSAmountToClear.Formula = "=IF(OR([@[OCF_Activity]] = """ + OCF_ACTIVITY_EXCLUDE1 + """),0,IF([@[OCF_Activity]]=""Direct"",-[@[CS-Owned BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]],-[@[CS-Owned BNS Balance]]-[@[Total BNS Balance]]-[@[Contra-BNS PC]]-[@[Clearing Acct Balance]]))"
    // Must override OCF Activity field if account is Direct so that BNS Amount to Clear field picks-up adjustment
    
    rngBNSAmountToClear(.Offset(0, -1).Activate)
    with( ActiveCell.Validation) {
        .Delete()
        .Add( Type=xlValidateWholeNumber, _)
            AlertStyle(=xlValidAlertStop, _)
            Operator(=xlEqual, _)
            Formula1(="0")
            //.ErrorTitle = "Zero-out clearing account balance!"
            .ErrorMessage( = "You must enter zero for the balance for this clearing account!")
            .InputMessage( = "You must enter zero for the balance for this clearing account!")
     }
    ActiveCell(.Value = 0)
    
    with( rngHeaders.Find(what=OCF_ACTIVITY_HEADER, lookat=xlWhole).Offset(rngClearingAcct.Row - 1)) {
        .Value( = "Indirect")
        .Interior(.Color = vbYellow)
        .ClearComments()
        .AddComment( "OCF Activity indicator overridden for this Direct clearing account to allow BNS Amount to Clear calculation.")
     }
    
    rngBNSAmountToClear(.Activate)
    
    Worksheets((OCF_ACTIVITY_TEMPLATE_SHEET_NAME).Protect Password=WORKSHEET_PASSWORD)
    Application.AutoCorrect.AutoFillFormulasInLists = true
    
    MsgBox prompt="Clearing account " + strClearingAcct + " balance cleared to zero!"
    
NO_RANGE_SELECTED()
     break}
    
//    rngHeaders.Find(what=BNS_YTD_BALANCE_HEADER, lookat=xlWhole).Offset(rngClearingAcct.Row - 1, 0).Activate
//    ActiveCell.Interior.Color = vbYellow
//    ActiveCell.Value = "ENTER BALANCE!"
 }

function HideAdminWorksheets(){
    var strInputPassword  
    var wsWorksheet  Worksheet
    
    Call( InitializeTemplateGlobals)
    
    strInputPassword( = InputBox(prompt="Please input the password to hide worksheets"))
    if( strInputPassword != WORKSHEET_PASSWORD ) {
        MsgBox( prompt="You did not enter the correct password.  Please contact the OCF Project Team administrator.")
         break}
     }
    
    for ( var wsWorksheet in Worksheets) {
        if( IsInArray(wsWorksheet.Name, HIDDEN_WORKSHEETS) ) {
            wsWorksheet(.Visible = xlSheetVeryHidden)
         }
    } wsWorksheet
    
 }

function UnHideAdminWorksheets(){
    var strInputPassword  
    var wsWorksheet  Worksheet
    
    Call( InitializeTemplateGlobals)
    
    strInputPassword( = InputBox(prompt="Please input the password to hide worksheets"))
    if( strInputPassword != WORKSHEET_PASSWORD ) {
        MsgBox( prompt="You did not enter the correct password.  Please contact the OCF Project Team administrator.")
         break}
     }
    
    for ( var wsWorksheet in Worksheets) {
        if( IsInArray(wsWorksheet.Name, HIDDEN_WORKSHEETS) ) {
            wsWorksheet.Visible = true
         }
    } wsWorksheet
    
    Worksheets("Input + Assumptions").Activate
    ActiveWorkbook(.Names("input_SAP_Document_Date").RefersToRange.Activate)
    
    
 }
