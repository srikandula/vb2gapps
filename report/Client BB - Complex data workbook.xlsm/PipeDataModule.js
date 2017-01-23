Attribute VB_Name( = "PipeDataModule")
Option Explicit  // Causes error if variable is not defined with var
 var ERROR_YIELD   = 99999999
// Use to highlight rows a different color
// see function subHighlightRow()
 Enum ROW_HIGHLIGHT
    Highlight_Off( = 1)
    Highlight_RED( = 2)
 Enum
Global var MAX_SAVE_ARRAY  Integer = 3000
Global vSaveSTPRArray(MAX_SAVE_ARRAY, 15)  Variant
Global lSaveSTRPArrayHighNumber  
Global vSaveJobArray(MAX_SAVE_ARRAY)  Variant
Global lSaveJobArrayHighNumber  
 function ColorText(CellToColor  , Color  , Value  )  Boolean{

    var CellLength  Integer
    var StartPosition  Integer
    var ColorLength  Integer
    
    CellLength( = Len(Range(CellToColor).Value))
    ColorLength( = Len(Value))
    StartPosition( = CellLength - ColorLength)
    
    Range((CellToColor).Select)
    with( ActiveCell.Characters(Start=StartPosition, Length=ColorLength).Font) {
                            //.Name = "Arial"
                            //.FontStyle = "Bold"
                            //.Size = 8
                            //.Strikethrough = false
                            //.Superscript = false
                            //.Subscript = false
                            //.OutlineFont = false
                            //.Shadow = false
                            //.Underline = xlUnderlineStyleNone
                            if( Color == "Red" || Color == "red" ) {
                                .Color( = -16776961)
                            else if( Color = "Blue" || Color = "blue" ) {
                                .Color( = -100520)
                            else if( Color = "DarkBlue" || Color = "Darkblue" || Color = "darkblue" ) {
                                .Color( = RGB(0, 0, 255))
                            else if( Color = "Purple" || Color = "purple" ) {
                                .Color( = -6279056)
                            else if( Color = "DarkPurple" || "darkpurple" || "Darkpurple" ) {
                                .Color( = -9174924)
                             }
                            //.TintAndShade = 0
                            //.ThemeFont = xlThemeFontNone
     }
    ColorText = true
 }

 function SaveSTPRRowNumber(p_bInitialize  Boolean, p_lSaveRowNumber  , p_sTabName  ){
    var k  Integer
    if( p_bInitialize ) {
        // Intializeing Array ONLY
        lSaveSTRPArrayHighNumber( = 0)
        Debug(.Print "Initialize Array")
    }else{
        // Save Row number that was modified
        lSaveSTRPArrayHighNumber( = lSaveSTRPArrayHighNumber + 1)
        for( k = 1 To 11) {
            vSaveSTPRArray(lSaveSTRPArrayHighNumber, k) = Worksheets(p_sTabName).Range("E" + Chr(64 + k) + p_lSaveRowNumber).Value
        } k
     }
 }
 function SaveJobowNumber(p_bInitialize  Boolean, p_lSaveRowNumber  , p_sTabName  ){
    var k  Integer
    if( p_bInitialize ) {
        // Intializeing Array ONLY
        lSaveSTRPArrayHighNumber( = 0)
        Debug(.Print "Initialize Array")
    }else{
        // Save Row number that was modified
        lSaveSTRPArrayHighNumber( = lSaveSTRPArrayHighNumber + 1)
        for( k = 1 To 4) {
            vSaveSTPRArray(lSaveSTRPArrayHighNumber, k) = Worksheets(p_sTabName).Range("E" + Chr(75 + k) + p_lSaveRowNumber).Value
        } k
     }
 } // --------------------------------------------------
// Comments Get Yield from Tab //Tb2a2 DATA//
// --------------------------------------------------
function CalcYield_2A2(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 12        // Table starts at this row
   iNumColumns = 12  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("O12").Value)
   iPipeSize( = Worksheets(sTabName).Range("O13").Value)
   sSeamType( = Worksheets(sTabName).Range("O14").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 12, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("O15").Value = lYield)
 }
// --------------------------------------------------
// Comments Get Yield from Tab //Tbl_2b3 DATA//
// --------------------------------------------------
function CalcYield_2b3(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 34        // Table starts at this row
   iNumColumns = 11  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("N35").Value)
   iPipeSize( = Worksheets(sTabName).Range("N36").Value)
   sSeamType( = Worksheets(sTabName).Range("N37").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 11, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("N38").Value = lYield)
 }
// --------------------------------------------------
// Comments Get Yield from Tab //Tbl2b4_1 DATA//
// --------------------------------------------------
function CalcYield_2b4_1(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 51        // Table starts at this row
   iNumColumns = 6  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("I50").Value)
   iPipeSize( = Worksheets(sTabName).Range("I51").Value)
   sSeamType( = Worksheets(sTabName).Range("I52").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 6, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("I53").Value = lYield)
 }
// --------------------------------------------------
// Comments Get Yield from Tab //Tbl2b4_2 DATA//
// --------------------------------------------------
function CalcYield_2b4_2(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 58        // Table starts at this row
   iNumColumns = 4  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("G57").Value)
   iPipeSize( = Worksheets(sTabName).Range("G58").Value)
   sSeamType( = Worksheets(sTabName).Range("G59").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 4, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("G60").Value = lYield)
 }
// --------------------------------------------------
// Comments Get Yield from Tab //2b4_Unknown_1 DATA//
// --------------------------------------------------
function CalcYield_2b4_Unknown_1(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 67        // Table starts at this row
   iNumColumns = 11  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("N67").Value)
   iPipeSize( = Worksheets(sTabName).Range("N68").Value)
   sSeamType( = Worksheets(sTabName).Range("N69").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 11, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("N70").Value = lYield)
 }

// --------------------------------------------------
// Comments Get Yield from Tab //2b4_Unknown_2 DATA//
// --------------------------------------------------
function CalcYield_2b4_Unknown_2(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 76        // Table starts at this row
   iNumColumns = 12  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("O76").Value)
   iPipeSize( = Worksheets(sTabName).Range("O77").Value)
   sSeamType( = Worksheets(sTabName).Range("O78").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 12, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("O79").Value = lYield)
 }

// --------------------------------------------------
// Comments Get Yield from Tab //Tbl2b4_3 DATA//
// --------------------------------------------------
function CalcYield_2b4_3(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 85        // Table starts at this row
   iNumColumns = 4  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("G84").Value)
   iPipeSize( = Worksheets(sTabName).Range("G85").Value)
   sSeamType( = Worksheets(sTabName).Range("G86").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 4, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("G87").Value = lYield)
 }
// --------------------------------------------------
// Comments Get Yield from Tab //Tbl2b4_4 DATA//
// --------------------------------------------------
function CalcYield_2b4_4(){
   var lYield  , iPipeSize  , sSeamType  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   var sOutput  
   iStartingRow = 101        // Table starts at this row
   iNumColumns = 10  // Number of columns in table
   sTabName = "Logic" //Name of sheet here
   sOutput( = "Pipe Data")
   vPurchaseDate( = Worksheets(sTabName).Range("M102").Value)
   iPipeSize( = Worksheets(sTabName).Range("M103").Value)
   sSeamType( = Worksheets(sTabName).Range("M104").Value)
   lYield( = fnxDetermineYield(sTabName, iStartingRow, 10, iPipeSize, sSeamType, vPurchaseDate, 1))
   Worksheets((sTabName).Range("M105").Value = lYield)
 }
// --------------------------------------------------
// Comments Get Thickness from Tab //Tbl3 PipeWTCurrentLogic//
// --------------------------------------------------
function CalcThickness_3_2(){
   var vThickness  Variant, iDiameter  , sLogic  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   iStartingRow = 145       // Table starts at this row
   iNumColumns = 8         // Number of columns in table
   sTabName( = "Logic")
   vPurchaseDate( = Worksheets(sTabName).Range("K145").Value)
   iDiameter( = Worksheets(sTabName).Range("K146").Value)
   sLogic( = Worksheets(sTabName).Range("K147").Value)
   // Only 8 columns
   vThickness( = fnxDetermineThickness(sTabName, iStartingRow, 8, iDiameter, sLogic, vPurchaseDate, 1))
   Worksheets((sTabName).Range("K148").Value = vThickness)
 }
// --------------------------------------------------
// Comments Get Thickness from Tab //Tbl3 PipeWTCurrentLogic//
// --------------------------------------------------
function CalcThickness_5B_SEAM(){
   var vThickness  Variant, iDiameter  , sLogic  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer
   iStartingRow = 187        // Table starts at this row
   iNumColumns = 5         // Number of columns in table
   sTabName( = "Logic")
   vPurchaseDate( = Worksheets(sTabName).Range("H187").Value)
   iDiameter( = Worksheets(sTabName).Range("H188").Value)
   sLogic( = Worksheets(sTabName).Range("H190").Value)
   // Only 8 columns
   vThickness( = fnxDetermineThickness(sTabName, iStartingRow, 5, iDiameter, sLogic, vPurchaseDate, 1))
   Worksheets((sTabName).Range("H189").Value = vThickness)

 }
// --------------------------------------------------
// Comments Get Thickness from Tab //Tbl3 PipeWTCurrentLogic//
// --------------------------------------------------
// START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
function CalcThickness_UnkDate(){
   var vThickness  Variant, iDiameter  , sLogic  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer, sStartColumn  , sEndColumn  
   iStartingRow = 206        // Table starts at this row
   iNumColumns = 3         // Number of columns in table
   sTabName( = "Logic")
   sStartColumn( = "A")
   sEndColumn( = "B")
   vPurchaseDate( = Worksheets(sTabName).Range("F207").Value)
   iDiameter( = Worksheets(sTabName).Range("F208").Value)
   sLogic( = Worksheets(sTabName).Range("F209").Value)
   // Only 8 columns
   vThickness( = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn))
   Worksheets((sTabName).Range("F210").Value = vThickness)

 }
// --------------------------------------------------
// Comments Get Thickness from Tab //Tbl3 PipeWTCurrentLogic//
// --------------------------------------------------
// START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
function CalcThickness_1st(){
   var vThickness  Variant, iDiameter  , sLogic  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer, sStartColumn  , sEndColumn  
   iStartingRow = 233        // Table starts at this row
   iNumColumns = 3         // Number of columns in table
   sTabName( = "Logic")
   sStartColumn( = "A")
   sEndColumn( = "B")
   vPurchaseDate( = Worksheets(sTabName).Range("F234").Value)
   iDiameter( = Worksheets(sTabName).Range("F235").Value)
   sLogic( = Worksheets(sTabName).Range("F236").Value)
   // Only 8 columns
   vThickness( = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn))
   Worksheets((sTabName).Range("F237").Value = vThickness)

 }
// --------------------------------------------------
// Comments Get Thickness from Tab //Tbl3 PipeWTCurrentLogic//
// --------------------------------------------------
// START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
function CalcThickness_2nd(){
   var vThickness  Variant, iDiameter  , sLogic  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer, sStartColumn  , sEndColumn  
   iStartingRow = 260        // Table starts at this row
   iNumColumns = 3         // Number of columns in table
   sTabName( = "Logic")
   sStartColumn( = "A")
   sEndColumn( = "B")
   vPurchaseDate( = Worksheets(sTabName).Range("F261").Value)
   iDiameter( = Worksheets(sTabName).Range("F262").Value)
   sLogic( = Worksheets(sTabName).Range("F263").Value)
   // Only 8 columns
   vThickness( = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn))
   Worksheets((sTabName).Range("F264").Value = vThickness)

 }
// --------------------------------------------------
// Comments Get Thickness from Tab //Tbl3 PipeWTCurrentLogic//
// --------------------------------------------------
// START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
function CalcThickness_3rd(){
   var vThickness  Variant, iDiameter  , sLogic  , vPurchaseDate  Variant
   var sTabName  , iStartingRow  Integer, iNumColumns  Integer, sStartColumn  , sEndColumn  
   iStartingRow = 287        // Table starts at this row
   iNumColumns = 3         // Number of columns in table
   sTabName( = "Logic")
   sStartColumn( = "A")
   sEndColumn( = "B")
   vPurchaseDate( = Worksheets(sTabName).Range("F288").Value)
   iDiameter( = Worksheets(sTabName).Range("F289").Value)
   sLogic( = Worksheets(sTabName).Range("F290").Value)
   // Only 8 columns
   vThickness( = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn))
   Worksheets((sTabName).Range("F291").Value = vThickness)
 }
// --------------------------------------------------
// Comments fnxDetermineYield -
// Params   p_sWorksheet name of worksheet (i.e. Pipe Data);
//
// Returns  long (-1 implies error, 0 implies cannot be determined)
// Created  10/25/2012 example
// --------------------------------------------------
 function fnxDetermineYield(p_sWorksheet  , p_iStartingRow  Integer, p_iNumColumns  Integer, _{
        p_iPipeSize  , p_sSeamType  , p_vPurchasDate  Variant, i_UseDateShift  Integer)  
    On( Error GoTo PROC_ERR)
    var PossibleYields(100)  , iYieldIndex  Integer, lBestYield  
    var iCurrentRow  Integer, iRowFound  Integer, sPipeCellValue  , iCurColumn  Integer
    var dMinDate  Date, dMaxDate  Date, vDateVal  Variant, dParamDate  Date, idx  Integer
    // default answer for Yield is 0 .. which means cannot be determined
    lBestYield( = 0)
    iRowFound( = 0)
    // The first row of data starts 2 rows after the header
    iCurrentRow( = p_iStartingRow + 2)
    // First find mathcing Pipe Size / Seam Type Row
    Do while Len(Trim(Worksheets(p_sWorksheet).Range("A" + Trim(Str(iCurrentRow))).Value)) > 0) {
        sPipeCellValue = Worksheets(p_sWorksheet).Range("A" + Trim(Str(iCurrentRow))).Value
        // Check to ensure we//re looking at a numeric value
        if( IsNumeric(sPipeCellValue) ) {
            // See if paramater matches spreadsheet
            if( Val(sPipeCellValue) == p_iPipeSize ) {
                // Now see if Seam Type matches
                if( p_sSeamType == Trim(Worksheets(p_sWorksheet).Range("B" + Trim(Str(iCurrentRow))).Value) ) {
                    iRowFound( = iCurrentRow)
                     Do
                 }
             }
         }
        iCurrentRow( = iCurrentRow + 1)
    Loop()
    // Found matching Pipe Size and Seam Type
    if( iRowFound > 0 ) {
        // Now that we found the right row .. time to get the yield
        // .. get all possible yields and store in array PossibleYields
        // .. then get minimum value of all Yields
        iYieldIndex( = 0)
        // Unknown date is translated into 1/1/1900
        if( IsDate(p_vPurchasDate) ) {
            dParamDate( = p_vPurchasDate)
        }else{
            dParamDate( = #1/1/1900#)
         }
        // See if Purchase Date is within Range of Install Dates
        for( iCurColumn = 3 To p_iNumColumns) {
            // Chr(65) = "A"
            // Get Min Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(p_iStartingRow))).Value
            if( IsDate(vDateVal) ) {
                dMinDate( = DateValue(vDateVal))
            }else{
                dMinDate( = #1/1/1900#)
             }
            // Get Max Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(p_iStartingRow + 1))).Value
            if( IsDate(vDateVal) ) {
                // Add 10 years to compensate for Install Datee
                // dMaxDate = DateAdd("YYYY", 10, DateValue(vDateVal))
                dMaxDate( = DateValue(vDateVal))
            }else{
                dMaxDate( = #1/1/2100#)
             }
            // Failing here
            if( (dParamDate >== dMinDate And dParamDate <== dMaxDate) || _
                DateAdd("d", -1, DateAdd("YYYY", 10, DateValue(dParamDate))) >= dMinDate And dParamDate <= dMaxDate ) {
                //DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate ) {
                // This cell is in range
                iYieldIndex( = iYieldIndex + 1)
                PossibleYields(iYieldIndex) = Val(Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(iRowFound))).Value)
             }
        } iCurColumn
        // Find lowest value in Array .. Use 1st selection as starting point
        lBestYield( = IIf(PossibleYields(1) = 0, ERROR_YIELD, PossibleYields(1)))
        // Loop through all possible Yiels values .. select lowest value
        if( i_UseDateShift == 1 ) {
            for( idx = 1 To iYieldIndex) {
                if( IIf(PossibleYields(idx) == 0, ERROR_YIELD, PossibleYields(idx)) < lBestYield ) {
                    lBestYield( = PossibleYields(idx))
                 }
            } idx
            fnxDetermineYield( = IIf(lBestYield = ERROR_YIELD, 0, lBestYield))
        }else{
            fnxDetermineYield( = IIf(lBestYield = ERROR_YIELD, 0, lBestYield))
         }
    }else{
        fnxDetermineYield( = 0)
     }
PROC_EXIT()
     function{
PROC_ERR()
    fnxDetermineYield( = -1)
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in fnxDetermineYield"
    Resume( PROC_EXIT)
 }
 function xTestYield()  {
  xTestYield( = fnxDetermineYield("Tbl2 PipeYSCurrentLogic_Mas", 37, 10, 24, "DSAW", #1/1/1980#, 1))
 }
// --------------------------------------------------
// Comments fnxDetermineThickness -
// Params   p_sWorksheet name of worksheet (i.e. Tbl3 PipeWTCurrentLogic);
//
// Returns  long (-1 implies error, 0 implies cannot be determined)
// Created  10/25/2012 example
// --------------------------------------------------
 function fnxDetermineThickness(p_sWorksheet  , p_iStartingRow  Integer, p_iNumColumns  Integer, _{
        p_iDiameter  , p_sLogic  Variant, p_vPurchasDate  Variant, i_UseDateShift  Integer)  Variant
    On( Error GoTo PROC_ERR)
    var PossibleThickness(100)  Variant, iDiamIndex  Integer, vBestDiameter  Variant
    var iCurrentRow  Integer, iRowFound  Integer, sDiamCellValue  , iCurColumn  Integer
    var dMinDate  Date, dMaxDate  Date, vDateVal  Variant, dParamDate  Date, idx  Integer
    var dThickness  
    var Test  
    // default answer for Yield is 0 .. which means cannot be determined
    vBestDiameter( = 0)
    iRowFound( = 0)
    // The first row of data starts 2 rows after the header
    iCurrentRow( = p_iStartingRow + 2)
    // First find mathcing Diameter / Seam Type Row
    Do while Len(Trim(Worksheets(p_sWorksheet).Range("A" + Trim(Str(iCurrentRow))).Value)) > 0) {
        sDiamCellValue = Worksheets(p_sWorksheet).Range("A" + Trim(Str(iCurrentRow))).Value
        // Check to ensure we//re looking at a numeric value
        if( IsNumeric(sDiamCellValue) ) {
            // See if paramater matches spreadsheet
            if( Val(sDiamCellValue) == p_iDiameter ) {
                // Now see if Seam Type matches (p_sLogic is the Seam)
                if( p_sLogic == Trim(Worksheets(p_sWorksheet).Range("B" + Trim(Str(iCurrentRow))).Value) || _
                        Len(Trim(Worksheets(p_sWorksheet).Range("B" + Trim(Str(iCurrentRow))).Value)) = 0 ) {
                    iRowFound( = iCurrentRow)
                     Do
                 }
             }
         }
        iCurrentRow( = iCurrentRow + 1)
    Loop()
    // Found matchin Pipe Size and Seam Type
    if( iRowFound > 0 ) {
        // Now that we found the right row .. time to get the yield
        // .. get all possible yields and store in array PossibleThickness
        // .. then get minimum value of all Yields
        iDiamIndex( = 0)
        // Unknown date is translated into 1/1/1900
        if( IsDate(p_vPurchasDate) ) {
            dParamDate( = p_vPurchasDate)
        }else{
            dParamDate( = #1/1/1800#)
         }
        
        // See if Purchase Date is within Range of Install Dates
        for( iCurColumn = 3 To p_iNumColumns) {
            //////////
            // Get Min/Max values in Columns
            //////////
            // Chr(65) = "A"
            // Get Min Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(p_iStartingRow))).Value
            if( IsDate(vDateVal) ) {
                dMinDate( = DateValue(vDateVal))
            }else{
                dMinDate( = #1/1/1800#)
             }
            // Get Max Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(p_iStartingRow + 1))).Value
            if( IsDate(vDateVal) ) {
                // Add 10 years to compensate for Install Datee
                dMaxDate( = DateValue(vDateVal))
            }else{
                dMaxDate( = #1/1/2100#)
             }
            //////////////
            // Determine if Purchase date is within bounds of min/max dates
            //////////////
            if( (dParamDate >== dMinDate And dParamDate <== dMaxDate) || _
                DateAdd("d", -1, DateAdd("YYYY", 10, DateValue(dParamDate))) >= dMinDate And dParamDate <= dMaxDate ) {
                //DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate ) {
                //DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate ) {
                // This cell is in range
                iDiamIndex( = iDiamIndex + 1)
                PossibleThickness(iDiamIndex) = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(iRowFound))).Value
             }
        } iCurColumn
        // Find lowest value in Array .. Use 1st selection as starting point
        vBestDiameter( = PossibleThickness(1))
        // Look for miniumum .. if alpha selected .. then always take alpha .. no minimum needed
        //if( i_UseDateShift == 1 then choose lowest value within range, otherwise just choose single value
        if( i_UseDateShift == 1 ) {
            if( IsNumeric(vBestDiameter) ) {
                // Loop through all possible Thickness values
                // .. field investig will always win
                for( idx = 1 To iDiamIndex) {
                    // Always choose non numeric thichkness
                    if( IsNumeric(PossibleThickness(idx)) ) {
                        if( CDec(PossibleThickness(idx)) < vBestDiameter ) {
                            vBestDiameter( = PossibleThickness(idx))
                         }
                    }else{
                        // field investig always selected
                        // vBestDiameter = PossibleThickness(idx)
                        //
                        //  break
                     }
                } idx
             }
         }
        fnxDetermineThickness( = vBestDiameter)
    }else{
        fnxDetermineThickness( = 0)
     }
PROC_EXIT()
     function{
PROC_ERR()
    fnxDetermineThickness( = -1)
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in fnxDetermineThickness"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------
// Comments fnxDetermineThickness -
// Params   p_sWorksheet name of worksheet (i.e. Tbl3 PipeWTCurrentLogic);
//
// Returns  long (-1 implies error, 0 implies cannot be determined)
// Created  08/15/2011 example
// --------------------------------------------------
 function fnxDetermineThickness1(p_sWorksheet  , p_iStartingRow  Integer, p_iNumColumns  Integer, _{
        p_iDiameter  , p_sLogic  Variant, p_vPurchasDate  Variant, sStartColumn  , sEndColumn  )  Variant
    On( Error GoTo PROC_ERR)
    var PossibleThickness(100)  Variant, iDiamIndex  Integer, vBestDiameter  Variant
    var iCurrentRow  Integer, iRowFound  Integer, sDiamCellValue  , iCurColumn  Integer
    var dMinDate  Date, dMaxDate  Date, vDateVal  Variant, dParamDate  Date, idx  Integer
    var dThickness  
    var Test  
    // default answer for Yield is 0 .. which means cannot be determined
    vBestDiameter( = 0)
    iRowFound( = 0)
    // The first row of data starts 2 rows after the header
    iCurrentRow( = p_iStartingRow + 2)
    // First find mathcing Diameter / Seam Type Row
    Do while Len(Trim(Worksheets(p_sWorksheet).Range("A" + Trim(Str(iCurrentRow))).Value)) > 0) {
        sDiamCellValue = Worksheets(p_sWorksheet).Range("A" + Trim(Str(iCurrentRow))).Value
        // Check to ensure we//re looking at a numeric value
        if( IsNumeric(sDiamCellValue) ) {
            // See if paramater matches spreadsheet
            if( Val(sDiamCellValue) == p_iDiameter ) {
                // Now see if Seam Type matches (p_sLogic is the Seam)
                if( p_sLogic == Trim(Worksheets(p_sWorksheet).Range("B" + Trim(Str(iCurrentRow))).Value) || _
                        Len(Trim(Worksheets(p_sWorksheet).Range("B" + Trim(Str(iCurrentRow))).Value)) = 0 ) {
                    iRowFound( = iCurrentRow)
                     Do
                 }
             }
         }
        iCurrentRow( = iCurrentRow + 1)
    Loop()
    // Found matchin Pipe Size and Seam Type
    if( iRowFound > 0 ) {
        // Now that we found the right row .. time to get the yield
        // .. get all possible yields and store in array PossibleThickness
        // .. then get minimum value of all Yields
        iDiamIndex( = 0)
        // Unknown date is translated into 1/1/1900
        if( IsDate(p_vPurchasDate) ) {
            dParamDate( = p_vPurchasDate)
        }else{
            dParamDate( = #1/1/1800#)
         }
        // See if Purchase Date is within Range of Install Dates
        for( iCurColumn = 3 To p_iNumColumns) {
            //////////
            // Get Min/Max values in Columns
            //////////
            // Chr(65) = "A"
            // Get Min Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(p_iStartingRow))).Value
            if( IsDate(vDateVal) ) {
                dMinDate( = DateValue(vDateVal))
            }else{
                dMinDate( = #1/1/1800#)
             }
            // Get Max Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(p_iStartingRow + 1))).Value
            if( IsDate(vDateVal) ) {
                // Add 10 years to compensate for Install Datee
                dMaxDate( = DateValue(vDateVal))
            }else{
                dMaxDate( = #1/1/2100#)
             }
            //////////////
            // Determine if Purchase date is within bounds of min/max dates
            //////////////
            if( (dParamDate >== dMinDate And dParamDate <== dMaxDate) || _
                DateAdd("d", -1, DateAdd("YYYY", 10, DateValue(dParamDate))) >= dMinDate And dParamDate <= dMaxDate ) {
                //DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate ) {
                //DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate ) {
                // This cell is in range
                iDiamIndex( = iDiamIndex + 1)
                PossibleThickness(iDiamIndex) = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) + Trim(Str(iRowFound))).Value
             }
        } iCurColumn
        // Find lowest value in Array .. Use 1st selection as starting point
        vBestDiameter( = PossibleThickness(1))
        // Look for miniumum .. if alpha selected .. then always take alpha .. no minimum needed
        if( IsNumeric(vBestDiameter) ) {
            // Loop through all possible Thickness values
            // .. field investig will always win
            for( idx = 1 To iDiamIndex) {
                // Always choose non numeric thichkness
                if( IsNumeric(PossibleThickness(idx)) ) {
                    if( CDec(PossibleThickness(idx)) < vBestDiameter ) {
                        vBestDiameter( = PossibleThickness(idx))
                     }
                }else{
                    // field investig always selected
                    // vBestDiameter = PossibleThickness(idx)
                    //
                    //  break
                 }
            } idx
         }
        fnxDetermineThickness1( = vBestDiameter)
    }else{
        fnxDetermineThickness1( = 0)
     }
PROC_EXIT()
     function{
PROC_ERR()
    fnxDetermineThickness1( = -1)
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in fnxDetermineThickness1"
    Resume( PROC_EXIT)
 }

 function xTestThick()  Variant{
//  function fnxDetermineThickness(p_sWorksheet  , p_iStartingRow  Integer, p_iNumColumns  Integer, _{
//        p_iDiameter  Integer, p_sLogic  , p_vPurchasDate  Variant)  Variant
  xTestThick( = fnxDetermineThickness("Tbl3 PipeWTCurrentLogic_Mas", 46, 8, 18, "", #1/1/1950#, 1))
 }

//Option Explicit     // Causes error if variable is not defined with var
////////////////////////////////////////////Basic Functions//////////////////////////////////////////////
// --------------------------------------------------
// Comments fnxIsItIn - Return boolean if 1st paramater is contained in next set of paramaters
// Params   p_lTest is the var to be tested,
//           ParamArray vars is the test set
// Returns  Boolean
// Created  08/15/2011 example
// --------------------------------------------------
 function fnxIsItIn(p_lTest  Variant, ParamArray a_Test())  Boolean{
    On( Error GoTo PROC_ERR)
    var lIncrement  Variant
    fnxIsItIn = false   // Default that p_lTest is not in ParamArray
    // Cycle through ALL lon integers in paramater a_Test
    for ( var lIncrement in a_Test) {
        // if( test va p_lTest is contained in any of the elements of a_Test, return true
        if( p_lTest == lIncrement ) {
            fnxIsItIn = true
             function{
         }
    } lIncrement
    // All variables have been tested
PROC_EXIT()
     function{
PROC_ERR()
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in fnxIsItIn"
    Resume( PROC_EXIT)
 }
 function fnxTestForNull()  Boolean{
    var a
    a( = Worksheets("Pipe Data").Range("O20").Value)
    Debug(.Print a)
    if( Len(Trim(a)) == 0 ) {
        fnxTestForNull = true
    }else{
        fnxTestForNull = false
     }
 }
// --------------------------------------------------
// Comments fnxFindLastRow - Return Last Row, example fnxFindLastRow("Pipe Data", "GG", "OD 1")
// Params   p_sWorksheet name of worksheet (i.e. Pipe Data);
//           p_sColumn is the column to search (i.e. "GG"), p_sColumnHeader is the column header (i.e. "OD 1")
// Returns  long (1st blank row after column header)
// Created  08/15/2011 example
// --------------------------------------------------
 function fnxFindLastRow(p_sWorksheet  , p_sColumn  , p_sColumnHeader  )  {
    On( Error GoTo PROC_ERR)
    var sHeader  
    var lLastRow  
    sHeader = Worksheets(p_sWorksheet).Range(p_sColumn + "3").Value
    if( sHeader != p_sColumnHeader ) {
        fnxFindLastRow( = 0)
        MsgBox "Error (Cannot Find Header Column" + p_sColumnHeader + ")", vbExclamation + vbOKOnly, _
                "Error( in fnxFindLastRow")
    }else{
        lLastRow( = 4)
        Do while Len(Trim(Worksheets(p_sWorksheet).Range(Trim(p_sColumn) + Trim(Str(lLastRow))).Value)) > 0) {
            lLastRow( = lLastRow + 1)
        Loop()
     }
    fnxFindLastRow( = lLastRow - 1)
PROC_EXIT()
     function{
PROC_ERR()
    fnxFindLastRow( = 0)
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in fnxFindLastRow"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------
// Comments subHighlightRow -
// Params  
// Created  08/14/2011 1008 AM
// Modified
// --------------------------------------------------
 function subHighlightRow(p_sWorksheet  , p_lRowNumber  , p_eHighlight  ROW_HIGHLIGHT){
    On( Error GoTo PROC_ERR)
    if( p_lRowNumber > 3 ) {
        Worksheets(p_sWorksheet).Range(Trim(Str(p_lRowNumber)) + "" + Trim(Str(p_lRowNumber))).Select
        Select( Case p_eHighlight)
            Case( ROW_HIGHLIGHT.Highlight_Off)
                Selection(.Font.Color = vbBlack)
                Selection.Font.Bold = false
            Case( ROW_HIGHLIGHT.Highlight_RED)
                Selection(.Font.Color = vbRed)
                Selection.Font.Bold = true
            Case }else{
                Selection(.Font.Color = vbRed)
                Selection.Font.Bold = true
         Select
     }
PROC_EXIT()
     break}
PROC_ERR()
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in subHighlightRow"
    Resume( PROC_EXIT)
 }
 function a()  Boolean{
    subHighlightRow( "Pipe Data", 4, Highlight_Off)
 }

