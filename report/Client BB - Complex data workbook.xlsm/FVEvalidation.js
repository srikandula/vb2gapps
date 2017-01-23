Attribute VB_Name( = "FVEvalidation")
function turnRangeRed(rng  Range){
    rng(.Font.ColorIndex = 3)
 }

function setValidation(rng  Range)  {

var fieldname  
var subrng  Range
var errorMsg  
var fieldArray()  Variant
var rangeName  
var resultRange  Range
var subAddress  
var installDate  Variant
var fittingMAOPrange  Range
var buildArray()  Variant
var MaxWorkingPressure_low  Variant
var MaxWorkingPressure_high  Variant


errorMsg = ""()

fieldArray = Array(("SMYS", "OD 1", "OD 2", "LS Factor", "Seam Type", "Fitting Rating", _)
                "Installed CL", "Installed CL Design Factor", "Today//s CL", "Today//s CL Design Factor", _
                "Fitting( MAOP", "Design Factor", "WT 1", "WT 2", _)
                "Remove( From MAOP Report ""R"" or ""D""", "Component"))
                
MaxWorkingPressure_low = Range(("=MaxWorkingPressure_low").Value)
MaxWorkingPressure_high = Range(("=MaxWorkingPressure_high").Value)

                
for ( var subrng in rng) {

    subAddress = subrng.AddressLocal(rowabsolute=false, columnabsolute=false)
    fieldname( = Sheets("Pipe Data").Cells(2, subrng.Column))
    if( (subrng.row > 2) And ((subrng.Column >== 102) || (subrng.Column == 8)) ) {
        if( inArray(fieldname, fieldArray) And hasValidation(subrng) ) {
            subrng(.Validation.Delete)
         }
        Select( Case fieldname)
            Case( "Component")
                if( featureIsOther(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=ComponentFeatureType_FVE")
                }else{
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=ComponentFeature_FVE")
                 }
            Case( "SMYS")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=SMYS_FVE")
            Case( "OD 1", "OD 2", "Feature")
                ODvalidation( fieldname, subrng)
            Case( "LS Factor")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=LSFactor_FVE")
            Case( "Seam Type")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=SeamType_FVE")
            Case( "Fitting Rating")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=FittingRating_FVE")
            Case( "Installed CL")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=ClassLocation_FVE")
            Case( "Installed CL Design Factor")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=DesignFactor_FVE")
            Case "Today//s CL"
                subrng(.Validation.Add Type=xlValidateList, Formula1="=ClassLocationNoBlank_FVE")
            Case "Today//s CL Design Factor"
                subrng(.Validation.Add Type=xlValidateList, Formula1="=DesignFactor_FVE")
            
            Case( "Fitting MAOP")
            //    subrng.Validation.Add Type=xlValidateList, Formula1="=FittingMAOP_FVE"
                On Error Resume }
                if( isSkidMount(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=modelDynamic")
                else if( isHPR(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=HPRdynamic")
                else if( hasMaxPressure(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateDecimal, Operator=xlBetween, Formula1="=MaxWorkingPressure_low", Formula2="=MaxWorkingPressure_high")
                }else{
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=fittingDynamic")
                 }
                On( Error GoTo 0)
            Case( "Design Factor")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=DesignFactor_FVE")
            Case( "Remove From MAOP Report ""R"" or ""D""")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=RemoveFromMAOPReport_FVE")
            Case( "WT 1")
                subrng(.Validation.Add Type=xlValidateDecimal, Formula1=0.1, Formula2=1.5, Operator=xlBetween)
            Case( "WT 2")
                subrng(.Validation.Add Type=xlValidateDecimal, Formula1=0.1, Formula2=1.5, Operator=xlBetween)
         Select
     }
    
    //set validation for Fitting MAOP if change was made to Feature Type or Figure/Model#
    if( (fieldname == "Type" || fieldname == "Figure - Model #" || fieldname == "Max Working Pressure") And (subrng.row > 2) ) {
         fittingMAOPrange = subrng.Parent.Cells(subrng.row, fittingMAOPColumn)
        if( hasValidation(fittingMAOPrange) ) {
            fittingMAOPrange(.Validation.Delete)
         }
        //On Error Resume }
        if( isSkidMount(subrng.row) ) {
            fittingMAOPrange(.Validation.Add Type=xlValidateList, Formula1="=modelDynamic")
        else if( isHPR(subrng.row) ) {
            fittingMAOPrange(.Validation.Add Type=xlValidateList, Formula1="=HPRdynamic")
        else if( hasMaxPressure(subrng.row) ) {
            fittingMAOPrange(.Validation.Add Type=xlValidateDecimal, Operator=xlBetween, Formula1="=MaxWorkingPressure_low", Formula2="=MaxWorkingPressure_high")
        }else{
            fittingMAOPrange(.Validation.Add Type=xlValidateList, Formula1="=fittingDynamic")
         }
        //On Error GoTo 0
     }
    
    
    if( inArray(fieldname, fieldArray) And (subrng.row > 2) And (subrng.Column >== 102) ) {
        //turnRangeRed subrng
        Select( Case fieldname)
            Case( "WT 1")
                if( IsNumeric(subrng.Value) ) {
                    if( subrng.Value < 0.1 || subrng.Value > 1.5 ) {
                        errorMsg( = addError(errorMsg, fieldname, subrng))
                     }
                else if( subrng.Value != "N/A" ) {
                    errorMsg( = addError(errorMsg, fieldname, subrng))
                 }
            Case( "WT 2")
                if( IsNumeric(subrng.Value) ) {
                    if( subrng.Value < 0.1 || subrng.Value > 1.5 ) {
                        errorMsg( = addError(errorMsg, fieldname, subrng))
                     }
                else if( subrng.Value != "N/A" ) {
                    errorMsg( = addError(errorMsg, fieldname, subrng))
                 }
            Case( "Fitting MAOP")
                if( hasMaxPressure(subrng.row) ) {
                    if( IsNumeric(subrng.Value) ) {
                        if( subrng.Value < MaxWorkingPressure_low || subrng.Value > MaxWorkingPressure_high ) {
                            errorMsg( = addError(errorMsg, fieldname, subrng))
                         }
                    else if( subrng.Value != "N/A" ) {
                        errorMsg( = addError(errorMsg, fieldname, subrng))
                     }
                }else{
                    rangeName( = Replace(subrng.Validation.Formula1, "=", ""))
                     fittingMAOPrange = fittingMAOPcheck(subrng, rangeName)
                    if( fittingMAOPrange Is Nothing ) {
                        if( ! subrng.Value == "" ) {
                            errorMsg( = addError(errorMsg, fieldname, subrng))
                         }
                    }else{
                         resultRange = fittingMAOPrange.Find(subrng.Value, lookat=xlWhole)
                        if( resultRange Is Nothing ) {
                            errorMsg( = addError(errorMsg, fieldname, subrng))
                         }
                     }
                 }
            Case }else{
                rangeName( = Replace(subrng.Validation.Formula1, "=", ""))
                 resultRange = Range(rangeName).Find(subrng.Value, lookat=xlWhole)
                if( resultRange Is Nothing ) {
                    errorMsg( = addError(errorMsg, fieldname, subrng))
                    if( fieldname == "Remove From MAOP Report ""R"" or ""D""" ) {
                        Application.EnableEvents = false
                        subrng(.ClearContents)
                        Application.EnableEvents = true
                        errorMsg = errorMsg + "Value deleted." + vbNewLine
                     }
                 }
         Select
        
        if( errorMsg != "" ) {
            Application.EnableEvents = false
            subrng(.ClearContents)
            Application.EnableEvents = true
         }
     }
}

setValidation = errorMsg()

 }


function addError(errorMsg  , fieldname  , rng  Range){

if( errorMsg == "" ) {
    errorMsg = "Invalid data deleted from the following cells " + vbNewLine
 }

addError = errorMsg + vbNewLine + "Field " + fieldname + vbNewLine
addError = addError + "Value " + rng.Value + vbNewLine
addError = addError + "Address " + rng.Address + vbNewLine


 }

function inArray(searchVal  , searchSet()  Variant)  Boolean{

var i  
inArray = false
for( i = 0 To UBound(searchSet)) {
    if( UCase(searchVal) == UCase(searchSet(i)) ) {
        inArray = true
         function{
     }
}

 }


function hasValidation(cellobj  Range)  Boolean{

    On Error Resume }
        if( cellobj.SpecialCells(xlCellTypeSameValidation).Cells.Count < 1 ) {
            hasValidation = false
        }else{
            hasValidation = true
         }
    On( Error GoTo 0)
 }

function initializeValidation(){

var lastRow  
var sht  Worksheet
var validationRange  Range
var subrng  Range
var fieldname  
var fieldArray()  Variant
var fittingMAOPrange  Range

//sample formula for dynamic validation
//=OFFSET(INDIRECT(ADDRESS(MATCH(H2,OFFSET(A1,0,0,COUNTA(AA),1),0),1)),0,1,COUNTIF(AA,"="&H2),1)
//where H2 is the source cell, A1 and AA are the source range
//can make H2 depend on current cell using =OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),yOffset,xOffset,1,1)
//backref with =fittingRatingRelative
//so=OFFSET(INDIRECT(ADDRESS(MATCH(fittingRatingRelative,//FVE Validation//!$L$3$L$50,0)+2,13,1,true,"FVE Validation")),0,0,COUNTIF(//FVE Validation//!$L$3$L$50,"="&fittingRatingRelative),1)


 sht = Sheets("pipe data")
lastRow = sht.Range(addressForLastRow).(xlUp).row
 validationRange = sht.Range("cx3fc" + lastRow)

fieldArray = Array(("SMYS", "OD 1", "OD 2", "LS Factor", "Seam Type", "Fitting Rating", _)
                "Installed CL", "Installed CL Design Factor", "Today//s CL", "Today//s CL Design Factor", _
                "Fitting( MAOP", "Design Factor", "WT 1", "WT 2", _)
                "Remove( From MAOP Report ""R"" or ""D""", "Component"))

for ( var subrng in validationRange) {
    fieldname( = sht.Cells(2, subrng.Column).Value)
    if( inArray(fieldname, fieldArray) ) {
        subrng(.Validation.Delete)
     }
    if( (! hasValidation(subrng)) And ((subrng.Column >== 102) || (subrng.Column == 8)) ) {
        Select( Case fieldname)
            Case( "Component")
                if( featureIsOther(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=ComponentFeatureType_FVE")
                }else{
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=ComponentFeature_FVE")
                 }
            Case( "SMYS")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=SMYS_FVE")
            Case( "OD 1", "OD 2", "Feature")
                ODvalidation( fieldname, subrng)
            Case( "LS Factor")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=LSFactor_FVE")
            Case( "Seam Type")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=SeamType_FVE")
            Case( "Fitting Rating")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=FittingRating_FVE")
            Case( "Installed CL")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=ClassLocation_FVE")
            Case( "Installed CL Design Factor")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=DesignFactor_FVE")
            Case "Today//s CL"
                subrng(.Validation.Add Type=xlValidateList, Formula1="=ClassLocationNoBlank_FVE")
            Case "Today//s CL Design Factor"
                subrng(.Validation.Add Type=xlValidateList, Formula1="=DesignFactor_FVE")
            Case( "Fitting MAOP")
            //    subrng.Validation.Add Type=xlValidateList, Formula1="=FittingMAOP_FVE"
                On Error Resume }
                if( isSkidMount(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=modelDynamic")
                else if( isHPR(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=HPRdynamic")
                else if( hasMaxPressure(subrng.row) ) {
                    subrng(.Validation.Add Type=xlValidateDecimal, Operator=xlBetween, Formula1="=MaxWorkingPressure_low", Formula2="=MaxWorkingPressure_high")
                }else{
                    subrng(.Validation.Add Type=xlValidateList, Formula1="=fittingDynamic")
                 }
                On( Error GoTo 0)
                
            Case( "Design Factor")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=DesignFactor_FVE")
            Case( "Remove From MAOP Report ""R"" or ""D""")
                subrng(.Validation.Add Type=xlValidateList, Formula1="=RemoveFromMAOPReport_FVE")
            Case( "WT 1")
                subrng(.Validation.Add Type=xlValidateDecimal, Formula1=0.1, Formula2=1.5, Operator=xlBetween)
            Case( "WT 2")
                subrng(.Validation.Add Type=xlValidateDecimal, Formula1=0.1, Formula2=1.5, Operator=xlBetween)
         Select
     }
}


 }


function deleteValidation(){

var lastRow  
var sht  Worksheet
var validationRange  Range
var subrng  Range
var fieldname  
var fieldArray()  Variant



 sht = Sheets("pipe data")
lastRow = sht.Range(addressForLastRow).(xlUp).row
 validationRange = sht.Range("cx3ez" + lastRow)

fieldArray = Array(("SMYS", "OD 1", "OD 2", "LS Factor", "Seam Type", "Fitting Rating", _)
                "Installed CL", "Installed CL Design Factor", "Today//s CL", "Today//s CL Design Factor", _
                "Fitting( MAOP", "Design Factor", "WT 1", "WT 2", _)
                "Remove( From MAOP Report ""R"" or ""D"""))
                
                
for ( var subrng in validationRange) {
    fieldname( = sht.Cells(2, subrng.Column).Value)
    if( inArray(fieldname, fieldArray) ) {
        subrng(.Validation.Delete)
     }
}


 }
function testprop(){
//setValidationStatus (false)
Debug.Print( validationStatus)

 }

function validationStatus()  Boolean{
var sht  Worksheet
var cprop  CustomProperty

 sht = Sheets("fve validation")
for ( var cprop in sht.CustomProperties) {
    if( cprop.Name == "validationStatus" ) {
        validationStatus( = cprop.Value)
         function{
     }
}


 }

function setValidationStatus(Value  Boolean){
var sht  Worksheet
var cprop  CustomProperty

 sht = Sheets("fve validation")
for ( var cprop in sht.CustomProperties) {
    if( cprop.Name == "validationStatus" ) {
        cprop(.Value = Value)
         break}
     }
}

 }

//need to figure out how to trim @crlf from isHPR result
function isHPR(row  )  Boolean{
var sht  Worksheet
var feature  Variant
var featureType  Variant
var HPRval  Variant


 sht = Sheets("Pipe Data")
feature = sht(.Cells(row, buildFeatureColumn))
featureType = sht(.Cells(row, buildFeatureTypeColumn))

if( feature == "FarmTapRegSet" ) {
    HPRval( = featureType)
}else{
    HPRval( = "")
 }
isHPR = Len((HPRval) > 0)

 }

function installDate(row  )  Variant{

var sht  Worksheet

 sht = Sheets("Pipe Data")

installDate = sht(.Cells(row, installDateCol))


 }

function isSkidMount(row  )  Boolean{
var sht  Worksheet
var feature  Variant
var featureType  Variant

 sht = Sheets("Pipe Data")
feature = sht(.Cells(row, buildFeatureColumn))
featureType = sht(.Cells(row, buildFeatureTypeColumn))

isSkidMount = ((feature = "Meter") And ((featureType = "Orifice Skid Mnt - Gas Well") || (featureType = "OrificeSkidMntGasWell")))

 }

function featureIsOther(row  )  Boolean{
var sht  Worksheet
var feature  Variant

 sht = Sheets("Pipe Data")
feature = sht(.Cells(row, buildFeatureColumn))
featureIsOther = (feature( = "Other"))

 }
function hasMaxPressure(row  )  Boolean{
var sht  Worksheet
var maxWorkingPressure  Variant

 sht = Sheets("Pipe Data")
maxWorkingPressure = sht(.Cells(row, maxWorkingPressureColumn))
hasMaxPressure = maxWorkingPressure != ""

 }

function testInstall(){
var datevar  Variant
datevar = installDate((3))
Debug.Print( datevar)
Debug.Print( datevar < #1/2/2013#)
 }

function lastColumn(sht  Worksheet){
    lastColumn = sht.Range(addressForLastColumn).(xlToLeft).Column
 }

function blankRow(row  )  Boolean{
    var sht  Worksheet
    var rng  Range
    var subrng  Range
     sht = Sheets("Pipe Data")
     rng = sht.Range(sht.Cells(row, 1), sht.Cells(row, lastColumn(sht)))
    for ( var subrng in rng) {
        if( VarType(subrng.Value) == vbError ) {
            blankRow = false
             function{
         }
        if( subrng.Value != "" ) {
            blankRow = false
             function{
         }
    }
    blankRow = true
 }

function fittingMAOPcheck(rng  Range, namedRange  )  Range{

var fittingRatingRange  Range
var fittingRatingSubRange  Range
var startRow  
var endRow  
var rowOffset  
var startMAOP  Range
var endMAOP  Range
var valueOffset  
var validationTable  


Select Case( namedRange)
    Case( "modelDynamic")
        valueOffset( = -102)
        validationTable( = "SMMS_FVE")
    Case( "HPRdynamic")
        valueOffset( = -120)
        validationTable( = "HPR_FVE")
    Case( "fittingDynamic")
        valueOffset( = -10)
        validationTable( = "FittingRating_FVE")
 Select


rowOffset = 0()


 fittingRatingRange = Range(validationTable)
 fittingRatingSubRange = fittingRatingRange.Find(rng.Offset(0, valueOffset).Value)
if( ! fittingRatingSubRange Is Nothing ) {
    startRow( = fittingRatingSubRange.row)
    while fittingRatingSubRange.Offset(rowOffset, 0).Value = fittingRatingSubRange.Value) {
        rowOffset( = rowOffset + 1)
    Wend()
    with( ThisWorkbook.Sheets("FVE Validation")) {
         startMAOP = .Cells(fittingRatingSubRange.row, fittingRatingSubRange.Column + 1)
         endMAOP = .Cells(fittingRatingSubRange.row + rowOffset - 1, fittingRatingSubRange.Column + 1)
         fittingMAOPcheck = .Range(startMAOP, endMAOP)
     }
}else{
     fittingMAOPcheck = Nothing
 }

 }

function fittingMAOPresultRange(subrng  Range)  Range{
    var validationRange  Range
     validationRange = fittingMAOPcheck(subrng)
     resultRange = subrng.Find(tempValue, lookat=xlWhole)
 }


function ODvalidation(fieldname  Variant, rng  Range){

Select Case( fieldname)
    Case( "Feature")
        Select( Case rng.Value)
            Case( "Sleeve")
                rng(.Offset(columnoffset=feature_OD1_offset).Validation.Delete)
                rng(.Offset(columnoffset=feature_OD1_offset).Validation.Add Type=xlValidateList, Formula1="=ODlong_FVE")
                rng(.Offset(columnoffset=feature_OD2_offset).Validation.Delete)
                rng(.Offset(columnoffset=feature_OD2_offset).Validation.Add Type=xlValidateList, Formula1="=ODlong_FVE")
            Case }else{
                rng(.Offset(columnoffset=feature_OD1_offset).Validation.Delete)
                rng(.Offset(columnoffset=feature_OD1_offset).Validation.Add Type=xlValidateList, Formula1="=ODshort_FVE")
                rng(.Offset(columnoffset=feature_OD2_offset).Validation.Delete)
                rng(.Offset(columnoffset=feature_OD2_offset).Validation.Add Type=xlValidateList, Formula1="=ODshort_FVE")
         Select
    Case( "OD 1")
        Select( Case rng.Offset(columnoffset=(-1 * feature_OD1_offset)).Value)
            Case( "Sleeve")
                rng(.Validation.Add Type=xlValidateList, Formula1="=ODlong_FVE")
            Case }else{
                rng(.Validation.Add Type=xlValidateList, Formula1="=ODshort_FVE")
         Select
    Case( "OD 2")
        Select( Case rng.Offset(columnoffset=(-1 * feature_OD2_offset)).Value)
            Case( "Sleeve")
                rng(.Validation.Add Type=xlValidateList, Formula1="=ODlong_FVE")
            Case }else{
                rng(.Validation.Add Type=xlValidateList, Formula1="=ODshort_FVE")
         Select

 Select
      
 }
