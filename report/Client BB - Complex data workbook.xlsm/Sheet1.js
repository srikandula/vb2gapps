Attribute VB_Name( = "Sheet1")
Attribute VB_Base( = "0{00020820-0000-0000-C000-000000000046}")
Attribute VB_GlobalNameSpace = false
Attribute VB_Creatable = false
Attribute VB_PredeclaredId = true
Attribute VB_Exposed = true
Attribute VB_TemplateDerived = false
Attribute VB_Customizable = true
//****************************************************************************************************************
//****************************************************************************************************************
//
//    Purpose The purpose of this function is to provide an automated process in suggesting a "Seam", "SMYS"
//             and "Wall Thickness" for pipe, in addition to "SMYS" and "Wall Thickness" for Tee//s, Field Bend//s,
//             Manufactured Bends, and Reducers.
//
//    Version V24 - Phase I
//
//    Modified Date 12/12/2012 920AM
//
//   Log
//       - Updated OD from 0.5 to 0.54
//       - Update 1.32 O.D. to 1.315" O.D. - 12/4/2012
//       - Moved Seam type Exception Logic up - 12/3/2012
//       - Fixed WT2 Blank or 0 suggestion.  Forced WT2 to calculate each time.
//       - Add "Problem" in suggestioned Columns if Pipe is prior purchase or recondition salvaged
//       - Minor tweaks to address macro forcing a Problem suggestion - 11/30/2012
//       - Added conservative ERW Logic with Test date < 7/1/1961 - 11/29/2012
//       - Added suggestions to be forces if auto-populated by macro - 11/29/2012
//       - Added Cap as a new feature for the macro to analyze (Caps are to be treated the same as Mfg Bends - 11/28/2012
//       - Fixed Fitting "MAOP did not pass" comment - 11/28/2012
//       - Updated Seam Type fix for 4" and less to reflect Unknown SMYS from Build in the Exception section - 11/26/2012
//       - Updated new Default values to Unknown per new template - 11/26/2012
//       - Single Row button to return suggestions regardless if a spec value is known - 11/2/2012
//       - Revised Exception Logic - 10/24/2012
//       - Suggestions will not be returned if known specs for Pipes (this logic is already in place for fittings)
//       - Modified Seam Type Dates - 10/18/2012
//       - Added Unk Install Data W.T. 10/18/2012
//       - Fixed autopopulation. - 10/16/2012
//       - Fixed Seam Type function to handle Prior Purchase = "Unkown"
//
//    Required External Modules / Functions
//            - PipeDataModule
//                1.) fnxDetermineYield
//                2.) fnxDetermineThickness
//                3.) fnxDetermineThickness1
//                4.) fnxIsItIn
//                5.) fnxTestForNull
//                6.) fnxFindLastRow
//                7.) subHighlightRow
//
//    Internal Functions
//            - ClearSuggestions
//            - Automation_PFL
//            - SingleSelect_PFL
//            - fnxProcessPipeRow
//                1.) Pipe_Seam()
//                2.) Pipe_SMYS()
//                3.) Pipe_WT()
//                4.) Exceptions()
//                5.) Reconditioned_Salvaged_Pipe()
//                5.) Fittings_Source_WT_1_2 (p_lRowIndex)
//                6.) Fittings_SourceWT()
//                7.) Fittings_TimeFrame1_3()
//                8.) Fittings_TimeFrame4()
//                9.) Fittings_TimeFrame5()
//                10.) Fittings_TimeFrame6()
//                11.) AutoPopulate()
//
//    Output Results
//            - Pipe Data Tab (3 suggestion columns on FVE Side)
//
//    Steps for fnxProcessPipeRow function{
//            1.) Enter in acceptable //types//, column references and table locations
//            2.) Begin stages of analysis to determine Seam, SMYS and W.T. for Pipe
//            3.) Begin stages of analysis to determine SMYS and W.T. for Fittngs
//            4.) Obtain Both Source W.T.s for fittings
//            5.) Choose one Source W.T.
//            6.) Time Frame Logic for Fittings
//            7.) Auto populate function
//
//*****************************************************************************************************************
//*****************************************************************************************************************
Option Explicit()
 btest  Boolean
//Begining and ending row variables
 lLastRowdDiameter  Variant
 iPipeDataLastRow  Integer
 iForceSuggestion  Integer
 iForceSMYS  Integer
 iForceWT  Integer
 iForceSeamType  Integer
//Column Variables
 sFeature  
 sSalvaged  
 iClassLocation  Variant
 sPriorOperator  
 dDiameter  
 dDiameter2  
 dDiameterSmall  
 vPurchaseYr  Variant
 vUserPurchaseYr  Variant
 vSuggestPurchaseYr  Variant
 vInstallYr  Variant
 sSeam  
 sUserSeam  
 sSuggestSeam  
 vSMYS  Variant
 vUserSMYS  Variant
 vWT  Variant
 vWT2  Variant
 vUserWT  Variant
 vUserWT2  Variant
 vLsFactor  Variant
 vSourceWT  Variant
 vSource2WT  Variant
 vSourceWT1  Variant
 vSourceWT2  Variant
 sSourceFeature1  
 sSourceFeature2  
 sSourceFeature3  
 iRowsAway1  Variant
 iRowsAway2  Variant
 iRowsAway3  Variant
//Variables subject to change (Adjustment period between Purchase + Install Date)
 lChangeT  
 lPYearAdjust  
//Variables assigned to Columns
 vBuildWTColumn  Variant
 vBuildWT2Column  Variant
 vBuildSMYSColumn  Variant
 vBuildSeamTypeColumn  Variant
 FeatureColumn  
 iFeatureColumn  Integer
 SalvagedColumn  
 sClassLocationColumn  
 sPriorOperatorColumn  
 sDiamterColumn  
 sDiamterColumn2  
 sUserSeamColumn  
 sInstallYrColumn  
 sTypeColumn  
 vSMYSUserColumn  Variant
 vWTUserColumn  Variant
 vWT2UserColumn  Variant
 vLsFactorColumn  Variant
 vPurchaseDateColumn  Variant
 vFittingMAOPColumn  Variant
 svDesignFactorInstalledClassColumn  
 vSuggestedSMYSColumn  Variant
 vSuggestedWTColumn  Variant
 vSuggestedWT2Column  Variant
 vSuggestedSeamColumn  Variant
 AngleColumn  Variant
//Copy/Paste function{
// iTemplateFound  Integer (! currently in use)
//SME Comments variable
 iSMEComments  Integer
//Pipe/Fitting Logic (starting Row Logic Tables)
 iFVEoption  Integer
 vCurrentLogicStep  Variant
 vSequenceIterator  Variant
 itbl2A2  Integer
 itbl2B3  Integer
 itbl2B4_1  Integer
 itbl2B4_2  Integer
 itbl2B4_Unk1  Integer
 itbl2B4_Unk2  Integer
 itbl2B4_3  Integer
 itbl2B4_4  Integer
 itbl3_2  Integer
 itbl4B  Integer
 itbl5B  Integer
 iUnknownTbl  Integer
 iFirstChoice  Integer
 iSecond_StandardChoice  Integer
 iThirdChoice  Integer
//MAOP Calculation variables
 vMAOP  Variant
 vnewvMAOP  Variant
 vDesignFactor  Variant
 vDesignFactorInstalledClass  Variant
//Invalid entry/Problem variables
 sValidSourceWT  Integer
 sValidSource2WT  Integer
 UnknownInstDate  Integer
 UnknownPurchDate  Integer
 UnknownSeam  Integer
 UnknownSMYS  Integer
 UnknownWT  Integer
 UnknownWT2  Integer
 Problem1_Salvaged  Integer
 Problem2_PriorPurchase  Integer
 Problem3_SMYS  Integer
 Problem4_WT  Integer
 Problem5_Seam  Integer
 Problem6_Exceptions  Integer
 Problem7_SourceWT  Integer
 Problem8_NewSeam  Integer
 DoNotAutoPopSeam  Integer
 DoNotAutoPopSMYS  Integer
 DoNotAutoPopWT  Integer
 DoNotAutoPopWT2  Integer
 SMYS35K  Integer
 Pipe30Inch  Integer
 var SETTING_NOT_APPLICABLE   = 99999
 ManualException  Integer
 iProblemAutoPop  Integer
 sProblemType  
 UseOD2  Integer
// --------------------------------------------------------------------------------------------------------------
// Comments Clears all suggested data in "Suggestion Columns" only
//           This does not clear suggested values that are inserted into other columns (AutoPopulate function){
// Created  11/21/2011 example
// --------------------------------------------------------------------------------------------------------------
function ClearSuggestions(){
var lTotalRows  
var lColumnMax  
var lCurrentRow  
var iColumn  Integer
    //Turn off iForceSuggestion variable
    iForceSuggestion( = 0)
    //Establish Last Row
    lTotalRows = Worksheets("Pipe Data").Cells(Rows.Count, 102).(xlUp).row
    // Current Row of PFL
    lCurrentRow( = 3)
On Error( GoTo Errorcatch)
    //Empty Suggestion Columns
    Worksheets("Pipe Data").Range("DA" + lCurrentRow + "" + "DA" + lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DD" + lCurrentRow + "" + "DD" + lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DF" + lCurrentRow + "" + "DF" + lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DL" + lCurrentRow + "" + "DL" + lTotalRows) = Empty
PROC_EXIT()
     break}
Errorcatch()
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Clearing Suggestions"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments Automation_PFL (AKA All Rows Button) will loop through every row of PFL sheet and execute
//           suggestions on each row
// Returns  Values in suggested fields
// Created  08/21/2011 example
// --------------------------------------------------------------------------------------------------------------
function Automation_PFL(){
On Error( GoTo PROC_ERR)
var lCurrentRow  
var lMod  
var lTotalRows  
    //Row with data starts at Row 3
    lCurrentRow( = 3)
    //Establish Last Row
    lTotalRows = Worksheets("Pipe Data").Cells(Rows.Count, 102).(xlUp).row
    //Empty Suggestion Columns
    Worksheets("Pipe Data").Range("DA" + lCurrentRow + "" + "DA" + lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DD" + lCurrentRow + "" + "DD" + lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DF" + lCurrentRow + "" + "DF" + lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DL" + lCurrentRow + "" + "DL" + lTotalRows) = Empty
    //Update status bar every 10 records
    lMod( = 10)
    //Turn Screen Off
    Application.ScreenUpdating = false
    //Turns off screen updating
    Application.DisplayStatusBar = true
    //Makes sure that the statusbar is visible
    Application(.StatusBar = "Macro Running! (1% Complete)")
    //Loop from Top data row until first blank row (currently using FVE Component Column to measure last row
    Do while Len(Trim(Worksheets("Pipe Data").Range("CX" + Trim(Str(lCurrentRow))).Value)) != Empty) {
        if( ! fnxProcessPipeRow(lCurrentRow) ) {
            //Error if function returns false
             Do
         }
        //Go to } Row
        lCurrentRow( = lCurrentRow + 1)
        if( lCurrentRow Mod lMod == 0 ) {
            Application.StatusBar = "... " + Str(Int(lCurrentRow / lTotalRows * 100)) + "%"
         }
    Loop()
    //Turn Status bar off upon fishing run.
    Application.StatusBar = false
PROC_EXIT()
     break}
PROC_ERR()
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Automation_PFL"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments SingleSelect_PFL (AKA Single Row Button) will execute the macro on the row that is Selected only
// Returns  Values in suggested fields
// Created  08/15/2011 example
// --------------------------------------------------------------------------------------------------------------
function SingleSelect_PFL(){
On Error( GoTo PROC_ERR)
var lCurrentRow  
var lTotalRows  
    //Turn on iForceSuggestion variable
    iForceSuggestion( = 1)
    //Row with data starts at Row
    lCurrentRow( = ActiveCell.row)
    //Establish Last Row
    lTotalRows = Worksheets("Pipe Data").Cells(Rows.Count, 102).(xlUp).row
    //Empty Suggestion Columns
    Worksheets("Pipe Data").Range("DA" + lCurrentRow + "" + "DA" + lCurrentRow) = Empty
    Worksheets("Pipe Data").Range("DD" + lCurrentRow + "" + "DD" + lCurrentRow) = Empty
    Worksheets("Pipe Data").Range("DF" + lCurrentRow + "" + "DF" + lCurrentRow) = Empty
    Worksheets("Pipe Data").Range("DL" + lCurrentRow + "" + "DL" + lCurrentRow) = Empty
    //Added logic to highlig`ht rows in error in red
    if( fnxProcessPipeRow(lCurrentRow) ) {
        //Row Successfully Processed, turn off red/bold font
        //subHighlightRow "Pipe Data", lCurrentRow, ROW_HIGHLIGHT.Highlight_Off
    }else{
        //Error in Row .. highlight as red
        subHighlightRow( "Pipe Data", lCurrentRow, ROW_HIGHLIGHT.Highlight_RED)
     }
    //Reset cursor to the FVE section
    Worksheets(("Pipe Data").Cells(lCurrentRow, 105).Select)
PROC_EXIT()
     break}
PROC_ERR()
    // Error message incorrect, use this as stub
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in SingleSelect_PFL"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments Automaed proccess to return Pipe (Seam, SMYS + Wall Thickness) and Fittings (SMYS + W.T.)
// Returns  Values in suggested fields
// Created  06/21/2012 example
// --------------------------------------------------------------------------------------------------------------
 function fnxProcessPipeRow(ByVal p_lRowIndex  )  Boolean{
On Error( GoTo PROC_ERR)
var iCurrentRow  
var iRowsAway1  Variant
var iRowsAway2  
    //Use Line No. as indicator of LastRow
    lLastRowdDiameter( = p_lRowIndex)
    iProblemAutoPop( = 0)
    ManualException( = 0)
    //*********************// Column Definitions //*********************//
    if( ! ColumnDefinitions() ) {
        fnxProcessPipeRow = false
         function{
     }
    //-----------------------------------------------------------------------------
    //--  1.) Enter in acceptable Feature Types
    //-----------------------------------------------------------------------------
    if( fnxIsItIn(sFeature, "Pipe", "Field Bend", "Mfg Bend", "Reducer", "Tee", "Cap") _
    And Range(vFittingMAOPColumn + CStr(lLastRowdDiameter)).Value = "N/A" _
    And Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value != 0 _
    And Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value != "N/A" ) {
    //----------------------------------------------------------------------------------------
    //--  2.) Begin stages of analysis to determine Seam, SMYS and W.T. for Pipe + Fld Bend --
    //----------------------------------------------------------------------------------------
        if( fnxIsItIn(sFeature, "Pipe", "Field Bend") ) {
            //*********************// SEAM Analysis //*********************//
            if( ! Pipe_Seam() ) {
                fnxProcessPipeRow = false
                 function{
             }
            //*********************// SMYS Analysis //*********************//
            if( ! Pipe_SMYS() ) {
                fnxProcessPipeRow = false
                 function{
             }
            //*********************// W.T. Analysis //*********************//
            if( ! Pipe_WT() ) {
                fnxProcessPipeRow = false
                 function{
             }
             //***************// Exceptions Analysis //***************//
            if( ! Exceptions() ) {
                fnxProcessPipeRow = false
                 function{
             }
//-----------------------------------------------------------------------------
//--  3.) Begin stages of analysis to determine SMYS and W.T. for Fittngs    --
//-----------------------------------------------------------------------------
        }else{
        //CHOOSE LARGER OD FOR SMYS
            //Currently handles Tee, Field Bends, Mfg Bend and Reducer (W)
            if( fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") ) {
                // Adjustment Dates back to 0 Becuase they should only be delt with pipes
                // time change to 0 Years for Fittings
                lChangeT( = 0)
                // adjustment time for Purchase year to 0 Years
                // adjustment time for Purchase year to 0 Years
                if( vInstallYr >== DateValue("3/1/1940") ) {
                    lPYearAdjust( = 0)
                 }
                // Design Factor
                if( (Worksheets("Pipe Data").Range("DR" + Trim(Str(p_lRowIndex))).Value == "Error" _
                || Worksheets("Pipe Data").Range("DR" + Trim(Str(p_lRowIndex))).Value = "N/A" _
                || Worksheets("Pipe Data").Range("DR" + Trim(Str(p_lRowIndex))).Value = "0") _
                And (Worksheets("Pipe Data").Range("DT" + Trim(Str(p_lRowIndex))).Value = "N/A" _
                || Worksheets("Pipe Data").Range("DT" + Trim(Str(p_lRowIndex))).Value = "Error" _
                || Worksheets("Pipe Data").Range("DT" + Trim(Str(p_lRowIndex))).Value = "0") ) {
                    vDesignFactor( = "Error")
                }else{
                     //Process 15.)
                    if( Worksheets("Pipe Data").Range(svDesignFactorInstalledClassColumn + Trim(Str(p_lRowIndex))).Value != "" _
                    And Worksheets("Pipe Data").Range(svDesignFactorInstalledClassColumn + Trim(Str(p_lRowIndex))).Value != "N/A" _
                    And Worksheets("Pipe Data").Range(svDesignFactorInstalledClassColumn + Trim(Str(p_lRowIndex))).Value != "ERROR" ) {
                        vDesignFactor = Worksheets("Pipe Data").Range("DR" + Trim(Str(p_lRowIndex))).Value
                    }else{
                        vDesignFactor = Worksheets("Pipe Data").Range("DT" + Trim(Str(p_lRowIndex))).Value
                     }
                 }
                //if( design factor not valid, insert comments into FVE comments
                if( vDesignFactor == "Error" ) {
                    Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value = "D.F. N/A"
                    Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value = "D.F. N/A"
                    if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Design Factor Problem.*") ) {
                        //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "Design Factor Problem. ")
                        Range("EV" + CStr(lLastRowdDiameter)).Select
                        with( ActiveCell) {
                            .Characters((Len(.Value) + 1).Insert UCase("Design Factor Problem. "))
                         }
                        btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "Design Factor Problem. ")
                     }
                }else{
//-----------------------------------------------------------------------------
//--  4.) Obtain Both Source W.T.s for fittings                              --
//-----------------------------------------------------------------------------
                //***************// Source W.T. (1&2) Analysis //***************//
                if( ! Fittings_Source_WT_1_2(p_lRowIndex) ) {
                    fnxProcessPipeRow = false
                     function{
                 }
                //Error Handling (Instances where source W.T. is Blank or N/A) (G)
                if( sValidSourceWT == 1 ) {
//-----------------------------------------------------------------------------
//--  5.) Choose one Source W.T.                                             --
//-----------------------------------------------------------------------------
                    //***************// Primary Source W.T. Analysis //***************//
                    if( ! Fittings_SourceWT() ) {
                        fnxProcessPipeRow = false
                         function{
                     }
                    //Determine Class Location and choose nesign Factor if Class Location = 4
                    if( iClassLocation == "4" And (DateValue(vInstallYr) < DateValue("7/1/1961")) ) {
                        vDesignFactor( = 0.5)
                        if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61*") ) {
                            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61 ")
                            Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61 "))
                             }
                            btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61 ")
                         }
                     }
//-----------------------------------------------------------------------------
//--  6.) Time Frame Logic for Fittings                                      --
//-----------------------------------------------------------------------------
                    //*************************// Time Frame 1-3 //*************************//
                    if( ! Fittings_TimeFrame1_3() ) {
                        fnxProcessPipeRow = false
                         function{
                     }
                    //Establish Variables and formulas for Time Frames (4-6)
                      // MAOP From MAOP of R Field (MAOP of R will default to 0 if no values are entered in by user)
                    vMAOP = Worksheets("Pipe Data").Range("EK" + Trim(Str(p_lRowIndex))).Value
                    //Default Current Logic
                    vCurrentLogicStep( = SETTING_NOT_APPLICABLE)
                    //Default Iterator to stage 1 very first instance
                    vSequenceIterator( = "Start")
                    //Error Handling for SMYS
                    if( vSMYS Like "Problem" ) {
                        vSMYS( = 0)
                     }
                    //**************************// Time Frame 4 //**************************//
                    //Time Frame 4 3/1/1940 - 4/11/1963
                    if( ! Fittings_TimeFrame4() ) {
                        fnxProcessPipeRow = false
                         function{
                     }
                    //**************************// Time Frame 5 //**************************//
                    //Time Frame 5 4/12/1963 - 10/31/1968
                    if( ! Fittings_TimeFrame5() ) {
                        fnxProcessPipeRow = false
                         function{
                     }
                    //**************************// Time Frame 6 //**************************//
                    //Time Frame 6 >= 11/1/1968
                    //if( ! Fittings_TimeFrame6() ) {
                    //    fnxProcessPipeRow = false
                    //     function{
                    // }
                }else{
                    //**************************// Source W.T. ! Available //**************************//
                    if( ! SourceWT_NA() ) {
                        fnxProcessPipeRow = false
                         function{
                     }
                 }
                //See Logic Tables (PRUF should output NA for 0.54 Diameter
                if( vWT == 0 || dDiameter == 0.54 ) {
                    vWT( = "N/A")
                 }
             }
            // (W) if( Worksheets("Pipe Data").Range("H" + Trim(Str(p_lRowIndex))).Value == "Tee"
             } //Ends handling of fittings (W)
         } //Ends Pipe/Other Handler (A)
        //***************// Salvaged Analysis //***************//
        if( ! Salvaged() ) {
            fnxProcessPipeRow = false
             function{
         }
        //***************// Prior Purchase Analysis //***************//
        if( ! PriorPurchase() ) {
            fnxProcessPipeRow = false
             function{
         }
        //***************// Problems Analysis //***************//
        if( ! Problems() ) {
            fnxProcessPipeRow = false
             function{
         }
          //Output suggestions Suggestion fields
        if( fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") ) {
            //Populate SMYS
            if( UnknownSMYS == 1 || iForceSuggestion == 1 || iForceSMYS == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                if( (Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 _
                || Problem3_SMYS = 1 || Problem7_SourceWT = 1) ) {
                    //if( Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                        Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopSMYS( = 1)
                    // }
                }else{
                    if( iForceSuggestion == 1 ) {
                        DoNotAutoPopSMYS( = 1)
                     }
                    if( Problem1_Salvaged != 1 And Problem2_PriorPurchase != 1 ) {
                        Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value = vSMYS
                     }
                 }
             }
            //Populate WT
            if( UnknownWT == 1 || iForceSuggestion == 1 || iForceWT == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                if( (Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 || Problem4_WT == 1 || Problem7_SourceWT == 1) ) {
                    //if( Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                        Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopWT( = 1)
                    // }
                }else{
                    if( iForceSuggestion == 1 ) {
                        DoNotAutoPopWT( = 1)
                     }
                    if( Problem1_Salvaged != 1 And Problem2_PriorPurchase != 1 ) {
                        if( UseOD2 == 1 ) {
                            Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value = vWT2
                        }else{
                            Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value = vWT
                         }
                     }
                 }
             }
             //Populate WT2
            if( (UnknownWT2 == 1 || iForceSuggestion == 1 || iForceWT == 1) And Range("DI" + CStr(lLastRowdDiameter)).Value != "N/A" ) {
                if( (Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 || Problem4_WT == 1 || Problem7_SourceWT == 1) ) {
                    //if( Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                        Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopWT2( = 1)
                    // }
                }else{
                    if( iForceSuggestion == 1 ) {
                        DoNotAutoPopWT2( = 1)
                     }
                    if( Problem1_Salvaged != 1 And Problem2_PriorPurchase != 1 ) {
                        if( UseOD2 == 1 ) {
                            Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value = vWT
                        }else{
                            Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value = vWT2
                         }
                     }
                 }
             }
        else if( fnxIsItIn(sFeature, "Pipe", "Field Bend") ) {
            //SMYS
            if( UnknownSMYS == 1 || iForceSuggestion == 1 || iForceSMYS == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                if( (ManualException == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 || Problem3_SMYS == 1 || Problem6_Exceptions == 1) ) {
                    //if( Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                        Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopSMYS( = 1)
                    // }
                }else{
                    if( iForceSuggestion == 1 ) {
                        DoNotAutoPopSMYS( = 1)
                     }
                    if( Problem1_Salvaged != 1 And Problem2_PriorPurchase != 1 ) {
                        Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value = vSMYS
                     }
                 }
             }
            //W.T.
            if( UnknownWT == 1 || iForceSuggestion == 1 || iForceWT == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                if( (ManualException == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 || Problem4_WT == 1 || Problem6_Exceptions == 1) ) {
                    //if( Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                        Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopWT( = 1)
                    // }
                }else{
                    if( iForceSuggestion == 1 ) {
                        DoNotAutoPopWT( = 1)
                     }
                    if( Problem1_Salvaged != 1 And Problem2_PriorPurchase != 1 ) {
                        Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value = vWT
                     }
                 }
             }
            //SEAM TYPE
            if( UnknownSeam == 1 || iForceSuggestion == 1 || iForceSeamType == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                if( (ManualException == 1 || Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 || Problem5_Seam == 1 || Problem6_Exceptions == 1) ) {
                    //if( Problem1_Salvaged == 1 || Problem2_PriorPurchase == 1 ) {
                        Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopSeam( = 1)
                    // }
                }else{
                    if( iForceSuggestion == 1 ) {
                        DoNotAutoPopSeam( = 1)
                     }
                    if( Problem1_Salvaged != 1 And Problem2_PriorPurchase != 1 ) {
                        Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value = sSuggestSeam
                     }
                 }
             }
         }
//-----------------------------------------------------------------------------
//--  7.) Auto populate function                                             --
//-----------------------------------------------------------------------------
        // Handle Auto-populate for instances where Seam, SMYMS or W.T. are blank
        //if( (Problem3_SMYS == 3 || Problem4_WT == 4) ) {
            if( ! AutoPopulate() ) {
                fnxProcessPipeRow = false
                 function{
             }
        // }
     }
//**************************************************************************************************************//
    //Indicate true (code executed without errors)
    fnxProcessPipeRow = true
//Handle All Error Descriptions
PROC_EXIT()
     function{
PROC_ERR()
    fnxProcessPipeRow = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in fnxProcessPipeRow"
    Resume( PROC_EXIT)
 }
//**************************************************************************************************************//
//**************************************************************************************************************//
// -- function Definitions --{
//**************************************************************************************************************//
//**************************************************************************************************************//
// --------------------------------------------------------------------------------------------------------------
// Comments This function defines column positions and variables for input data
// Returns  Column locations
// Created  06/14/2012 example
// --------------------------------------------------------------------------------------------------------------
 function ColumnDefinitions()  Boolean{
    //Last row in Sheet based on Feature Number in FVE
    iPipeDataLastRow = Worksheets("Pipe Data").Cells(Rows.Count, 102).(xlUp).row
// Columns Here
    //Build WT
    vBuildWTColumn( = "P")
    //Build WT2
    vBuildWT2Column( = "R")
    //Build SMYS
    vBuildSMYSColumn( = "U")
    //Build SEAM
    vBuildSeamTypeColumn( = "S")
    //PriorPurchaser
    sPriorOperatorColumn( = "CY")
    //Diameter
    sDiamterColumn( = "DH")
    //Diameter
    sDiamterColumn2( = "DI")
    //Seam User
    sUserSeamColumn( = "DK")
    //SMYS User
    vSMYSUserColumn( = "CZ")
    //WT User
    vWTUserColumn( = "DC")
    //WT User2
    vWT2UserColumn( = "DE")
    //Install Year
    sInstallYrColumn( = "DZ")
    // Seam Factor
    vLsFactorColumn( = "DM")
    //Class Location
    sClassLocationColumn( = "DS")
    //Angle Column for Mfg Bends
    AngleColumn( = "AD")
    //Reconditioned / Salvaged?
    SalvagedColumn( = "AZ")
    //Purchase Date of Feature
    vPurchaseDateColumn( = "AV")
    //Fitting MAOP
    vFittingMAOPColumn( = "DY")
    //Feature
    FeatureColumn( = "CX")
    //Installed CL Design Factor
    svDesignFactorInstalledClassColumn( = "DR")
    //Feature Column Number
    iFeatureColumn( = 102)
    //Suggested SMYS
    vSuggestedSMYSColumn( = "DA")
    //Suggested WT
    vSuggestedWTColumn( = "DD")
    //Suggested WT2
    vSuggestedWT2Column( = "DF")
    //Suggested Seam
    vSuggestedSeamColumn( = "DL")
    // Feature Column
    sFeature = Range(FeatureColumn + CStr(lLastRowdDiameter)).Value
    // Salvaged Column
    sSalvaged = Range(SalvagedColumn + CStr(lLastRowdDiameter)).Value
    // SME comments variable to 0 by default
    iSMEComments( = 0)
    // time change to +10 Years
    lChangeT( = 10)
    // adjustment time for Purchase year to -10 Years
    lPYearAdjust( = -10)
    //Check to ensure users only enter approved values in Col sPriorOperatorColumn
    // PriorPurchase to No if not already specified
    if( Range(sPriorOperatorColumn + CStr(lLastRowdDiameter)).Value != Empty ) {
        sPriorOperator = Range(sPriorOperatorColumn + CStr(lLastRowdDiameter)).Value
    }else{
        Range(sPriorOperatorColumn + CStr(lLastRowdDiameter)).Value = "No"
        sPriorOperator( = "No")
     }
    // dDiameter equal to last row in dDiameter Column
    if( Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value == "PipeOD+(2*SleeveWT)+0.25" || Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value == 0 || Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value == "N/A" || Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" ) {
        dDiameter( = 0)
    }else{
        dDiameter = Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value
     }
    // dDiameter2 equal to last row in dDiameter Column
    if( Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value == "PipeOD+(2*SleeveWT)+0.25" || Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value == 0 || Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value == "N/A" || Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value Like "*unknown*" ) {
        dDiameter2( = 0)
    }else{
        dDiameter2 = Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value
     }
    //Pull in User Specified Seam
    sUserSeam = Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value
    //Pull in User Specified SMYS
    vUserSMYS = Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value
    //Pull in User Specified W.T.
    vUserWT = Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value
    //Pull in User Specified W.T.2
    vUserWT2 = Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value
    // Seam Factor
    vLsFactor = Range(vLsFactorColumn + CStr(lLastRowdDiameter)).Value
    //Class Location
    iClassLocation = Range(sClassLocationColumn + CStr(lLastRowdDiameter)).Value
     //Choose Starting row for the following variables (Look at "Logic" Tab)
       //Note for( all "tbl..." variables - Starting Row will be the first of the two Date rows at the top of the table.) {
       //Note for( the remaining 4 variables, set the starting row to one row above the title (2 rows above where the data starts).) {
    itbl2A2( = 12)
    itbl2B3( = 18)
    itbl2B4_1( = 24)
    itbl2B4_2( = 42)
    itbl2B4_Unk1( = 66)
    itbl2B4_Unk2( = 77)
    itbl2B4_3( = 92)
    itbl2B4_4( = 112)
    itbl3_2( = 156)
    itbl4B( = 196)
    itbl5B( = 201)
    iUnknownTbl( = 220)
    iFirstChoice( = 247)
    iSecond_StandardChoice( = 274)
    iThirdChoice( = 301)
    //Default all Problem Variables to 0
    Problem1_Salvaged( = 0)
    Problem2_PriorPurchase( = 0)
    Problem3_SMYS( = 0)
    Problem4_WT( = 0)
    Problem5_Seam( = 0)
    Problem6_Exceptions( = 0)
    Problem7_SourceWT( = 0)
    Problem8_NewSeam( = 0)
    DoNotAutoPopSeam( = 0)
    DoNotAutoPopSMYS( = 0)
    DoNotAutoPopWT( = 0)
    DoNotAutoPopWT2( = 0)
    SMYS35K( = 0)
    Pipe30Inch( = 0)
    UnknownSeam( = 0)
    UnknownSMYS( = 0)
    UnknownWT( = 0)
    UnknownWT2( = 0)
    UnknownInstDate( = 0)
    UnknownPurchDate( = 0)
    //Default Manual Exception to 0
    ManualException( = 0)
    UseOD2( = 0)
    //Turn off iForceSuggestion variable unless there is a 1 in the Rat//le columns
    iForceSuggestion( = 0)
    if( Worksheets("Pipe Data").Range("DB" + Trim(Str(lLastRowdDiameter))).Value == 1 ) {
        iForceSMYS( = 1)
    }else{
        iForceSMYS( = 0)
     }
    if( Worksheets("Pipe Data").Range("DG" + Trim(Str(lLastRowdDiameter))).Value == 1 ) {
        iForceWT( = 1)
    }else{
        iForceWT( = 0)
     }
    if( Worksheets("Pipe Data").Range("DN" + Trim(Str(lLastRowdDiameter))).Value == 1 ) {
        iForceSeamType( = 1)
    }else{
        iForceSeamType( = 0)
     }
    //vSequenceIterator = 0
    //Check if Seam Type is Unknown
    if( fnxIsItIn(sFeature, "Pipe", "Field Bend") And Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value != "" And Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value != "NA" And Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value != "N/A" ) {
        vInstallYr = DateAdd("d", 1, (DateAdd("yyyy", 10, Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value)))
    }else{
        UnknownPurchDate( = 1)
        if( Range(sInstallYrColumn + CStr(lLastRowdDiameter)).Value == Empty _
        || Range(sInstallYrColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
        || Range(sInstallYrColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" ) {
            UnknownInstDate( = 1)
            vInstallYr( = DateValue("1/1/1930"))
        }else{
            vInstallYr = Range(sInstallYrColumn + CStr(lLastRowdDiameter)).Value
         }
     }
    // }
    // Purchase date to whatever Install date is minus 10 years (-10 yrs + 1 day)
//Process 16.) SMYS and W.T.
    if( Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value != "" And Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value != "NA" ) {
        vPurchaseYr = Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value
    }else{
        UnknownPurchDate( = 1)
        vPurchaseYr( = DateAdd("d", 1, (DateAdd("yyyy", lPYearAdjust, vInstallYr))))
     }
    if( ((Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value Like "*N/A*") _
    And (Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value = 0 || Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value = "" || Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value Like "*N/A*")) ) {
        UnknownSeam( = 1)
    }else{
        UnknownSeam( = 0)
     }
    //Check if SMYS is Unknown for Step 1 of fitting logic
    if( ((Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value == "N/A") _
    And (Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value = 0 || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value = "" || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value = "N/A")) ) {
        UnknownSMYS( = 1)
    }else{
        UnknownSMYS( = 0)
     }
    //Check if W.T. is Unknown for Step 1 of fitting logic
    if( ((Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value == "N/A") _
    And (Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value = "" || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value = Empty) || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value = "N/A" || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*") ) {
        UnknownWT( = 1)
    }else{
        UnknownWT( = 0)
     }
    //Check if W.T.2 is Unknown for Step 1 of fitting logic
    if( ((Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value == "N/A")) _
    And ((Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value = "" || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value = Empty || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value = "N/A" || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*")) _
    And (Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value != "" And Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value != Empty And Range(sDiamterColumn2 + CStr(lLastRowdDiameter)).Value != "N/A" And Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value != "0") ) {
        UnknownWT2( = 1)
    }else{
        UnknownWT2( = 0)
     }
    ColumnDefinitions = true
PROC_EXIT()
     function{
PROC_ERR()
    ColumnDefinitions = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Column Definitions"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function returns a suggested Seam for Pipes
// Returns  Pipe - Seam
// Created  02/27/2012 example
// --------------------------------------------------------------------------------------------------------------
 function Pipe_Seam()  Boolean{
    //Check if there is a Prior Operator (B)
    if( sPriorOperator == "Yes" And fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 3.5, 4.5) ) {
        sSuggestSeam( = "Furnace Butt Weld")
    else if( sPriorOperator = "Yes" And fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, _
      12(.75, 14, 16, 18, 20, 22, _)
      24, 26, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42) ) {
        sSuggestSeam( = "Lap Weld")
    //All cases with no prior operator - //Begin Seam calculation logic for Pipe
    else if( fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375) And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") ) {
        sSuggestSeam( = "Furnace Butt Weld")
    else if( dDiameter = 3.5 And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) <= DateValue("2/12/1983") ) {
        sSuggestSeam( = "Furnace Butt Weld")
    else if( dDiameter = 3.5 And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) >= DateValue("2/13/1983") ) {
        sSuggestSeam( = "Seamless/Electric Resistance Weld")
    else if( dDiameter = 4.5 And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) <= DateValue("10/25/1977") ) {
        sSuggestSeam( = "Furnace Butt Weld")
    else if( dDiameter = 4.5 And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) >= DateValue("10/26/1977") ) {
        sSuggestSeam( = "Seamless/Electric Resistance Weld")
    else if( fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, 12.75) _
      And (sPriorOperator = "No" || sPriorOperator = "N/A") _
      And DateValue(vInstallYr) <= DateValue("12/30/1940") ) {
        sSuggestSeam( = "Lap Weld")
    else if( (fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, 12.75) _
      And (sPriorOperator = "No" || sPriorOperator = "N/A") _
      And DateValue(vInstallYr) >= DateValue("12/31/1940")) ) {
        sSuggestSeam( = "Seamless/Electric Resistance Weld")
    //14" Diameter
    else if( dDiameter = "14" ) {
        sSuggestSeam( = "Seamless/Electric Resistance Weld")
    //16" Diameter
    else if( dDiameter = "16" _
      And (sPriorOperator = "No" || sPriorOperator = "N/A") _
      And DateValue(vInstallYr) <= DateValue("12/30/1958") ) {
        sSuggestSeam( = "AO Smith")
    else if( dDiameter = "16" _
      And (sPriorOperator = "No" || sPriorOperator = "N/A") _
      And (DateValue(vInstallYr) >= DateValue("12/31/1958")) ) {
        sSuggestSeam( = "Seamless/Electric Resistance Weld")
    //18" Diameter
    else if( dDiameter = "18" ) {
        sSuggestSeam( = "Seamless/Electric Resistance Weld/Double Submerged Arc Weld")
    //20-24 + 26 Diameter
    else if( fnxIsItIn(dDiameter, 20, 22, 24, 26) And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) <= DateValue("12/30/1958") ) {
        sSuggestSeam( = "Single Submerged Arc Weld/AO Smith")
    else if( fnxIsItIn(dDiameter, 20, 22, 24) And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") _
      And (DateValue(vInstallYr) >= DateValue("12/31/1958")) ) {
        sSuggestSeam( = "Seamless/Double Submerged Arc Weld")
    //26"
    else if( fnxIsItIn(dDiameter, 26) And (sPriorOperator = "No" || sPriorOperator = "N/A" || sPriorOperator = "Unknown") _
      And (DateValue(vInstallYr) >= DateValue("12/31/1958")) ) {
        sSuggestSeam( = "Double Submerged Arc Weld")
    else if( fnxIsItIn(dDiameter, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42) _
      And (sPriorOperator = "No" || sPriorOperator = "N/A") ) {
        sSuggestSeam( = "Double Submerged Arc Weld")
    }else{
        sSuggestSeam( = "Problem 5")
        Problem5_Seam( = 1)
     } //Ends handling of Seam for Pipe (B)
    //Choose ERW if Seamless/ERW and class location = 1 and install date < 1961
     if( Range("DQ" + CStr(lLastRowdDiameter)).Value Like "*1*" || ((Range("DQ" + CStr(lLastRowdDiameter)).Value == "" || Range("DQ" + CStr(lLastRowdDiameter)).Value == "N/A" || Range("DQ" + CStr(lLastRowdDiameter)).Value == "Unknown") And iClassLocation Like "*1*") ) {
        if( Range("EA" + CStr(lLastRowdDiameter)).Value < DateValue("7/1/1961") And Range("EA" + CStr(lLastRowdDiameter)).Value != "N/A" And sSuggestSeam == "Seamless/Electric Resistance Weld" ) {
            sSuggestSeam( = "Electric Resistance Weld")
            if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Most conservative seam type is ERW due to possibility of low frequency ERW in Class 1 location and a test pressure ratio of < 1.251.*" ) {
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert "Most conservative seam type is ERW due to possibility of low frequency ERW in Class 1 location and a test pressure ratio of < 1.251.  ")
                 }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "Most conservative seam type is ERW due to possibility of low frequency ERW in Class 1 location and a test pressure ratio of < 1.251.  ")
             }
         }
     }
    
        //Exceptions for Seam Types ***********************************************
                if( dDiameter < 5 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 400 _
            And DateValue(vInstallYr) >= DateValue("10/13/1964") ) {
                if( sUserSeam == "Furnace Butt Weld" ) {
                    if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Furnace Butt Weld with MAOP > 400 psig, SME Review Required.  *") ) {
                        Range("EV" + CStr(lLastRowdDiameter)).Select
                                with( ActiveCell) {
                                    .Characters((Len(.Value) + 1).Insert UCase("Furnace Butt Weld with MAOP > 400 psig, SME Review Required.  "))
                                 }
                        btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "Furnace Butt Weld with MAOP > 400 psig, SME Review Required.  ")
                     }
                 }
                //Force Seam Type
                if( dDiameter < 3 ) {
                    sUserSeam( = "Seamless")
                    sSeam( = "Seamless")
                    sSuggestSeam( = "Seamless")
                }else{
                    sUserSeam( = "Seamless/Electric Resistance Weld")
                    sSeam( = "Seamless/Electric Resistance Weld")
                    sSuggestSeam( = "Seamless/Electric Resistance Weld")
                 }
             }
            //Exceptions for Steam Types (SMYS)
            if( vUserSMYS > 30000 And vUserSMYS != "Unknown" And vUserSMYS != "unknown" And ! (vSMYS Like "*Problem*") And dDiameter < 5 ) {
                if( dDiameter < 3 ) {
                    sUserSeam( = "Seamless")
                    sSeam( = "Seamless")
                    sSuggestSeam( = "Seamless")
                }else{
                    sUserSeam( = "Seamless/Electric Resistance Weld")
                    sSeam( = "Seamless/Electric Resistance Weld")
                    sSuggestSeam( = "Seamless/Electric Resistance Weld")
                 }
             }
            //A53
            if( sSeam == "Furnace Butt Weld" And Range("T" + CStr(lLastRowdDiameter)).Value == "ASTM A-53" And dDiameter <== 4 ) {
                vSMYS( = 30000)
             }
            //New Seam Types
            if( fnxIsItIn(sUserSeam, "Spiral Weld post 1966", "Polyethylene Pipe", "Special 0.95", "Special 0.90", "Special 0.85") ) {
                Problem8_NewSeam( = 1)
                vWT( = "Problem 8")
                sSeam( = "Problem 8")
                vSMYS( = "Problem 8")
             }
    
    //Choose Users inputed seam (if there is one) for further SMYS and W.T. calculations
    //if( Range("S" + CStr(lLastRowdDiameter)).Value == Empty || Range("S" + CStr(lLastRowdDiameter)).Value Like "*Unknown*" _

    if( (Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value == 0 || Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || IsEmpty(Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value) == true) ) {
       sSeam( = sSuggestSeam)
    }else{
        // Seam type to User Seam Type
        sSeam( = sUserSeam)
        //Exceptions for new seam types
        if( sUserSeam == "Electric Fusion Weld" ) {
            sSeam( = "Double Submerged Arc Weld")
         }
     }
    Pipe_Seam = true
PROC_EXIT()
     function{
PROC_ERR()
    Pipe_Seam = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Pipe-Seam"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function returns a suggested SMYS for Pipes
// Returns  Pipe - SMYS
// Created  02/27/2012 example
// --------------------------------------------------------------------------------------------------------------
 function Pipe_SMYS()  Boolean{
    //Begin SMYS calculation logic for Pipe (C)
    //Determine if Purchase date is known
    if( dDiameter == 3.5 _
      And (sSeam = "Furnace Butt Weld" || sSeam Like "*Unknown*") ) {
        vSMYS( = fnxDetermineYield("Logic", itbl2A2, 12, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    else if( dDiameter = 4.5 _
      And (sSeam = "Furnace Butt Weld" || sSeam Like "*Unknown*") ) {
        vSMYS( = fnxDetermineYield("Logic", itbl2B3, 12, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    //Unknown Date
    else if( (Range(sInstallYrColumn + CStr(lLastRowdDiameter)).Value = Empty || Range("F" + CStr(lLastRowdDiameter)).Value Like "*Unknown*") _
      And fnxIsItIn(dDiameter, 2.375, 3.5) And (sSeam = sSeam Like "*Unknown*") ) {
        vSMYS( = 25000)
    //NEW TABLE
    else if( fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375) _
      And (sSeam = "Furnace Butt Weld" || sSeam Like "*Unknown*") ) {
        vSMYS( = fnxDetermineYield("Logic", itbl2B4_1, 4, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    else if( fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 3.5, 4.5, 6.625, 8.625) _
      And sSeam != "Furnace Butt Weld" _
      And (sSeam != "Lap Weld") ) { // And dDiameter != 6.625 || dDiameter != 8.625) ) { //Confirm with Jim
        vSMYS( = fnxDetermineYield("Logic", itbl2B4_2, 11, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    else if( fnxIsItIn(dDiameter, 4.5, 6.625, 8.625, 10.75, 12.75) And sSeam = "Lap Weld" ) {
        vSMYS( = fnxDetermineYield("Logic", itbl2B4_Unk1, 5, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    else if( fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26) And sSeam Like "*Unknown*" ) {
        vSMYS( = fnxDetermineYield("Logic", itbl2B4_Unk2, 12, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    else if( fnxIsItIn(dDiameter, 12.75, 16, 20, 22, 24, 26) _
      And (sSeam = "Single Submerged Arc Weld" || sSeam = "AO Smith" || sSeam = "Single Submerged Arc Weld/AO Smith") ) {
        vSMYS( = fnxDetermineYield("Logic", itbl2B4_3, 5, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    else if( dDiameter >= 10.75 And fnxIsItIn(sSeam, "Double Submerged Arc Weld", "Electric Resistance Weld", "Seamless", _
        "Seamless/Double Submerged Arc Weld", "Seamless/Electric Resistance Weld", "Seamless/Electric Resistance Weld/Double Submerged Arc Weld") ) {
        vSMYS( = fnxDetermineYield("Logic", itbl2B4_4, 11, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
    }else{
    //Default to "Problem" if SMYS is not captures in any of the prior logic sets
        vSMYS( = "Problem 3")
        Problem3_SMYS( = 1)
     } //Ends handling of SMYS for Pipe (C)
    //Handle when the Year is not specified. This will overide previous logic regardless.
    if( dDiameter <== 3.5 _
      And sSeam = "Furnace Butt Weld" And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vSMYS( = 25000)
     } //Ends handling of SMYS for Pipe with no Date specified
    if( vSMYS == 0 ) {
        vSMYS( = "Problem 3")
        Problem3_SMYS( = 1)
    }else{
        vSMYS( = vSMYS)
     }
    Pipe_SMYS = true
PROC_EXIT()
     function{
PROC_ERR()
    Pipe_SMYS = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Pipe-SMYS"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function returns a suggested Wall Thickness for Pipes
// Returns  Pipe - W.T.
// Created  08/21/2011 example
// --------------------------------------------------------------------------------------------------------------
 function Pipe_WT()  Boolean{
    //Begin W.T. calculation logic for Pipe
    //Handle Unknown Install Dates
    if( dDiameter == 3.5 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {   //(D)
        vWT( = 0.141)
    else if( dDiameter = 4.5 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        if( sSeam == "Seamless" || sSeam == "Furnace Butt Weld" ) {
            vWT( = 0.148)
        }else{
            vWT( = 0.141)
         }
    else if( dDiameter = 6.625 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.156)
    else if( dDiameter = 8.625 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.172)
    else if( dDiameter = 10.75 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.188)
    else if( dDiameter = 12.75 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.203)
    else if( dDiameter = 16 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.219)
    else if( dDiameter = 18 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.25)
    else if( dDiameter = 20 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.25)
    else if( dDiameter = 22 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.25)
    else if( dDiameter = 24 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.25)
    else if( dDiameter = 26 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.281)
    else if( dDiameter = 30 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.25)
    else if( dDiameter = 32 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.375)
    else if( dDiameter = 34 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.25)
    else if( dDiameter = 36 And (UnknownInstDate != 0 And UnknownPurchDate != 0) ) {
        vWT( = 0.312)
    }else{
        vWT( = fnxDetermineThickness("Logic", itbl3_2, 9, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate))
     } //Ends handling of SMYS for Pipe (D)
    if( vWT == 0 || vWT == "error" ) {
        vWT( = "Problem 4")
        Problem4_WT( = 1)
    }else{
        vWT( = vWT)
     }
    // }
    Pipe_WT = true
PROC_EXIT()
     function{
PROC_ERR()
    Pipe_WT = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Pipe-W.T."
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function handles all of the exceptions for pipe (4", 6", 8")
// Returns  Pipe - Exceptions
// Created  10/24/2012 example
// --------------------------------------------------------------------------------------------------------------
 function Exceptions()  Boolean{
            //Special Case for 4" with 970 MAOP
            if( dDiameter == 4.5 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 970 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("6/14/1962") And DateValue(vPurchaseYr) <= DateValue("10/12/1964") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("6/14/1962") And DateValue(vInstallYr) <= DateValue("10/11/1974")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
            //Special Case for 4" with 678 MAOP
            if( dDiameter == 4.5 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 678 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("1/1/1930") And DateValue(vPurchaseYr) <= DateValue("12/31/1930") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("1/1/1930") And DateValue(vInstallYr) <= DateValue("12/30/1940")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
            //Special Case for 4" with 665 MAOP
            if( dDiameter == 4.5 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 665 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("6/14/1962") And DateValue(vPurchaseYr) <= DateValue("10/12/1964") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("6/14/1962") And DateValue(vInstallYr) <= DateValue("10/11/1974")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
            //Special Case for 6" with 489 MAOP
            if( dDiameter == 6.625 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 489 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) = DateValue("6/17/1948") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("6/17/1948") And DateValue(vInstallYr) <= DateValue("6/16/1958")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
            //Special Case for 8" with 451 MAOP
            if( dDiameter == 8.625 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 451 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) = DateValue("12/21/1945") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("12/21/1945") And DateValue(vInstallYr) <= DateValue("12/20/1955")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
            //Special Case for 8" with 1045 MAOP
            if( dDiameter == 8.625 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 1045 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) <= DateValue("8/22/1932") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) <= DateValue("8/21/1942")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
            //Special Case for 12" with 420 MAOP
            if( dDiameter == 12.75 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 420 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) = DateValue("8/7/1941") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("8/7/1941") And DateValue(vInstallYr) <= DateValue("8/6/1951")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
            //Special Case for 16" with 590 MAOP
            if( dDiameter == 16 _
            And Range("EK" + CStr(lLastRowdDiameter)).Value > 590 _
            And( (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("1/1/1953") And DateValue(vPurchaseYr) <= DateValue("12/31/1954") _)
            || (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("1/1/1953") And DateValue(vInstallYr) <= DateValue("12/30/1964")) _
            || (UnknownInstDate != 0 And UnknownPurchDate != 0)) ) {
                if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") ) {
                    Range("EV" + CStr(lLastRowdDiameter)).Select
                            with( ActiveCell) {
                                .Characters((Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  "))
                             }
                    btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                 }
             }
//??? Jim
            //18" Diameter
            //if( dDiameter == 18 ) {
            //    vSMYS = "Problem"
            //    sSuggestSeam = "Problem"
            //    vWT = "Problem"
            //    if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Exception for Diameter*" ) {
            //        //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|6| Exception for Diameter " + dDiameter + ", See PRUPF Table 3 notes.  ")
            //        Range("EV" + CStr(lLastRowdDiameter)).Select
            //                with( ActiveCell) {
            //                    .Characters(Len(.Value) + 1).Insert "|6| Exception for Diameter " + dDiameter + ", See PRUPF Table 2 notes.  "
            //                 }
            //        btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "|6| Exception for Diameter " + dDiameter + ", See PRUPF Table 2 notes.  ")
            //        //Selection.Font.Color = vbBlue
            //        //Selection.Font.Bold = true
            //     }
            // }
            //Exceptions for Steam Types (MAOP-R)
    Exceptions = true
PROC_EXIT()
     function{
PROC_ERR()
    Exceptions = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Exceptions Function"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function Handles actions for the Reconditioned/Salvaged Column
// Returns  "Problem" in suggestion fields based on given criteria
// Created  07/24/2012 example
// --------------------------------------------------------------------------------------------------------------
 function Salvaged()  Boolean{
//if( Salvaged column is anything but yes (No or Unknown) and Feature is Pipe, Bend (or Tee, Reducer, Mfg Bend + Date < 3/1/1940)
//) { macro should return Problem
    if( (sSalvaged == "Yes" || sSalvaged == "yes") _
    And( ((fnxIsItIn(sFeature, "Pipe", "Field Bend") _)
    || (fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") And (DateValue(vInstallYr) < DateValue("3/1/1940"))))) ) {
        Problem1_Salvaged( = 1)
        DoNotAutoPopSeam( = 1)
        DoNotAutoPopWT( = 1)
        DoNotAutoPopSMYS( = 1)
        DoNotAutoPopWT2( = 1)
    //else if( (sSalvaged = "" || sSalvaged = "unknown" || sSalvaged = "Unknown" || IsEmpty(sSalvaged) = true) _
    //And ((fnxIsItIn(sFeature, "Pipe", "Field Bend"))) ) {
    //    if( ! Range("EV" + CStr(lLastRowdDiameter)).value Like "*Reconditioned/Salvage is marked as Unknown, requires technical peer review. *" ) {
    //        Range("EV" + CStr(lLastRowdDiameter)).value = (Range("EV" + CStr(lLastRowdDiameter)).value + "Reconditioned/Salvage is marked as Unknown, requires technical peer review.  ")
    //        Worksheets("Pipe Data").Range("EV" + CStr(lLastRowdDiameter)).Select
    //        Selection.Font.Color = vbBlue
    //        Selection.Font.Bold = true
    //     }
     }
    Salvaged = true
PROC_EXIT()
     function{
PROC_ERR()
    Salvaged = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Salvaged"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function Handles actions for the Prior Purchased Column
// Returns  "Problem" in suggestion fields based on given criteria
// Created  06/18/2012 example
// --------------------------------------------------------------------------------------------------------------
 function PriorPurchase()  Boolean{
//if( Salvaged column is anything but yes (No or Unknown) and Feature is Pipe, Bend (or Tee, Reducer, Mfg Bend + Date < 3/1/1940)
//) { macro should return Problem
    if( (sPriorOperator == "Yes" || sPriorOperator == "yes") ) {
        Problem2_PriorPurchase( = 1)
        DoNotAutoPopSeam( = 1)
        DoNotAutoPopWT( = 1)
        DoNotAutoPopSMYS( = 1)
        DoNotAutoPopWT2( = 1)
     }
    PriorPurchase = true
PROC_EXIT()
     function{
PROC_ERR()
    PriorPurchase = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in PriorPurchase"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function Handles actions for tall distinct problems a user may run into
// Returns  List of what type of Problem is occuring with description
// Created  06/13/2012 example
// --------------------------------------------------------------------------------------------------------------
 function Problems()  Boolean{
//Default iProblemAutoPop to 0
iProblemAutoPop = 0()
sProblemType = "Problems"()
    if( Problem1_Salvaged == 1 ) {
    Problem1_Salvaged( = 1)
     //Insert FVE Comments Value in Blue
    if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Feature is marked as Reconditioned / Salvaged.*") ) {
        //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|1| Feature is marked as Reconditioned / Salvaged. ")
        Range("EV" + CStr(lLastRowdDiameter)).Select
            with( ActiveCell) {
                .Characters((Len(.Value) + 1).Insert UCase("|1| Feature is marked as Reconditioned / Salvaged. "))
             }
        btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|1| Feature is marked as Reconditioned / Salvaged. ")
        //Selection.Font.Color = vbBlue
        //Selection.Font.Bold = true
     }
    iProblemAutoPop( = 1)
    sProblemType = sProblemType + " |1|"
     }
    if( Problem2_PriorPurchase == 1 ) {
         if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Feature is marked as a Prior Purchase.*") ) {
            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|2| Feature is marked as a Prior Purchase.  ")
            Range("EV" + CStr(lLastRowdDiameter)).Select
            with( ActiveCell) {
                .Characters((Len(.Value) + 1).Insert UCase("|2| Feature is marked as a Prior Purchase.  "))
             }
            btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|2| Feature is marked as a Prior Purchase.  ")
            //Selection.Font.Color = vbBlue
            //Selection.Font.Bold = true
         }
        iProblemAutoPop( = 2)
        sProblemType = sProblemType + " |2|"
     }
    
    if( Problem1_Salvaged != 1 And Problem2_PriorPurchase != 1 ) {
        if( Problem3_SMYS == 1 ) {
             if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*SMYS not possible under given circumstances, refer to PRUPF.*") ) {
                //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|3| SMYS not possible under given circumstances, refer to PRUPF.  ")
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert UCase("|3| SMYS not possible under given circumstances, refer to PRUPF.  "))
                 }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|3| SMYS not possible under given circumstances, refer to PRUPF.  ")
                //Selection.Font.Color = vbBlue
                //Selection.Font.Bold = true
             }
            iProblemAutoPop( = 3)
            sProblemType = sProblemType + " |3|"
         }
        if( Problem4_WT == 1 ) {
             if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*W.T. suggestion not available, or not possible under given circumstances, refer to PRUPF.*") ) {
                //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value) // + "|4| W.T. not possible under given circumstances, refer to PRUPF.  ")
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert UCase("|4| W.T. suggestion not available, or not possible under given circumstances, refer to PRUPF.  "))
                 }
                Worksheets("Pipe Data").Range("EV" + CStr(lLastRowdDiameter)).Select
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "W.T. suggestion not available, or not possible under given circumstances, refer to PRUPF.  ")
                //Selection.Font.Color = vbBlue
                //Selection.Font.Bold = true
             }
            iProblemAutoPop( = 4)
            sProblemType = sProblemType + " |4|"
         }
        if( Problem5_Seam == 1 ) {
             if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Seam not possible under given circumstances, refer to PRUPF.*") ) {
                //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|5| Seam not possible under given circumstances, refer to PRUPF.  ")
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert UCase("|5| Seam not possible under given circumstances, refer to PRUPF.  "))
                 }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|5| Seam not possible under given circumstances, refer to PRUPF.  ")
                //Selection.Font.Color = vbBlue
                //Selection.Font.Bold = true
             }
            iProblemAutoPop( = 5)
            sProblemType = sProblemType + " |5|"
         }
        if( Problem6_Exceptions == 1 ) {
            if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Exception for Diameter*" ) {
                //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|6| Exception for Diameter " + dDiameter + ", See PRUPF Table 2 + 3 notes.  ")
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters(Len(.Value) + 1).Insert "|6| Exception for Diameter " + dDiameter + ", See PRUPF Table 2 + 3 notes.  "
                 }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "|6| Exception for Diameter " + dDiameter + ", See PRUPF Table 2 + 3 notes.  ")
                //Selection.Font.Color = vbBlue
                //Selection.Font.Bold = true
             }
            iProblemAutoPop( = 6)
            sProblemType = sProblemType + " |6|"
         }
        if( Problem7_SourceWT == 1 ) {
             if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Could not find valid Source W.T.*") ) {
                //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|7| Could not find valid Source W.T.  ")
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert UCase("|7| Could not find valid Source W.T.  "))
                 }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|7| Could not find valid Source W.T.  ")
                //Selection.Font.Color = vbBlue
                //Selection.Font.Bold = true
             }
            iProblemAutoPop( = 7)
            sProblemType = sProblemType + " |7|"
         }
        if( Problem8_NewSeam == 1 ) {
             if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Seam type not yet in PRUPF*") ) {
                //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "|8| Seam type not yet in PRUPF.  ")
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert UCase("|8| Seam type not yet in PRUPF.  "))
                 }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "|8| Seam type not yet in PRUPF.  ")
                //Selection.Font.Color = vbBlue
                //Selection.Font.Bold = true
             }
            iProblemAutoPop( = 8)
            sProblemType = sProblemType + " |8|"
         }
        //Check if Mfg Bend has an Angle < 30 degrees (This is information only, no action!)
        if( sFeature == "Mfg Bend" And Worksheets("Pipe Data").Range(AngleColumn + Trim(Str(lLastRowdDiameter))).Value < 30 ) {
            if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Manufactured bend angle is small enough such that feature may be field bend instead - Ref PFL Build Guideline AKM-MAOP-404G.*" ) {
                Range("EV" + CStr(lLastRowdDiameter)).Select
                with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert "Manufactured bend angle is small enough such that feature may be field bend instead - Ref PFL Build Guideline AKM-MAOP-404G.  ")
                 }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "Manufactured bend angle is small enough such that feature may be field bend instead - Ref PFL Build Guideline AKM-MAOP-404G.  ")
             }
         }
     }
    Problems = true
PROC_EXIT()
     function{
PROC_ERR()
    Problems = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Problems"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function calculates Source W.T. 1 and Source W.T. 2 to prep for comparison
// Returns  Fittings - Source W.T. 1 + 2
// Created  02/27/2012 example
// --------------------------------------------------------------------------------------------------------------
 function Fittings_Source_WT_1_2(p_lRowIndex  )  Boolean{
    var iCurrentRow  
    iCurrentRow( = p_lRowIndex)
    iRowsAway1( = 0)
    iRowsAway2( = 0)
    vSourceWT( = 0)
    vSourceWT1( = 0)
    vSourceWT2( = 0)
    //Choose Larger OD for Reducers and Tees by defaulting dDiameter to larger OD
    if( sFeature == "Reducer" || sFeature == "Tee" ) {
        if( dDiameter2 > dDiameter ) {
            dDiameter( = dDiameter2)
            UseOD2( = 1)
         }
     }
    //SourceWT1 with matching Source Diameter1 for Reducers with matching Diameters (Mfg Bend, Tee, Reducer)
    if( sFeature == "Reducer" || sFeature == "Tee" || sFeature == "Mfg Bend" || sFeature == "Cap" ) {
        Do while vSourceWT1 = 0) {
            if( p_lRowIndex == 3 ) { // (E)
                vSourceWT1( = SETTING_NOT_APPLICABLE)
                sSourceFeature1( = "Last Row")
            }else{
                if( (Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow - 1))).Value == "Pipe" _
                || Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow - 1))).Value = "Field Bend") _
                   And Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow - 1))).Value != "N/A" _
                   And Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow - 1))).Value != "" _
                   And dDiameter = Worksheets("Pipe Data").Range(sDiamterColumn + Trim(Str(iCurrentRow - 1))).Value ) {
                    vSourceWT1 = Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow - 1))).Value
                    sSourceFeature1 = Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow - 1))).Value
                    iCurrentRow( = iCurrentRow + 1)
                else if( (iCurrentRow = 2) ) {
                    vSourceWT1( = SETTING_NOT_APPLICABLE)
                }else{
                    iCurrentRow( = iCurrentRow - 1)
                    iRowsAway1( = iRowsAway1 + 1)
                 }
             } //(E)
        Loop()
     }
    //Establish Source2 WT1 for Mfg Bends and Tees
    //Reset Current Row
    iCurrentRow( = p_lRowIndex)
    //Establish 2st Source WT by going down in the spreadhseet (Mfg Bend, Tee)
    if( sFeature == "Mfg Bend" || sFeature == "Tee" || sFeature == "Reducer" || sFeature == "Cap" ) {
        Do while vSourceWT2 = 0) {
            if( p_lRowIndex == iPipeDataLastRow ) { //(F)
                vSourceWT2( = SETTING_NOT_APPLICABLE)
                sSourceFeature3( = "Last Row")
            }else{
                if( (Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow + 1))).Value == "Pipe" _
                || Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow + 1))).Value = "Field Bend") _
                And Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow + 1))).Value != "N/A" _
                And Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow + 1))).Value != "" _
                And dDiameter = Worksheets("Pipe Data").Range(sDiamterColumn + Trim(Str(iCurrentRow + 1))).Value ) {
                    vSourceWT2 = Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow + 1))).Value
                    sSourceFeature2 = Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow + 1))).Value
                    iCurrentRow( = iCurrentRow + 1)
                else if( iCurrentRow = iPipeDataLastRow ) {
                    vSourceWT2( = SETTING_NOT_APPLICABLE)
                }else{
                    iCurrentRow( = iCurrentRow + 1)
                    iRowsAway2( = iRowsAway2 + 1)
                 }
             } //(F)
        Loop()
     }
    //Identify OD2 small OD
    if( UseOD2 == 0 ) {
        dDiameterSmall( = dDiameter2)
    }else{
        dDiameterSmall = Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value //PULL BACK OD1 FROM PFL HERE
     }
    //Source WT for Reducer for smaller OD**************
            //for( Tee OD, Just put nothing and default to Standard Wall as in notes) {
    iRowsAway3( = 0)
    vSource2WT( = 0)
    //Reset Current Row
    iCurrentRow( = p_lRowIndex)
    //Establish OD2 Source WT by going down in the spreadhseet (Reducer)
    if( sFeature == "Reducer" ) {
        Do while vSource2WT = 0) {
            if( p_lRowIndex == iPipeDataLastRow ) { //(F)
                vSource2WT( = SETTING_NOT_APPLICABLE)
                sSourceFeature2( = "Last Row")
            }else{
                if( (Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow + 1))).Value == "Pipe" _
                || Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow + 1))).Value = "Field Bend") _
                And Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow + 1))).Value != "N/A" _
                And Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow + 1))).Value != "" _
                And dDiameterSmall = Worksheets("Pipe Data").Range(sDiamterColumn2 + Trim(Str(iCurrentRow + 1))).Value ) {
                    vSource2WT = Worksheets("Pipe Data").Range(vWTUserColumn + Trim(Str(iCurrentRow + 1))).Value
                    sSourceFeature2 = Worksheets("Pipe Data").Range(FeatureColumn + Trim(Str(iCurrentRow + 1))).Value
                    iCurrentRow( = iCurrentRow + 1)
                else if( iCurrentRow = iPipeDataLastRow ) {
                    vSource2WT( = SETTING_NOT_APPLICABLE)
                }else{
                    iCurrentRow( = iCurrentRow + 1)
                    iRowsAway3( = iRowsAway3 + 1)
                 }
             } //(F)
        Loop()
     }
    //Determind if Source WT is valid or not
    if( ((vSourceWT1 == Empty || vSourceWT1 == "N/A") And (iRowsAway2 == 0 And vSourceWT2 != Empty And vSourceWT2 != "N/A") _
    || (vSourceWT2 = Empty || vSourceWT2 = "N/A") And (iRowsAway1 = 0 And vSourceWT1 != Empty And vSourceWT1 != "N/A")) _
    || ((vSourceWT1 != Empty And vSourceWT1 != "N/A") And (vSourceWT2 != Empty And vSourceWT2 != "N/A")) _
    And (vSourceWT1 != SETTING_NOT_APPLICABLE || vSourceWT2 != SETTING_NOT_APPLICABLE) ) {
        sValidSourceWT( = 1)
    }else{
        sValidSourceWT( = 0)
     }
    //Determind if Source2 WT is valid or not
    if( vSource2WT == Empty || vSource2WT == "N/A" || vSource2WT == 0 || vSource2WT == SETTING_NOT_APPLICABLE ) {
        sValidSource2WT( = 0)
    }else{
        sValidSource2WT( = 1)
     }
    Fittings_Source_WT_1_2 = true
PROC_EXIT()
     function{
PROC_ERR()
    Fittings_Source_WT_1_2 = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in FITTINGS-Source W.T. 1&2"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function determines the source W.T. from comparing Source W.T. 1 and Source W.T. 2
// Returns  Fittings - Source W.T.
// Created  02/27/2012 example
// --------------------------------------------------------------------------------------------------------------
 function Fittings_SourceWT()  Boolean{
//Determine Source WT by comparing and choosing best of two Source WT
        if( vSourceWT1 == SETTING_NOT_APPLICABLE And vSourceWT2 != SETTING_NOT_APPLICABLE ) {  //(H)
            vSourceWT( = vSourceWT2)
        else if( vSourceWT2 = SETTING_NOT_APPLICABLE And vSourceWT1 != SETTING_NOT_APPLICABLE ) {
            vSourceWT( = vSourceWT1)
        else if( vSourceWT2 = SETTING_NOT_APPLICABLE And vSourceWT1 != SETTING_NOT_APPLICABLE ) {
            vSourceWT( = SETTING_NOT_APPLICABLE)
        else if( vSourceWT1 != SETTING_NOT_APPLICABLE And vSourceWT2 != SETTING_NOT_APPLICABLE ) {
            // (I)
            if( (sSourceFeature1 == "Pipe" || sSourceFeature1 == "Field Bend") _
              And (sSourceFeature2 != "Pipe" And sSourceFeature2 != "Field Bend") ) {
                vSourceWT( = vSourceWT1)
            else if( (sSourceFeature2 = "Pipe" || sSourceFeature2 = "Field Bend") _
              And (sSourceFeature1 != "Pipe" And sSourceFeature1 != "Field Bend") ) {
                vSourceWT( = vSourceWT2)
            else if( (sSourceFeature1 = "Pipe" || sSourceFeature1 = "Field Bend") _
              And (sSourceFeature2 = "Pipe" || sSourceFeature2 = "Field Bend") ) {
                //Choose Source WT when both WT are Priorities
                //source-fitting-source = select lowest of the two source WT to serve as the source.
                //source-other-fitting-other-other-other-source = select lowest of the two source WTs.
                //source-fitting-other-source = select the contiguous source.
                if( (iRowsAway1 == iRowsAway2) || (iRowsAway1 > 0 And iRowsAway2 > 0) ) {
                    //Choose to take Maximum or Minimum Source Wall Thickness
                    if( vSourceWT1 < vSourceWT2 ) {
                        vSourceWT( = vSourceWT2)
                    else if( vSourceWT1 > vSourceWT2 ) {
                        vSourceWT( = vSourceWT1)
                    else if( vSourceWT1 = vSourceWT2 ) {
                        vSourceWT( = vSourceWT1)
                     }
                else if( iRowsAway1 = 0 And iRowsAway2 > 0 ) {
                    vSourceWT( = vSourceWT1)
                else if( iRowsAway1 > 0 And iRowsAway2 = 0 ) {
                    vSourceWT( = vSourceWT2)
                 }
             } //(I)
         } //(H)
        Fittings_SourceWT = true
PROC_EXIT()
     function{
PROC_ERR()
    Fittings_SourceWT = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Fittings_SourceWT"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function returns a suggested SMYS + Wall Thickness for Fittings
//           for Time Frames 1-3
// Returns  Fitting - SMYS + W.T. Time Frames 1-3
// Created  08/21/2011 example
// --------------------------------------------------------------------------------------------------------------
 function Fittings_TimeFrame1_3()  Boolean{
//Proccess 22)
    //1.) (Time Frame 1 Unknown Install Date)
    if( (Range("DZ" + CStr(lLastRowdDiameter)).Value == Empty _
    || Range("DZ" + CStr(lLastRowdDiameter)).Value Like "" _
    || Range("DZ" + CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    || Range("DZ" + CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And (Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value = Empty _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "" _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*") ) {
        vWT( = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
        vWT2( = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
        vSMYS( = 24000)
   //  }
    //2.) (Time Frame 2 < 1/1/30)
    else if( ((Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value = Empty _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "" _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And( DateValue(vInstallYr) < DateValue("1/1/1940")) _)
    || DateValue(vPurchaseYr) < DateValue("1/1/1930") ) {
    //(DateValue(vPurchaseYr) < DateValue("1/1/1930")) || (DateValue(vInstallYr) < DateValue("1/1/1940")) ) {
        //Just choose Standard Wall Thickness!
        vWT( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
        vWT2( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
        vSMYS( = 24000)
    // }
    //3.) (Time Frame 3 1/1/30 - 2/28/40)
    else if( ((Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value = Empty _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "" _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    || Range(vPurchaseDateColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And( DateValue(vInstallYr) >= DateValue("1/1/1940") And DateValue(vInstallYr) <= DateValue("2/28/1940")) _)
    || DateValue(vPurchaseYr) >= DateValue("1/1/1930") And DateValue(vPurchaseYr) <= DateValue("2/28/1940") ) {
    //else if( ((DateValue(vPurchaseYr) >= DateValue("1/1/1930") || DateValue(vInstallYr) >= DateValue("1/1/1940")) _
    //And (DateValue(vPurchaseYr) <= DateValue("2/28/1940") || DateValue(vInstallYr) <= DateValue("2/28/1940"))) _
    //And (Range("DZ" + CStr(lLastRowdDiameter)).Value != Empty _
    //And Range("DZ" + CStr(lLastRowdDiameter)).Value != "Unknown" _
    //And Range("DZ" + CStr(lLastRowdDiameter)).Value != "unknown") ) {
        //Just choose Standard Wall Thickness!
        vWT( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
        vWT2( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
        vSMYS( = 30000)
     }
    Fittings_TimeFrame1_3 = true
PROC_EXIT()
     function{
PROC_ERR()
    Fittings_TimeFrame1_3 = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Fittings_TimeFrame1-3"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function returns a suggested SMYS + Wall Thickness for Fittings
//           for Time Frame 4 3/1/40 - 4/11/63
// Returns  Fitting - SMYS + W.T. Time Frame 4
// Created  08/21/2011 example
// --------------------------------------------------------------------------------------------------------------
 function Fittings_TimeFrame4()  Boolean{
    //New vMAOP Calc (0 for first iteration)
    vnewvMAOP( = 0)
    //4.)Time Frame 4 3/1/40 - 4/11/63
    //Starting Step 1
    if( (DateValue(vInstallYr) >== DateValue("3/1/1940")) _
    And (DateValue(vInstallYr) <= DateValue("4/11/1963")) ) {
    //|| (sFeature = "Tee" And DateValue(vInstallYr) >= DateValue("3/1/1940")) ) {
    //OD LARGE
    //Step 1
        if( (vSourceWT <== fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 30000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 2
        if( ((vSourceWT > fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
        And( vSourceWT <= fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _)
        And vSequenceIterator != "Complete")) _
        || vSequenceIterator = "InComplete" ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 30000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
     //Step 3
        if( ((vSourceWT > fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) _
          And (vSequenceIterator != "Complete")) _
          || (vSequenceIterator = "InComplete") ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 30000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 4
        if( vSequenceIterator == "InComplete" ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 35000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 5
        if( vSequenceIterator == "InComplete" ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 35000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 6
        if( vSequenceIterator == "InComplete" ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 35000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
         //Output comment if MAOP is never met
        if( vSequenceIterator == "InComplete" And (UnknownSMYS == 1 || UnknownWT == 1) ) {
            if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" ) {
                 //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                 Range("EV" + CStr(lLastRowdDiameter)).Select
                 with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. ")
                  }
                 btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
              }
         } //
        
        //OD SMALL
        if( dDiameterSmall != 0 And (sFeature == "Reducer" || sFeature == "Tee") ) { //And UnknownWT2 == 1 ) {
             //Choose Smaller Diameter
             //if( UseOD2 == 1 ) {
             //    if( Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value == "PipeOD+(2*SleeveWT)+0.25" ) {
             //        dDiameter = 0
             //    }else{
             //        dDiameter = Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value
             //     }
             //}else{
             //    dDiameter = dDiameterSmall
             // }
            // SourceWT2 if unknown to Unkown Wt from table 7
            //if( sValidSource2WT == 0 ) {
            //    vSource2WT = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            // }
               //Default vSequenceIterator to InComplete
               //vSequenceIterator = "InComplete"
               //Step 1
                   //if( (vSource2WT <== fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
                   //And sFeature != "Tee") ) {
                       //Use Known WT if available
                   vWT2( = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Use Known SMYS from larger OD
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
               // }
           //Step 2
               //if( ((vSource2WT > fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
               //And vSource2WT <= fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _

               if( vSequenceIterator != "Complete" ) {
                   //Use either Standard Choice WT
                       vWT2( = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
                }
            //Step 3
               if( vSequenceIterator != "Complete" ) {
                   //Use either Standard Choice WT
                       vWT2( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
                }
                //Step 3
               if( vSequenceIterator != "Complete" ) {
                   //Use either Standard Choice WT
                    vWT2( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
                }
         //Stop here!
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
               // }
           // }
           //Output comment if MAOP is never met
                if( vSequenceIterator == "InComplete" And (UnknownSMYS == 1 || UnknownWT == 1 || UnknownWT2 == 1) ) {
                    if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" ) {
                        //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                         Range("EV" + CStr(lLastRowdDiameter)).Select
                         with( ActiveCell) {
                            .Characters((Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. ")
                          }
                        btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
                     }
                 } //
         }
     }
    Fittings_TimeFrame4 = true
PROC_EXIT()
     function{
PROC_ERR()
    Fittings_TimeFrame4 = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Fittings_TimeFrame4"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function returns a suggested SMYS + Wall Thickness for Fittings
//           for Time Frame 5 4/12/1963 - 10/31/1968
// Returns  Fitting - SMYS + W.T. Time Frame 5
// Created  08/21/2011 example
// --------------------------------------------------------------------------------------------------------------
 function Fittings_TimeFrame5()  Boolean{
    //Time Frame 5 4/12/1963 - 10/31/1968
    vnewvMAOP( = 0)
    if( (DateValue(vInstallYr) >== DateValue("4/12/1963")) ) {
    //And sFeature != "Tee") ) {    //(J)
            //Step 1
        if( (vSourceWT <== fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 35000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 2
        if( ((vSourceWT > fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
        And vSourceWT <= fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") And vSequenceIterator != "Complete") _
        || vSequenceIterator = "InComplete") ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 35000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
     //Step 3
        if( ((vSourceWT > fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) _
          And (vSequenceIterator != "Complete")) _
          || (vSequenceIterator = "InComplete") ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 35000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 4
        if( vSequenceIterator == "InComplete" ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                //Choose Source WT if larger than Extra Heavy WT
                if( vSourceWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") ) {
                    vWT( = vSourceWT)
                }else{
                    vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
                 }
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 35000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 5
        if( vSequenceIterator == "InComplete" ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                //Choose Source WT if larger than Extra Heavy WT
                if( vSourceWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") ) {
                    vWT( = vSourceWT)
                }else{
                    vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
                 }
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 42000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
    //Step 6
        if( vSequenceIterator == "InComplete" ) {
            //Use Known WT if available
            if( UnknownWT == 0 ) {
                vWT( = vUserWT)
            }else{
                //Choose Source WT if larger than Extra Heavy WT
                if( vSourceWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") ) {
                    vWT( = vSourceWT)
                }else{
                    vWT( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B"))
                 }
             }
            //Use Known SMYS if available
            if( UnknownSMYS == 0 ) {
                vSMYS( = vUserSMYS)
            }else{
                vSMYS( = 52000)
             }
            //Define MAOP
            vnewvMAOP( = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter)))
            //Test MAOP
            if( vnewvMAOP >== vMAOP ) {  //(M)
                vSequenceIterator( = "Complete")
            }else{
                vSequenceIterator( = "InComplete")
             }
         }
        //Output comment if MAOP is never met
        if( vSequenceIterator == "InComplete" And (UnknownSMYS == 1 || UnknownWT == 1) ) {
            if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" ) {
                //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                  Range("EV" + CStr(lLastRowdDiameter)).Select
                  with( ActiveCell) {
                     .Characters((Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. ")
                   }
                  btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
             }
         } //
        //Reset vSequenceIterator
        vSequenceIterator( = "Start")
            //OD SMALL
         if( dDiameter2 != 0 And (sFeature == "Reducer" || sFeature == "Tee") ) { // and UnknownWT2 == 1 ) {
             //Choose Smaller Diameter
             //if( UseOD2 == 1 ) {
             //    if( Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value == "PipeOD+(2*SleeveWT)+0.25" ) {
             //        dDiameter = 0
             //    }else{
             //        dDiameter = Range(sDiamterColumn + CStr(lLastRowdDiameter)).Value
             //     }
             //}else{
             //    dDiameter = dDiameterSmall
             // }
        // SourceWT2 if unknown to Unkown Wt from table 7
        //if( sValidSource2WT == 0 ) {
        //    vSource2WT = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
        // }
           //Default vSequenceIterator to InComplete
           //vSequenceIterator = "InComplete"
           //Step 1
               //if( (vSource2WT <== fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
               //And sFeature != "Tee") ) {
                   //Use Known WT if available
                   vWT2( = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Use Known SMYS from larger OD
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
               // }
           //Step 2
               if( vSequenceIterator != "Complete" ) {
                   //Use either Standard Choice WT
                       vWT2( = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
                }
            //Step 3
               if( vSequenceIterator != "Complete" ) {
                   //Use either Standard Choice WT
                       vWT2( = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
                }
                //Step 3
               if( vSequenceIterator != "Complete" ) {
                   //Use either Standard Choice WT
                    vWT2( = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B"))
                   //Define MAOP
                   vnewvMAOP( = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall)))
                   //Test MAOP
                   if( vnewvMAOP >== vMAOP ) {  //(M)
                       vSequenceIterator( = "Complete")
                   }else{
                       vSequenceIterator( = "InComplete")
                    }
                }
        //Stop here!
          // }
          //Output comment if MAOP is never met
          if( vSequenceIterator == "InComplete" And (UnknownSMYS == 1 || UnknownWT == 1 || UnknownWT2 == 1) ) {
              if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" ) {
                  //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                 Range("EV" + CStr(lLastRowdDiameter)).Select
                 with( ActiveCell) {
                    .Characters((Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. ")
                  }
                btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
               }
           }
         }
     } //Date
    Fittings_TimeFrame5 = true
PROC_EXIT()
     function{
PROC_ERR()
    Fittings_TimeFrame5 = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Fittings_TimeFrame5"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function will populate a suggestion in the appropriate field if there is no value there
//           currently.  It will also put a 1 in the rational, fill out the FVE comments and make it blue.
// Returns  Suggestion, Rat//le and Comments
// Created  6/18/2012 example
// --------------------------------------------------------------------------------------------------------------
 function SourceWT_NA()  Boolean{
   //Error Handling (Instances where source W.T. is Blank or N/A) (G)
                    // SMYS when W.T is Invalid
                    // Time Frame 1
                    if( (DateValue(vInstallYr) >== DateValue("4/12/1963")) ) {
                        vSMYS( = 35000)
                    // Time Frame 2
                    else if( (DateValue(vInstallYr) >= DateValue("3/1/1940")) _
                      And (DateValue(vInstallYr) <= DateValue("4/11/1963")) ) {
                        vSMYS( = 30000)
                    // Time Frame 3
                    else if( (DateValue(vInstallYr) >= DateValue("1/1/1930")) _
                      And (DateValue(vInstallYr) <= DateValue("2/28/1940")) ) {
                        vSMYS( = 30000)
                    // Time Frame 4
                    else if( (DateValue(vInstallYr) < DateValue("1/1/1930")) ) {
                        vSMYS( = 24000)
                     }
                    // Unknown Install Date
                    if( Range("F" + CStr(lLastRowdDiameter)).Value == Empty _
                             || Range("F" + CStr(lLastRowdDiameter)).Value = "" _
                             || Range("F" + CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
                             || Range("F" + CStr(lLastRowdDiameter)).Value Like "*unknown*" ) {
                        vSMYS( = 24000)
                     }
                    // Wall Thickness to invalid
                    vWT( = "Invld Src W.T.")
                    Problem7_SourceWT( = 1)
    SourceWT_NA = true
PROC_EXIT()
     function{
PROC_ERR()
    SourceWT_NA = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in SourceWT_NA"
    Resume( PROC_EXIT)
 }
// --------------------------------------------------------------------------------------------------------------
// Comments This function will populate a suggestion in the appropriate field if there is no value there
//           currently.  It will also put a 1 in the rational, fill out the FVE comments and make it blue.
// Returns  Suggestion, Rat//le and Comments
// Created  5/23/2012 example
// --------------------------------------------------------------------------------------------------------------
 function AutoPopulate()  Boolean{
    // Handle Auto-populate for instances where Seam, SMYMS or W.T. are blank
    // if( Seam is Blank (Auto Pop)
    if( ((Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildSeamTypeColumn + CStr(lLastRowdDiameter)).Value == "N/A") _
    And (Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value = 0 || Range("DK" + CStr(lLastRowdDiameter)).Value = Empty || Range("DK" + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range("DK" + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range("DK" + CStr(lLastRowdDiameter)).Value = "N/A")) _
    And Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value != 0 _
    And ! Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value Like "*error*" _
    And ! Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And ! Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value != Empty _
    And DoNotAutoPopSeam = 0 ) {
    //And Problem1_Salvaged != 1 _
    //And Problem2_PriorPurchase != 1 _
    //And Problem5_Seam != 1 _
    //And Problem6_Exceptions != 1
        Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Value = Range(vSuggestedSeamColumn + CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(sUserSeamColumn + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("DN" + CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DN" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
         //Update Catagory and Tier/Category Columns
        Range("EX" + CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("EY" + CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Pipe30Inch( = 1)
        // Make inserted Suggestion, Rational + FVE Comments Value Blue
        if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*SEAM was calculated by PRUPF Logic*" ) {
            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "SEAM was calculated by PRUPF Logic. ")
            Range("EV" + CStr(lLastRowdDiameter)).Select
            with( ActiveCell) {
               .Characters((Len(.Value) + 1).Insert "SEAM was calculated by PRUPF Logic. ")
             }
            btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "SEAM was calculated by PRUPF Logic. ")
            //Selection.Font.Color = vbBlue
            //Selection.Font.Bold = true
         }
     }
    // if( SMYS is Blank (Auto Pop)
    if( ((Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildSMYSColumn + CStr(lLastRowdDiameter)).Value == "N/A") _
    And (Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value = 0 || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value = "" || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value = "N/A")) _
    And Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value != "Field Bend" _
    And Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value != 0 _
    And ! Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value Like "*error*" _
    And ! Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And ! Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value != Empty _
    And DoNotAutoPopSMYS = 0 ) {
    //And Problem1_Salvaged != 1 _
    //And Problem2_PriorPurchase != 1 _
    //And Problem3_SMYS != 1 _
    //And Problem6_Exceptions != 1 ) {
        Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Value = Range(vSuggestedSMYSColumn + CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(vSMYSUserColumn + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("DB" + CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DB" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        //Update Catagory and Tier/Category Columns
        Range("EX" + CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("EY" + CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        SMYS35K( = 1)
        Pipe30Inch( = 1)
        // Make inserted Suggestion, Rational + FVE Comments Value Blue
        if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*SMYS was calculated by PRUPF Logic*" ) {
            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "SMYS was calculated by PRUPF Logic. ")
            Range("EV" + CStr(lLastRowdDiameter)).Select
            with( ActiveCell) {
               .Characters((Len(.Value) + 1).Insert "SMYS was calculated by PRUPF Logic. ")
             }
            btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "SMYS was calculated by PRUPF Logic. ")
            //Selection.Font.Color = vbBlue
            //Selection.Font.Bold = true
         }
     }
    // if( W.T. is Blank (Auto Pop)
    if( ((Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildWTColumn + CStr(lLastRowdDiameter)).Value == "N/A") _
    And (Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value = "" || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value = Empty) || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value = "N/A" || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value != 0 _
    And ! Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value Like "*error*" _
    And ! Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And ! Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value != Empty _
    And DoNotAutoPopWT = 0 ) {
    //And Problem1_Salvaged != 1 _
    //And Problem2_PriorPurchase != 1 _
    //And Problem4_WT != 1 _
    //And Problem6_Exceptions != 1 ) {
        Range(vWTUserColumn + CStr(lLastRowdDiameter)).Value = Range(vSuggestedWTColumn + CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(vWTUserColumn + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("DG" + CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DG" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        //Update Catagory and Tier/Category Columns
        Range("EX" + CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("EY" + CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Pipe30Inch( = 1)
        // Make inserted Suggestion, Rational + FVE Comments Value Blue
        if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*W.T. 1 was calculated by PRUPF Logic*" ) {
            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "W.T. 1 was calculated by PRUPF Logic. ")
            Range("EV" + CStr(lLastRowdDiameter)).Select
            with( ActiveCell) {
               .Characters((Len(.Value) + 1).Insert "W.T. 1 was calculated by PRUPF Logic. ")
             }
            btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "W.T. 1 was calculated by PRUPF Logic. ")
            //Selection.Font.Color = vbBlue
            //Selection.Font.Bold = true
         }
     }
       // if( W.T.2 is Blank (Auto Pop)
    if( ((Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value == Empty || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value == "" || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value Like "*unknown*" || Range(vBuildWT2Column + CStr(lLastRowdDiameter)).Value == "N/A") _
    And (Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value = "" || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value = Empty) || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value = "N/A" || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value Like "*Unknown*" || Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value != 0 _
    And ! Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value Like "*error*" _
    And ! Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And ! Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value != Empty _
    And DoNotAutoPopWT2 = 0 ) {
    //And Problem1_Salvaged != 1 _
    //And Problem2_PriorPurchase != 1 _
    //And Problem4_WT != 1 _
    //And Problem6_Exceptions != 1 ) {
        Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Value = Range(vSuggestedWT2Column + CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(vWT2UserColumn + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("DG" + CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DG" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        //Update Catagory and Tier/Category Columns
        Range("EX" + CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Range("EY" + CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" + CStr(lLastRowdDiameter)).Select
        Selection(.Font.Color = vbBlue)
        Selection.Font.Bold = true
        Pipe30Inch( = 1)
        // Make inserted Suggestion, Rational + FVE Comments Value Blue
        if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*W.T. 2 was calculated by PRUPF Logic*" ) {
            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "W.T. 2 was calculated by PRUPF Logic. ")
            Range("EV" + CStr(lLastRowdDiameter)).Select
            with( ActiveCell) {
               .Characters((Len(.Value) + 1).Insert "W.T. 2 was calculated by PRUPF Logic. ")
             }
            btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "W.T. 2 was calculated by PRUPF Logic. ")
            //Selection.Font.Color = vbBlue
            //Selection.Font.Bold = true
         }
     }
    // if( MAOP doesn//t pass, fill in SME comments
    if( iSMEComments == 1 And ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*W.T. Suggestion is based on maximum values and MAOP did not pass.*" ) {
        //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "W.T. Suggestion is based on maximum values and MAOP did not pass. ")
        Range("EV" + CStr(lLastRowdDiameter)).Select
        with( ActiveCell) {
           .Characters((Len(.Value) + 1).Insert "W.T. Suggestion is based on maximum values and MAOP did not pass. ")
         }
        btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "W.T. Suggestion is based on maximum values and MAOP did not pass. ")
        //Selection.Font.Color = vbBlue
        //Selection.Font.Bold = true
     }
    // if( W.T. Exceeds E.H. W.T., then fill in WT > EH Comments
    if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like "*This W.T. exceeds Extra Heavy W.T. for this diameter.*" _
    And IsNumeric(vWT) = true _
    And( vWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _)
    And (Worksheets("Pipe Data").Range(FeatureColumn + Trim(CStr(lLastRowdDiameter))).Value = "Tee" _
          || (Worksheets("Pipe Data").Range(FeatureColumn + Trim(CStr(lLastRowdDiameter))).Value = "Mfg Bend") _
          || (Worksheets("Pipe Data").Range(FeatureColumn + Trim(CStr(lLastRowdDiameter))).Value = "Reducer") _
          || (Worksheets("Pipe Data").Range(FeatureColumn + Trim(CStr(lLastRowdDiameter))).Value = "Cap")) ) {
        //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "This W.T. exceeds Extra Heavy W.T. for this diameter. ")
        Range("EV" + CStr(lLastRowdDiameter)).Select
        with( ActiveCell) {
           .Characters((Len(.Value) + 1).Insert "This W.T. exceeds Extra Heavy W.T. for this diameter. ")
         }
        btest = ColorText("EV" + CStr(lLastRowdDiameter), "DarkBlue", "This W.T. exceeds Extra Heavy W.T. for this diameter. ")
        //Selection.Font.Color = vbBlue
        //Selection.Font.Bold = true
     }
    //if( SMYS > 35K, "Requires Technical Pier Review", Only for fittings
    if( fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") _
    And( vSMYS > 35000 _)
    And SMYS35K = 1 ) {
        if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*SMYS > 35, Requires Technical Peer Review*") ) {
            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "SMYS > 35, Requires Technical Peer Review. ")
        Range("EV" + CStr(lLastRowdDiameter)).Select
        with( ActiveCell) {
           .Characters((Len(.Value) + 1).Insert UCase("SMYS > 35, Requires Technical Peer Review. "))
         }
        btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "SMYS > 35, Requires Technical Peer Review. ")
            //Selection.Font.Color = vbBlue
            //Selection.Font.Bold = true
         }
     }
//Process 23
    //if( Macro auto populates and the Diameter == 30.5, Requires Tech Peer review
    if( Pipe30Inch == 1 And dDiameter == 30.5 ) {
        if( ! Range("EV" + CStr(lLastRowdDiameter)).Value Like UCase("*Diameter == 30.5, Requires Technical Peer Review*") ) {
            //Range("EV" + CStr(lLastRowdDiameter)).Value = (Range("EV" + CStr(lLastRowdDiameter)).Value + "Diameter = 30.5, Requires Technical Peer Review. ")
            Range("EV" + CStr(lLastRowdDiameter)).Select
            with( ActiveCell) {
               .Characters((Len(.Value) + 1).Insert UCase("Diameter = 30.5, Requires Technical Peer Review. "))
             }
            btest = ColorText("EV" + CStr(lLastRowdDiameter), "purple", "Diameter = 30.5, Requires Technical Peer Review. ")
            //Selection.Font.Color = vbBlue
            //Selection.Font.Bold = true
         }
     }
        
    //Enter all non-logical scenerios here
    if( Worksheets("Pipe Data").Range(FeatureColumn + CStr(lLastRowdDiameter)).Value < 0 ) {
        MsgBox( "No Feature in FVE Section")
        AutoPopulate = false
    }else{
        AutoPopulate = true
     }
PROC_EXIT()
     function{
PROC_ERR()
    AutoPopulate = false
    MsgBox "Error (" + Err.Description + ")", vbExclamation + vbOKOnly, "Error in Auto Populate Function"
    Resume( PROC_EXIT)
 }

 function Worksheet_Change(ByVal Target  Range){

var errMsg  

if( validationStatus == true And ! blankRow(Target.row) ) {
    errMsg( = setValidation(Target))
    if( errMsg != "" ) {
        MsgBox( errMsg)
     }
 }


 }


