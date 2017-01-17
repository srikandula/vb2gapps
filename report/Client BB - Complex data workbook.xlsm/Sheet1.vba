Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'****************************************************************************************************************
'****************************************************************************************************************
'
'    Purpose: The purpose of this function is to provide an automated process in suggesting a "Seam", "SMYS"
'             and "Wall Thickness" for pipe, in addition to "SMYS" and "Wall Thickness" for Tee's, Field Bend's,
'             Manufactured Bends, and Reducers.
'
'    Version: V24 - Phase I
'
'    Modified Date: 12/12/2012 9:20AM
'
'   Log:
'       - Updated OD from 0.5 to 0.54
'       - Update 1.32 O.D. to 1.315" O.D. - 12/4/2012
'       - Moved Seam type Exception Logic up - 12/3/2012
'       - Fixed WT2 Blank or 0 suggestion.  Forced WT2 to calculate each time.
'       - Add "Problem" in suggestioned Columns if Pipe is prior purchase or recondition salvaged
'       - Minor tweaks to address macro forcing a Problem suggestion - 11/30/2012
'       - Added conservative ERW Logic with Test date < 7/1/1961 - 11/29/2012
'       - Added suggestions to be forces if auto-populated by macro - 11/29/2012
'       - Added Cap as a new feature for the macro to analyze (Caps are to be treated the same as Mfg Bends - 11/28/2012
'       - Fixed Fitting "MAOP did not pass" comment - 11/28/2012
'       - Updated Seam Type fix for 4" and less to reflect Unknown SMYS from Build in the Exception section - 11/26/2012
'       - Updated new Default values to Unknown per new template - 11/26/2012
'       - Single Row button to return suggestions regardless if a spec value is known - 11/2/2012
'       - Revised Exception Logic - 10/24/2012
'       - Suggestions will not be returned if known specs for Pipes (this logic is already in place for fittings)
'       - Modified Seam Type Dates - 10/18/2012
'       - Added Unk Install Data W.T. 10/18/2012
'       - Fixed autopopulation. - 10/16/2012
'       - Fixed Seam Type function to handle Prior Purchase = "Unkown"
'
'    Required External Modules / Functions:
'            - PipeDataModule
'                1.) fnxDetermineYield
'                2.) fnxDetermineThickness
'                3.) fnxDetermineThickness1
'                4.) fnxIsItIn
'                5.) fnxTestForNull
'                6.) fnxFindLastRow
'                7.) subHighlightRow
'
'    Internal Functions:
'            - ClearSuggestions
'            - Automation_PFL
'            - SingleSelect_PFL
'            - fnxProcessPipeRow
'                1.) Pipe_Seam()
'                2.) Pipe_SMYS()
'                3.) Pipe_WT()
'                4.) Exceptions()
'                5.) Reconditioned_Salvaged_Pipe()
'                5.) Fittings_Source_WT_1_2 (p_lRowIndex)
'                6.) Fittings_SourceWT()
'                7.) Fittings_TimeFrame1_3()
'                8.) Fittings_TimeFrame4()
'                9.) Fittings_TimeFrame5()
'                10.) Fittings_TimeFrame6()
'                11.) AutoPopulate()
'
'    Output Results:
'            - Pipe Data Tab (3 suggestion columns on FVE Side)
'
'    Steps for fnxProcessPipeRow Function:
'            1.) Enter in acceptable 'types', column references and table locations
'            2.) Begin stages of analysis to determine Seam, SMYS and W.T. for Pipe
'            3.) Begin stages of analysis to determine SMYS and W.T. for Fittngs
'            4.) Obtain Both Source W.T.s for fittings
'            5.) Choose one Source W.T.
'            6.) Time Frame Logic for Fittings
'            7.) Auto populate function
'
'*****************************************************************************************************************
'*****************************************************************************************************************
Option Explicit
Private btest As Boolean
'Begining and ending row variables
Private lLastRowdDiameter As Variant
Private iPipeDataLastRow As Integer
Private iForceSuggestion As Integer
Private iForceSMYS As Integer
Private iForceWT As Integer
Private iForceSeamType As Integer
'Column Variables
Private sFeature As String
Private sSalvaged As String
Private iClassLocation As Variant
Private sPriorOperator As String
Private dDiameter As Double
Private dDiameter2 As Double
Private dDiameterSmall As Double
Private vPurchaseYr As Variant
Private vUserPurchaseYr As Variant
Private vSuggestPurchaseYr As Variant
Private vInstallYr As Variant
Private sSeam As String
Private sUserSeam As String
Private sSuggestSeam As String
Private vSMYS As Variant
Private vUserSMYS As Variant
Private vWT As Variant
Private vWT2 As Variant
Private vUserWT As Variant
Private vUserWT2 As Variant
Private vLsFactor As Variant
Public vSourceWT As Variant
Public vSource2WT As Variant
Public vSourceWT1 As Variant
Public vSourceWT2 As Variant
Public sSourceFeature1 As String
Public sSourceFeature2 As String
Public sSourceFeature3 As String
Public iRowsAway1 As Variant
Public iRowsAway2 As Variant
Public iRowsAway3 As Variant
'Variables subject to change (Adjustment period between Purchase & Install Date)
Private lChangeT As Long
Private lPYearAdjust As Long
'Variables assigned to Columns
Public vBuildWTColumn As Variant
Public vBuildWT2Column As Variant
Public vBuildSMYSColumn As Variant
Public vBuildSeamTypeColumn As Variant
Public FeatureColumn As String
Public iFeatureColumn As Integer
Public SalvagedColumn As String
Public sClassLocationColumn As String
Public sPriorOperatorColumn As String
Public sDiamterColumn As String
Public sDiamterColumn2 As String
Public sUserSeamColumn As String
Public sInstallYrColumn As String
Public sTypeColumn As String
Public vSMYSUserColumn As Variant
Public vWTUserColumn As Variant
Public vWT2UserColumn As Variant
Public vLsFactorColumn As Variant
Public vPurchaseDateColumn As Variant
Public vFittingMAOPColumn As Variant
Public svDesignFactorInstalledClassColumn As String
Public vSuggestedSMYSColumn As Variant
Public vSuggestedWTColumn As Variant
Public vSuggestedWT2Column As Variant
Public vSuggestedSeamColumn As Variant
Public AngleColumn As Variant
'Copy/Paste Function
'Public iTemplateFound As Integer (Not currently in use)
'SME Comments variable
Public iSMEComments As Integer
'Pipe/Fitting Logic (starting Row Logic Tables)
Public iFVEoption As Integer
Public vCurrentLogicStep As Variant
Public vSequenceIterator As Variant
Public itbl2A2 As Integer
Public itbl2B3 As Integer
Public itbl2B4_1 As Integer
Public itbl2B4_2 As Integer
Public itbl2B4_Unk1 As Integer
Public itbl2B4_Unk2 As Integer
Public itbl2B4_3 As Integer
Public itbl2B4_4 As Integer
Public itbl3_2 As Integer
Public itbl4B As Integer
Public itbl5B As Integer
Public iUnknownTbl As Integer
Public iFirstChoice As Integer
Public iSecond_StandardChoice As Integer
Public iThirdChoice As Integer
'MAOP Calculation variables
Public vMAOP As Variant
Public vnewvMAOP As Variant
Public vDesignFactor As Variant
Public vDesignFactorInstalledClass As Variant
'Invalid entry/Problem variables
Public sValidSourceWT As Integer
Public sValidSource2WT As Integer
Public UnknownInstDate As Integer
Public UnknownPurchDate As Integer
Public UnknownSeam As Integer
Public UnknownSMYS As Integer
Public UnknownWT As Integer
Public UnknownWT2 As Integer
Public Problem1_Salvaged As Integer
Public Problem2_PriorPurchase As Integer
Public Problem3_SMYS As Integer
Public Problem4_WT As Integer
Public Problem5_Seam As Integer
Public Problem6_Exceptions As Integer
Public Problem7_SourceWT As Integer
Public Problem8_NewSeam As Integer
Public DoNotAutoPopSeam As Integer
Public DoNotAutoPopSMYS As Integer
Public DoNotAutoPopWT As Integer
Public DoNotAutoPopWT2 As Integer
Public SMYS35K As Integer
Public Pipe30Inch As Integer
Private Const SETTING_NOT_APPLICABLE As Long = 99999
Public ManualException As Integer
Public iProblemAutoPop As Integer
Public sProblemType As String
Public UseOD2 As Integer
' --------------------------------------------------------------------------------------------------------------
' Comments: Clears all suggested data in "Suggestion Columns" only
'           This does not clear suggested values that are inserted into other columns (AutoPopulate Function)
' Created : 11/21/2011 example
' --------------------------------------------------------------------------------------------------------------
Sub ClearSuggestions()
Dim lTotalRows As Long
Dim lColumnMax As Long
Dim lCurrentRow As Long
Dim iColumn As Integer
    'Turn off iForceSuggestion variable
    iForceSuggestion = 0
    'Establish Last Row
    lTotalRows = Worksheets("Pipe Data").Cells(Rows.Count, 102).End(xlUp).row
    'Set Current Row of PFL
    lCurrentRow = 3
On Error GoTo Errorcatch
    'Empty Suggestion Columns
    Worksheets("Pipe Data").Range("DA" & lCurrentRow & ":" & "DA" & lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DD" & lCurrentRow & ":" & "DD" & lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DF" & lCurrentRow & ":" & "DF" & lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DL" & lCurrentRow & ":" & "DL" & lTotalRows) = Empty
PROC_EXIT:
    Exit Sub
Errorcatch:
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Clearing Suggestions"
    Resume PROC_EXIT
End Sub
' --------------------------------------------------------------------------------------------------------------
' Comments: Automation_PFL (AKA All Rows Button) will loop through every row of PFL sheet and execute
'           suggestions on each row
' Returns : Values in suggested fields
' Created : 08/21/2011 example
' --------------------------------------------------------------------------------------------------------------
Sub Automation_PFL()
On Error GoTo PROC_ERR
Dim lCurrentRow As Long
Dim lMod As Long
Dim lTotalRows As Long
    'Row with data starts at Row 3
    lCurrentRow = 3
    'Establish Last Row
    lTotalRows = Worksheets("Pipe Data").Cells(Rows.Count, 102).End(xlUp).row
    'Empty Suggestion Columns
    Worksheets("Pipe Data").Range("DA" & lCurrentRow & ":" & "DA" & lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DD" & lCurrentRow & ":" & "DD" & lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DF" & lCurrentRow & ":" & "DF" & lTotalRows) = Empty
    Worksheets("Pipe Data").Range("DL" & lCurrentRow & ":" & "DL" & lTotalRows) = Empty
    'Update status bar every 10 records
    lMod = 10
    'Turn Screen Off
    Application.ScreenUpdating = False
    'Turns off screen updating
    Application.DisplayStatusBar = True
    'Makes sure that the statusbar is visible
    Application.StatusBar = "Macro Running! (1% Complete)"
    'Loop from Top data row until first blank row (currently using FVE Component Column to measure last row
    Do While Len(Trim(Worksheets("Pipe Data").Range("CX" & Trim(Str(lCurrentRow))).Value)) <> Empty
        If Not fnxProcessPipeRow(lCurrentRow) Then
            'Error if function returns false
            Exit Do
        End If
        'Go to Next Row
        lCurrentRow = lCurrentRow + 1
        If lCurrentRow Mod lMod = 0 Then
            Application.StatusBar = "... " & Str(Int(lCurrentRow / lTotalRows * 100)) & "%"
        End If
    Loop
    'Turn Status bar off upon fishing run.
    Application.StatusBar = False
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Automation_PFL"
    Resume PROC_EXIT
End Sub
' --------------------------------------------------------------------------------------------------------------
' Comments: SingleSelect_PFL (AKA Single Row Button) will execute the macro on the row that is Selected only
' Returns : Values in suggested fields
' Created : 08/15/2011 example
' --------------------------------------------------------------------------------------------------------------
Sub SingleSelect_PFL()
On Error GoTo PROC_ERR
Dim lCurrentRow As Long
Dim lTotalRows As Long
    'Turn on iForceSuggestion variable
    iForceSuggestion = 1
    'Row with data starts at Row
    lCurrentRow = ActiveCell.row
    'Establish Last Row
    lTotalRows = Worksheets("Pipe Data").Cells(Rows.Count, 102).End(xlUp).row
    'Empty Suggestion Columns
    Worksheets("Pipe Data").Range("DA" & lCurrentRow & ":" & "DA" & lCurrentRow) = Empty
    Worksheets("Pipe Data").Range("DD" & lCurrentRow & ":" & "DD" & lCurrentRow) = Empty
    Worksheets("Pipe Data").Range("DF" & lCurrentRow & ":" & "DF" & lCurrentRow) = Empty
    Worksheets("Pipe Data").Range("DL" & lCurrentRow & ":" & "DL" & lCurrentRow) = Empty
    'Added logic to highlig`ht rows in error in red
    If fnxProcessPipeRow(lCurrentRow) Then
        'Row Successfully Processed, turn off red/bold font
        'subHighlightRow "Pipe Data", lCurrentRow, ROW_HIGHLIGHT.Highlight_Off
    Else
        'Error in Row .. highlight as red
        subHighlightRow "Pipe Data", lCurrentRow, ROW_HIGHLIGHT.Highlight_RED
    End If
    'Reset cursor to the FVE section
    Worksheets("Pipe Data").Cells(lCurrentRow, 105).Select
PROC_EXIT:
    Exit Sub
PROC_ERR:
    ' Error message incorrect, use this as stub
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in SingleSelect_PFL"
    Resume PROC_EXIT
End Sub
' --------------------------------------------------------------------------------------------------------------
' Comments: Automaed proccess to return Pipe (Seam, SMYS & Wall Thickness) and Fittings (SMYS & W.T.)
' Returns : Values in suggested fields
' Created : 06/21/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function fnxProcessPipeRow(ByVal p_lRowIndex As Long) As Boolean
On Error GoTo PROC_ERR
Dim iCurrentRow As Long
Dim iRowsAway1 As Variant
Dim iRowsAway2 As Long
    'Use Line No. as indicator of LastRow
    lLastRowdDiameter = p_lRowIndex
    iProblemAutoPop = 0
    ManualException = 0
    '*********************' Column Definitions '*********************'
    If Not ColumnDefinitions() Then
        fnxProcessPipeRow = False
        Exit Function
    End If
    '-----------------------------------------------------------------------------
    '--  1.) Enter in acceptable Feature Types
    '-----------------------------------------------------------------------------
    If fnxIsItIn(sFeature, "Pipe", "Field Bend", "Mfg Bend", "Reducer", "Tee", "Cap") _
    And Range(vFittingMAOPColumn & CStr(lLastRowdDiameter)).Value = "N/A" _
    And Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value <> 0 _
    And Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value <> "N/A" Then
    '----------------------------------------------------------------------------------------
    '--  2.) Begin stages of analysis to determine Seam, SMYS and W.T. for Pipe & Fld Bend --
    '----------------------------------------------------------------------------------------
        If fnxIsItIn(sFeature, "Pipe", "Field Bend") Then
            '*********************' SEAM Analysis '*********************'
            If Not Pipe_Seam() Then
                fnxProcessPipeRow = False
                Exit Function
            End If
            '*********************' SMYS Analysis '*********************'
            If Not Pipe_SMYS() Then
                fnxProcessPipeRow = False
                Exit Function
            End If
            '*********************' W.T. Analysis '*********************'
            If Not Pipe_WT() Then
                fnxProcessPipeRow = False
                Exit Function
            End If
             '***************' Exceptions Analysis '***************'
            If Not Exceptions() Then
                fnxProcessPipeRow = False
                Exit Function
            End If
'-----------------------------------------------------------------------------
'--  3.) Begin stages of analysis to determine SMYS and W.T. for Fittngs    --
'-----------------------------------------------------------------------------
        Else
        'CHOOSE LARGER OD FOR SMYS
            'Currently handles Tee, Field Bends, Mfg Bend and Reducer (W)
            If fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") Then
                'Set Adjustment Dates back to 0 Becuase they should only be delt with pipes
                'Set time change to 0 Years for Fittings
                lChangeT = 0
                'Set adjustment time for Purchase year to 0 Years
                'Set adjustment time for Purchase year to 0 Years
                If vInstallYr >= DateValue("3/1/1940") Then
                    lPYearAdjust = 0
                End If
                'Set Design Factor
                If (Worksheets("Pipe Data").Range("DR" & Trim(Str(p_lRowIndex))).Value = "Error" _
                Or Worksheets("Pipe Data").Range("DR" & Trim(Str(p_lRowIndex))).Value = "N/A" _
                Or Worksheets("Pipe Data").Range("DR" & Trim(Str(p_lRowIndex))).Value = "0") _
                And (Worksheets("Pipe Data").Range("DT" & Trim(Str(p_lRowIndex))).Value = "N/A" _
                Or Worksheets("Pipe Data").Range("DT" & Trim(Str(p_lRowIndex))).Value = "Error" _
                Or Worksheets("Pipe Data").Range("DT" & Trim(Str(p_lRowIndex))).Value = "0") Then
                    vDesignFactor = "Error"
                Else
                     'Process 15.)
                    If Worksheets("Pipe Data").Range(svDesignFactorInstalledClassColumn & Trim(Str(p_lRowIndex))).Value <> "" _
                    And Worksheets("Pipe Data").Range(svDesignFactorInstalledClassColumn & Trim(Str(p_lRowIndex))).Value <> "N/A" _
                    And Worksheets("Pipe Data").Range(svDesignFactorInstalledClassColumn & Trim(Str(p_lRowIndex))).Value <> "ERROR" Then
                        vDesignFactor = Worksheets("Pipe Data").Range("DR" & Trim(Str(p_lRowIndex))).Value
                    Else
                        vDesignFactor = Worksheets("Pipe Data").Range("DT" & Trim(Str(p_lRowIndex))).Value
                    End If
                End If
                'If design factor not valid, insert comments into FVE comments
                If vDesignFactor = "Error" Then
                    Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value = "D.F. N/A"
                    Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value = "D.F. N/A"
                    If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Design Factor Problem.*") Then
                        'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "Design Factor Problem. ")
                        Range("EV" & CStr(lLastRowdDiameter)).Select
                        With ActiveCell
                            .Characters(Len(.Value) + 1).Insert UCase("Design Factor Problem. ")
                        End With
                        btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "Design Factor Problem. ")
                    End If
                Else
'-----------------------------------------------------------------------------
'--  4.) Obtain Both Source W.T.s for fittings                              --
'-----------------------------------------------------------------------------
                '***************' Source W.T. (1&2) Analysis '***************'
                If Not Fittings_Source_WT_1_2(p_lRowIndex) Then
                    fnxProcessPipeRow = False
                    Exit Function
                End If
                'Error Handling (Instances where source W.T. is Blank or N/A) (G)
                If sValidSourceWT = 1 Then
'-----------------------------------------------------------------------------
'--  5.) Choose one Source W.T.                                             --
'-----------------------------------------------------------------------------
                    '***************' Primary Source W.T. Analysis '***************'
                    If Not Fittings_SourceWT() Then
                        fnxProcessPipeRow = False
                        Exit Function
                    End If
                    'Determine Class Location and choose nesign Factor if Class Location = 4
                    If iClassLocation = "4" And (DateValue(vInstallYr) < DateValue("7/1/1961")) Then
                        vDesignFactor = 0.5
                        If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61*") Then
                            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61 ")
                            Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61 ")
                            End With
                            btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "Fitting selected to meet class 3 in a class 4 area - SME review required since 0.4 design factor can not be assumed prior to 7/1/61 ")
                        End If
                    End If
'-----------------------------------------------------------------------------
'--  6.) Time Frame Logic for Fittings                                      --
'-----------------------------------------------------------------------------
                    '*************************' Time Frame 1-3 '*************************'
                    If Not Fittings_TimeFrame1_3() Then
                        fnxProcessPipeRow = False
                        Exit Function
                    End If
                    'Establish Variables and formulas for Time Frames (4-6):
                      'Set MAOP From MAOP of R Field (MAOP of R will default to 0 if no values are entered in by user)
                    vMAOP = Worksheets("Pipe Data").Range("EK" & Trim(Str(p_lRowIndex))).Value
                    'Default Current Logic
                    vCurrentLogicStep = SETTING_NOT_APPLICABLE
                    'Default Iterator to stage 1 very first instance
                    vSequenceIterator = "Start"
                    'Error Handling for SMYS
                    If vSMYS Like "Problem" Then
                        vSMYS = 0
                    End If
                    '**************************' Time Frame 4 '**************************'
                    'Time Frame 4: 3/1/1940 - 4/11/1963
                    If Not Fittings_TimeFrame4() Then
                        fnxProcessPipeRow = False
                        Exit Function
                    End If
                    '**************************' Time Frame 5 '**************************'
                    'Time Frame 5: 4/12/1963 - 10/31/1968
                    If Not Fittings_TimeFrame5() Then
                        fnxProcessPipeRow = False
                        Exit Function
                    End If
                    '**************************' Time Frame 6 '**************************'
                    'Time Frame 6: >= 11/1/1968
                    'If Not Fittings_TimeFrame6() Then
                    '    fnxProcessPipeRow = False
                    '    Exit Function
                    'End If
                Else
                    '**************************' Source W.T. Not Available '**************************'
                    If Not SourceWT_NA() Then
                        fnxProcessPipeRow = False
                        Exit Function
                    End If
                End If
                'See Logic Tables (PRUF should output NA for 0.54 Diameter
                If vWT = 0 Or dDiameter = 0.54 Then
                    vWT = "N/A"
                End If
            End If
            ' (W) If Worksheets("Pipe Data").Range("H" & Trim(Str(p_lRowIndex))).Value = "Tee"
            End If 'Ends handling of fittings (W)
        End If 'Ends Pipe/Other Handler (A)
        '***************' Salvaged Analysis '***************'
        If Not Salvaged() Then
            fnxProcessPipeRow = False
            Exit Function
        End If
        '***************' Prior Purchase Analysis '***************'
        If Not PriorPurchase() Then
            fnxProcessPipeRow = False
            Exit Function
        End If
        '***************' Problems Analysis '***************'
        If Not Problems() Then
            fnxProcessPipeRow = False
            Exit Function
        End If
          'Output suggestions Suggestion fields
        If fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") Then
            'Populate SMYS
            If UnknownSMYS = 1 Or iForceSuggestion = 1 Or iForceSMYS = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                If (Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 _
                Or Problem3_SMYS = 1 Or Problem7_SourceWT = 1) Then
                    'If Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                        Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopSMYS = 1
                    'End If
                Else
                    If iForceSuggestion = 1 Then
                        DoNotAutoPopSMYS = 1
                    End If
                    If Problem1_Salvaged <> 1 And Problem2_PriorPurchase <> 1 Then
                        Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value = vSMYS
                    End If
                End If
            End If
            'Populate WT
            If UnknownWT = 1 Or iForceSuggestion = 1 Or iForceWT = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                If (Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Or Problem4_WT = 1 Or Problem7_SourceWT = 1) Then
                    'If Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                        Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopWT = 1
                    'End If
                Else
                    If iForceSuggestion = 1 Then
                        DoNotAutoPopWT = 1
                    End If
                    If Problem1_Salvaged <> 1 And Problem2_PriorPurchase <> 1 Then
                        If UseOD2 = 1 Then
                            Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value = vWT2
                        Else
                            Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value = vWT
                        End If
                    End If
                End If
            End If
             'Populate WT2
            If (UnknownWT2 = 1 Or iForceSuggestion = 1 Or iForceWT = 1) And Range("DI" & CStr(lLastRowdDiameter)).Value <> "N/A" Then
                If (Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Or Problem4_WT = 1 Or Problem7_SourceWT = 1) Then
                    'If Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                        Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopWT2 = 1
                    'End If
                Else
                    If iForceSuggestion = 1 Then
                        DoNotAutoPopWT2 = 1
                    End If
                    If Problem1_Salvaged <> 1 And Problem2_PriorPurchase <> 1 Then
                        If UseOD2 = 1 Then
                            Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value = vWT
                        Else
                            Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value = vWT2
                        End If
                    End If
                End If
            End If
        ElseIf fnxIsItIn(sFeature, "Pipe", "Field Bend") Then
            'SMYS
            If UnknownSMYS = 1 Or iForceSuggestion = 1 Or iForceSMYS = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                If (ManualException = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Or Problem3_SMYS = 1 Or Problem6_Exceptions = 1) Then
                    'If Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                        Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopSMYS = 1
                    'End If
                Else
                    If iForceSuggestion = 1 Then
                        DoNotAutoPopSMYS = 1
                    End If
                    If Problem1_Salvaged <> 1 And Problem2_PriorPurchase <> 1 Then
                        Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value = vSMYS
                    End If
                End If
            End If
            'W.T.
            If UnknownWT = 1 Or iForceSuggestion = 1 Or iForceWT = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                If (ManualException = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Or Problem4_WT = 1 Or Problem6_Exceptions = 1) Then
                    'If Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                        Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopWT = 1
                    'End If
                Else
                    If iForceSuggestion = 1 Then
                        DoNotAutoPopWT = 1
                    End If
                    If Problem1_Salvaged <> 1 And Problem2_PriorPurchase <> 1 Then
                        Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value = vWT
                    End If
                End If
            End If
            'SEAM TYPE
            If UnknownSeam = 1 Or iForceSuggestion = 1 Or iForceSeamType = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                If (ManualException = 1 Or Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Or Problem5_Seam = 1 Or Problem6_Exceptions = 1) Then
                    'If Problem1_Salvaged = 1 Or Problem2_PriorPurchase = 1 Then
                        Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value = sProblemType
                        DoNotAutoPopSeam = 1
                    'End If
                Else
                    If iForceSuggestion = 1 Then
                        DoNotAutoPopSeam = 1
                    End If
                    If Problem1_Salvaged <> 1 And Problem2_PriorPurchase <> 1 Then
                        Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value = sSuggestSeam
                    End If
                End If
            End If
        End If
'-----------------------------------------------------------------------------
'--  7.) Auto populate function                                             --
'-----------------------------------------------------------------------------
        ' Handle Auto-populate for instances where Seam, SMYMS or W.T. are blank
        'If (Problem3_SMYS = 3 Or Problem4_WT = 4) Then
            If Not AutoPopulate() Then
                fnxProcessPipeRow = False
                Exit Function
            End If
        'End If
    End If
'**************************************************************************************************************'
    'Indicate True (code executed without errors)
    fnxProcessPipeRow = True
'Handle All Error Descriptions
PROC_EXIT:
    Exit Function
PROC_ERR:
    fnxProcessPipeRow = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in fnxProcessPipeRow"
    Resume PROC_EXIT
End Function
'**************************************************************************************************************'
'**************************************************************************************************************'
' -- Function Definitions --
'**************************************************************************************************************'
'**************************************************************************************************************'
' --------------------------------------------------------------------------------------------------------------
' Comments: This function defines column positions and variables for input data
' Returns : Column locations
' Created : 06/14/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function ColumnDefinitions() As Boolean
    'Last row in Sheet based on Feature Number in FVE
    iPipeDataLastRow = Worksheets("Pipe Data").Cells(Rows.Count, 102).End(xlUp).row
'Set Columns Here
    'Build WT
    vBuildWTColumn = "P"
    'Build WT2
    vBuildWT2Column = "R"
    'Build SMYS
    vBuildSMYSColumn = "U"
    'Build SEAM
    vBuildSeamTypeColumn = "S"
    'PriorPurchaser
    sPriorOperatorColumn = "CY"
    'Diameter
    sDiamterColumn = "DH"
    'Diameter
    sDiamterColumn2 = "DI"
    'Seam User
    sUserSeamColumn = "DK"
    'SMYS User
    vSMYSUserColumn = "CZ"
    'WT User
    vWTUserColumn = "DC"
    'WT User2
    vWT2UserColumn = "DE"
    'Install Year
    sInstallYrColumn = "DZ"
    'Long Seam Factor
    vLsFactorColumn = "DM"
    'Class Location
    sClassLocationColumn = "DS"
    'Angle Column for Mfg Bends
    AngleColumn = "AD"
    'Reconditioned / Salvaged?
    SalvagedColumn = "AZ"
    'Purchase Date of Feature
    vPurchaseDateColumn = "AV"
    'Fitting MAOP
    vFittingMAOPColumn = "DY"
    'Feature
    FeatureColumn = "CX"
    'Installed CL Design Factor
    svDesignFactorInstalledClassColumn = "DR"
    'Feature Column Number
    iFeatureColumn = 102
    'Suggested SMYS
    vSuggestedSMYSColumn = "DA"
    'Suggested WT
    vSuggestedWTColumn = "DD"
    'Suggested WT2
    vSuggestedWT2Column = "DF"
    'Suggested Seam
    vSuggestedSeamColumn = "DL"
    'Set Feature Column
    sFeature = Range(FeatureColumn & CStr(lLastRowdDiameter)).Value
    'Set Salvaged Column
    sSalvaged = Range(SalvagedColumn & CStr(lLastRowdDiameter)).Value
    'Set SME comments variable to 0 by default
    iSMEComments = 0
    'Set time change to +10 Years
    lChangeT = 10
    'Set adjustment time for Purchase year to -10 Years
    lPYearAdjust = -10
    'Check to ensure users only enter approved values in Col sPriorOperatorColumn
    'Set PriorPurchase to No if not already specified
    If Range(sPriorOperatorColumn & CStr(lLastRowdDiameter)).Value <> Empty Then
        sPriorOperator = Range(sPriorOperatorColumn & CStr(lLastRowdDiameter)).Value
    Else
        Range(sPriorOperatorColumn & CStr(lLastRowdDiameter)).Value = "No"
        sPriorOperator = "No"
    End If
    'Set dDiameter equal to last row in dDiameter Column
    If Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value = "PipeOD+(2*SleeveWT)+0.25" Or Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value = 0 Or Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value = "N/A" Or Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Then
        dDiameter = 0
    Else
        dDiameter = Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value
    End If
    'Set dDiameter2 equal to last row in dDiameter Column
    If Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value = "PipeOD+(2*SleeveWT)+0.25" Or Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value = 0 Or Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value = "N/A" Or Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value Like "*unknown*" Then
        dDiameter2 = 0
    Else
        dDiameter2 = Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value
    End If
    'Pull in User Specified Seam
    sUserSeam = Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value
    'Pull in User Specified SMYS
    vUserSMYS = Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value
    'Pull in User Specified W.T.
    vUserWT = Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value
    'Pull in User Specified W.T.2
    vUserWT2 = Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value
    'Long Seam Factor
    vLsFactor = Range(vLsFactorColumn & CStr(lLastRowdDiameter)).Value
    'Class Location
    iClassLocation = Range(sClassLocationColumn & CStr(lLastRowdDiameter)).Value
     'Choose Starting row for the following variables (Look at "Logic" Tab)
       'Note: For all "tbl..." variables - Starting Row will be the first of the two Date rows at the top of the table.
       'Note: For the remaining 4 variables, set the starting row to one row above the title (2 rows above where the data starts).
    itbl2A2 = 12
    itbl2B3 = 18
    itbl2B4_1 = 24
    itbl2B4_2 = 42
    itbl2B4_Unk1 = 66
    itbl2B4_Unk2 = 77
    itbl2B4_3 = 92
    itbl2B4_4 = 112
    itbl3_2 = 156
    itbl4B = 196
    itbl5B = 201
    iUnknownTbl = 220
    iFirstChoice = 247
    iSecond_StandardChoice = 274
    iThirdChoice = 301
    'Default all Problem Variables to 0
    Problem1_Salvaged = 0
    Problem2_PriorPurchase = 0
    Problem3_SMYS = 0
    Problem4_WT = 0
    Problem5_Seam = 0
    Problem6_Exceptions = 0
    Problem7_SourceWT = 0
    Problem8_NewSeam = 0
    DoNotAutoPopSeam = 0
    DoNotAutoPopSMYS = 0
    DoNotAutoPopWT = 0
    DoNotAutoPopWT2 = 0
    SMYS35K = 0
    Pipe30Inch = 0
    UnknownSeam = 0
    UnknownSMYS = 0
    UnknownWT = 0
    UnknownWT2 = 0
    UnknownInstDate = 0
    UnknownPurchDate = 0
    'Default Manual Exception to 0
    ManualException = 0
    UseOD2 = 0
    'Turn off iForceSuggestion variable unless there is a 1 in the Rat'le columns
    iForceSuggestion = 0
    If Worksheets("Pipe Data").Range("DB" & Trim(Str(lLastRowdDiameter))).Value = 1 Then
        iForceSMYS = 1
    Else
        iForceSMYS = 0
    End If
    If Worksheets("Pipe Data").Range("DG" & Trim(Str(lLastRowdDiameter))).Value = 1 Then
        iForceWT = 1
    Else
        iForceWT = 0
    End If
    If Worksheets("Pipe Data").Range("DN" & Trim(Str(lLastRowdDiameter))).Value = 1 Then
        iForceSeamType = 1
    Else
        iForceSeamType = 0
    End If
    'vSequenceIterator = 0
    'Check if Seam Type is Unknown
    If fnxIsItIn(sFeature, "Pipe", "Field Bend") And Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value <> "" And Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value <> "NA" And Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value <> "N/A" Then
        vInstallYr = DateAdd("d", 1, (DateAdd("yyyy", 10, Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value)))
    Else
        UnknownPurchDate = 1
        If Range(sInstallYrColumn & CStr(lLastRowdDiameter)).Value = Empty _
        Or Range(sInstallYrColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
        Or Range(sInstallYrColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Then
            UnknownInstDate = 1
            vInstallYr = DateValue("1/1/1930")
        Else
            vInstallYr = Range(sInstallYrColumn & CStr(lLastRowdDiameter)).Value
        End If
    End If
    'End If
    'Set Purchase date to whatever Install date is minus 10 years (-10 yrs + 1 day)
'Process 16.) SMYS and W.T.
    If Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value <> "" And Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value <> "NA" Then
        vPurchaseYr = Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value
    Else
        UnknownPurchDate = 1
        vPurchaseYr = DateAdd("d", 1, (DateAdd("yyyy", lPYearAdjust, vInstallYr)))
    End If
    If ((Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value Like "*N/A*") _
    And (Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value = 0 Or Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value Like "*N/A*")) Then
        UnknownSeam = 1
    Else
        UnknownSeam = 0
    End If
    'Check if SMYS is Unknown for Step 1 of fitting logic
    If ((Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value = "N/A") _
    And (Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value = 0 Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value = "N/A")) Then
        UnknownSMYS = 1
    Else
        UnknownSMYS = 0
    End If
    'Check if W.T. is Unknown for Step 1 of fitting logic
    If ((Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value = "N/A") _
    And (Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value = Empty) Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value = "N/A" Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*") Then
        UnknownWT = 1
    Else
        UnknownWT = 0
    End If
    'Check if W.T.2 is Unknown for Step 1 of fitting logic
    If ((Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value = "N/A")) _
    And ((Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value = "N/A" Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*")) _
    And (Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value <> "" And Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value <> Empty And Range(sDiamterColumn2 & CStr(lLastRowdDiameter)).Value <> "N/A" And Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value <> "0") Then
        UnknownWT2 = 1
    Else
        UnknownWT2 = 0
    End If
    ColumnDefinitions = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    ColumnDefinitions = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Column Definitions"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function returns a suggested Seam for Pipes
' Returns : Pipe - Seam
' Created : 02/27/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function Pipe_Seam() As Boolean
    'Check if there is a Prior Operator (B)
    If sPriorOperator = "Yes" And fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 3.5, 4.5) Then
        sSuggestSeam = "Furnace Butt Weld"
    ElseIf sPriorOperator = "Yes" And fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, _
      12.75, 14, 16, 18, 20, 22, _
      24, 26, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42) Then
        sSuggestSeam = "Lap Weld"
    'All cases with no prior operator - 'Begin Seam calculation logic for Pipe
    ElseIf fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375) And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") Then
        sSuggestSeam = "Furnace Butt Weld"
    ElseIf dDiameter = 3.5 And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) <= DateValue("2/12/1983") Then
        sSuggestSeam = "Furnace Butt Weld"
    ElseIf dDiameter = 3.5 And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) >= DateValue("2/13/1983") Then
        sSuggestSeam = "Seamless/Electric Resistance Weld"
    ElseIf dDiameter = 4.5 And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) <= DateValue("10/25/1977") Then
        sSuggestSeam = "Furnace Butt Weld"
    ElseIf dDiameter = 4.5 And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) >= DateValue("10/26/1977") Then
        sSuggestSeam = "Seamless/Electric Resistance Weld"
    ElseIf fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, 12.75) _
      And (sPriorOperator = "No" Or sPriorOperator = "N/A") _
      And DateValue(vInstallYr) <= DateValue("12/30/1940") Then
        sSuggestSeam = "Lap Weld"
    ElseIf (fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, 12.75) _
      And (sPriorOperator = "No" Or sPriorOperator = "N/A") _
      And DateValue(vInstallYr) >= DateValue("12/31/1940")) Then
        sSuggestSeam = "Seamless/Electric Resistance Weld"
    '14" Diameter
    ElseIf dDiameter = "14" Then
        sSuggestSeam = "Seamless/Electric Resistance Weld"
    '16" Diameter
    ElseIf dDiameter = "16" _
      And (sPriorOperator = "No" Or sPriorOperator = "N/A") _
      And DateValue(vInstallYr) <= DateValue("12/30/1958") Then
        sSuggestSeam = "AO Smith"
    ElseIf dDiameter = "16" _
      And (sPriorOperator = "No" Or sPriorOperator = "N/A") _
      And (DateValue(vInstallYr) >= DateValue("12/31/1958")) Then
        sSuggestSeam = "Seamless/Electric Resistance Weld"
    '18" Diameter
    ElseIf dDiameter = "18" Then
        sSuggestSeam = "Seamless/Electric Resistance Weld/Double Submerged Arc Weld"
    '20-24 & 26 Diameter
    ElseIf fnxIsItIn(dDiameter, 20, 22, 24, 26) And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") _
      And DateValue(vInstallYr) <= DateValue("12/30/1958") Then
        sSuggestSeam = "Single Submerged Arc Weld/AO Smith"
    ElseIf fnxIsItIn(dDiameter, 20, 22, 24) And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") _
      And (DateValue(vInstallYr) >= DateValue("12/31/1958")) Then
        sSuggestSeam = "Seamless/Double Submerged Arc Weld"
    '26"
    ElseIf fnxIsItIn(dDiameter, 26) And (sPriorOperator = "No" Or sPriorOperator = "N/A" Or sPriorOperator = "Unknown") _
      And (DateValue(vInstallYr) >= DateValue("12/31/1958")) Then
        sSuggestSeam = "Double Submerged Arc Weld"
    ElseIf fnxIsItIn(dDiameter, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42) _
      And (sPriorOperator = "No" Or sPriorOperator = "N/A") Then
        sSuggestSeam = "Double Submerged Arc Weld"
    Else
        sSuggestSeam = "Problem 5"
        Problem5_Seam = 1
    End If 'Ends handling of Seam for Pipe (B)
    'Choose ERW if Seamless/ERW and class location = 1 and install date < 1961
     If Range("DQ" & CStr(lLastRowdDiameter)).Value Like "*1*" Or ((Range("DQ" & CStr(lLastRowdDiameter)).Value = "" Or Range("DQ" & CStr(lLastRowdDiameter)).Value = "N/A" Or Range("DQ" & CStr(lLastRowdDiameter)).Value = "Unknown") And iClassLocation Like "*1*") Then
        If Range("EA" & CStr(lLastRowdDiameter)).Value < DateValue("7/1/1961") And Range("EA" & CStr(lLastRowdDiameter)).Value <> "N/A" And sSuggestSeam = "Seamless/Electric Resistance Weld" Then
            sSuggestSeam = "Electric Resistance Weld"
            If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Most conservative seam type is ERW due to possibility of low frequency ERW in Class 1 location and a test pressure ratio of < 1.25:1.*" Then
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert "Most conservative seam type is ERW due to possibility of low frequency ERW in Class 1 location and a test pressure ratio of < 1.25:1.  "
                End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "Most conservative seam type is ERW due to possibility of low frequency ERW in Class 1 location and a test pressure ratio of < 1.25:1.  ")
            End If
        End If
    End If
    
        'Exceptions for Seam Types ***********************************************
                If dDiameter < 5 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 400 _
            And DateValue(vInstallYr) >= DateValue("10/13/1964") Then
                If sUserSeam = "Furnace Butt Weld" Then
                    If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Furnace Butt Weld with MAOP > 400 psig, SME Review Required.  *") Then
                        Range("EV" & CStr(lLastRowdDiameter)).Select
                                With ActiveCell
                                    .Characters(Len(.Value) + 1).Insert UCase("Furnace Butt Weld with MAOP > 400 psig, SME Review Required.  ")
                                End With
                        btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "Furnace Butt Weld with MAOP > 400 psig, SME Review Required.  ")
                    End If
                End If
                'Force Seam Type
                If dDiameter < 3 Then
                    sUserSeam = "Seamless"
                    sSeam = "Seamless"
                    sSuggestSeam = "Seamless"
                Else
                    sUserSeam = "Seamless/Electric Resistance Weld"
                    sSeam = "Seamless/Electric Resistance Weld"
                    sSuggestSeam = "Seamless/Electric Resistance Weld"
                End If
            End If
            'Exceptions for Steam Types (SMYS):
            If vUserSMYS > 30000 And vUserSMYS <> "Unknown" And vUserSMYS <> "unknown" And Not (vSMYS Like "*Problem*") And dDiameter < 5 Then
                If dDiameter < 3 Then
                    sUserSeam = "Seamless"
                    sSeam = "Seamless"
                    sSuggestSeam = "Seamless"
                Else
                    sUserSeam = "Seamless/Electric Resistance Weld"
                    sSeam = "Seamless/Electric Resistance Weld"
                    sSuggestSeam = "Seamless/Electric Resistance Weld"
                End If
            End If
            'A53
            If sSeam = "Furnace Butt Weld" And Range("T" & CStr(lLastRowdDiameter)).Value = "ASTM A-53" And dDiameter <= 4 Then
                vSMYS = 30000
            End If
            'New Seam Types:
            If fnxIsItIn(sUserSeam, "Spiral Weld post 1966", "Polyethylene Pipe", "Special 0.95", "Special 0.90", "Special 0.85") Then
                Problem8_NewSeam = 1
                vWT = "Problem 8"
                sSeam = "Problem 8"
                vSMYS = "Problem 8"
            End If
    
    'Choose Users inputed seam (if there is one) for further SMYS and W.T. calculations
    'If Range("S" & CStr(lLastRowdDiameter)).Value = Empty Or Range("S" & CStr(lLastRowdDiameter)).Value Like "*Unknown*" _

    If (Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value = 0 Or Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or IsEmpty(Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value) = True) Then
       sSeam = sSuggestSeam
    Else
        'Set Seam type to User Seam Type
        sSeam = sUserSeam
        'Exceptions for new seam types
        If sUserSeam = "Electric Fusion Weld" Then
            sSeam = "Double Submerged Arc Weld"
        End If
    End If
    Pipe_Seam = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Pipe_Seam = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Pipe-Seam"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function returns a suggested SMYS for Pipes
' Returns : Pipe - SMYS
' Created : 02/27/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function Pipe_SMYS() As Boolean
    'Begin SMYS calculation logic for Pipe (C)
    'Determine if Purchase date is known
    If dDiameter = 3.5 _
      And (sSeam = "Furnace Butt Weld" Or sSeam Like "*Unknown*") Then
        vSMYS = fnxDetermineYield("Logic", itbl2A2, 12, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    ElseIf dDiameter = 4.5 _
      And (sSeam = "Furnace Butt Weld" Or sSeam Like "*Unknown*") Then
        vSMYS = fnxDetermineYield("Logic", itbl2B3, 12, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    'Unknown Date
    ElseIf (Range(sInstallYrColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range("F" & CStr(lLastRowdDiameter)).Value Like "*Unknown*") _
      And fnxIsItIn(dDiameter, 2.375, 3.5) And (sSeam = sSeam Like "*Unknown*") Then
        vSMYS = 25000
    'NEW TABLE
    ElseIf fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375) _
      And (sSeam = "Furnace Butt Weld" Or sSeam Like "*Unknown*") Then
        vSMYS = fnxDetermineYield("Logic", itbl2B4_1, 4, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    ElseIf fnxIsItIn(dDiameter, 0.54, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 3.5, 4.5, 6.625, 8.625) _
      And sSeam <> "Furnace Butt Weld" _
      And (sSeam <> "Lap Weld") Then ' And dDiameter <> 6.625 Or dDiameter <> 8.625) Then 'Confirm with Jim
        vSMYS = fnxDetermineYield("Logic", itbl2B4_2, 11, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    ElseIf fnxIsItIn(dDiameter, 4.5, 6.625, 8.625, 10.75, 12.75) And sSeam = "Lap Weld" Then
        vSMYS = fnxDetermineYield("Logic", itbl2B4_Unk1, 5, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    ElseIf fnxIsItIn(dDiameter, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26) And sSeam Like "*Unknown*" Then
        vSMYS = fnxDetermineYield("Logic", itbl2B4_Unk2, 12, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    ElseIf fnxIsItIn(dDiameter, 12.75, 16, 20, 22, 24, 26) _
      And (sSeam = "Single Submerged Arc Weld" Or sSeam = "AO Smith" Or sSeam = "Single Submerged Arc Weld/AO Smith") Then
        vSMYS = fnxDetermineYield("Logic", itbl2B4_3, 5, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    ElseIf dDiameter >= 10.75 And fnxIsItIn(sSeam, "Double Submerged Arc Weld", "Electric Resistance Weld", "Seamless", _
        "Seamless/Double Submerged Arc Weld", "Seamless/Electric Resistance Weld", "Seamless/Electric Resistance Weld/Double Submerged Arc Weld") Then
        vSMYS = fnxDetermineYield("Logic", itbl2B4_4, 11, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    Else
    'Default to "Problem" if SMYS is not captures in any of the prior logic sets
        vSMYS = "Problem 3"
        Problem3_SMYS = 1
    End If 'Ends handling of SMYS for Pipe (C)
    'Handle when the Year is not specified. This will overide previous logic regardless.
    If dDiameter <= 3.5 _
      And sSeam = "Furnace Butt Weld" And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vSMYS = 25000
    End If 'Ends handling of SMYS for Pipe with no Date specified
    If vSMYS = 0 Then
        vSMYS = "Problem 3"
        Problem3_SMYS = 1
    Else
        vSMYS = vSMYS
    End If
    Pipe_SMYS = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Pipe_SMYS = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Pipe-SMYS"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function returns a suggested Wall Thickness for Pipes
' Returns : Pipe - W.T.
' Created : 08/21/2011 example
' --------------------------------------------------------------------------------------------------------------
Private Function Pipe_WT() As Boolean
    'Begin W.T. calculation logic for Pipe
    'Handle Unknown Install Dates
    If dDiameter = 3.5 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then   '(D)
        vWT = 0.141
    ElseIf dDiameter = 4.5 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        If sSeam = "Seamless" Or sSeam = "Furnace Butt Weld" Then
            vWT = 0.148
        Else
            vWT = 0.141
        End If
    ElseIf dDiameter = 6.625 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.156
    ElseIf dDiameter = 8.625 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.172
    ElseIf dDiameter = 10.75 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.188
    ElseIf dDiameter = 12.75 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.203
    ElseIf dDiameter = 16 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.219
    ElseIf dDiameter = 18 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.25
    ElseIf dDiameter = 20 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.25
    ElseIf dDiameter = 22 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.25
    ElseIf dDiameter = 24 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.25
    ElseIf dDiameter = 26 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.281
    ElseIf dDiameter = 30 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.25
    ElseIf dDiameter = 32 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.375
    ElseIf dDiameter = 34 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.25
    ElseIf dDiameter = 36 And (UnknownInstDate <> 0 And UnknownPurchDate <> 0) Then
        vWT = 0.312
    Else
        vWT = fnxDetermineThickness("Logic", itbl3_2, 9, dDiameter, sSeam, vPurchaseYr, UnknownPurchDate)
    End If 'Ends handling of SMYS for Pipe (D)
    If vWT = 0 Or vWT = "error" Then
        vWT = "Problem 4"
        Problem4_WT = 1
    Else
        vWT = vWT
    End If
    'End If
    Pipe_WT = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Pipe_WT = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Pipe-W.T."
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function handles all of the exceptions for pipe (4", 6", 8")
' Returns : Pipe - Exceptions
' Created : 10/24/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function Exceptions() As Boolean
            'Special Case for 4" with 970 MAOP
            If dDiameter = 4.5 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 970 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("6/14/1962") And DateValue(vPurchaseYr) <= DateValue("10/12/1964") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("6/14/1962") And DateValue(vInstallYr) <= DateValue("10/11/1974")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
            'Special Case for 4" with 678 MAOP
            If dDiameter = 4.5 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 678 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("1/1/1930") And DateValue(vPurchaseYr) <= DateValue("12/31/1930") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("1/1/1930") And DateValue(vInstallYr) <= DateValue("12/30/1940")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
            'Special Case for 4" with 665 MAOP
            If dDiameter = 4.5 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 665 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("6/14/1962") And DateValue(vPurchaseYr) <= DateValue("10/12/1964") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("6/14/1962") And DateValue(vInstallYr) <= DateValue("10/11/1974")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
            'Special Case for 6" with 489 MAOP
            If dDiameter = 6.625 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 489 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) = DateValue("6/17/1948") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("6/17/1948") And DateValue(vInstallYr) <= DateValue("6/16/1958")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
            'Special Case for 8" with 451 MAOP
            If dDiameter = 8.625 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 451 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) = DateValue("12/21/1945") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("12/21/1945") And DateValue(vInstallYr) <= DateValue("12/20/1955")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
            'Special Case for 8" with 1045 MAOP
            If dDiameter = 8.625 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 1045 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) <= DateValue("8/22/1932") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) <= DateValue("8/21/1942")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
            'Special Case for 12" with 420 MAOP
            If dDiameter = 12.75 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 420 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) = DateValue("8/7/1941") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("8/7/1941") And DateValue(vInstallYr) <= DateValue("8/6/1951")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
            'Special Case for 16" with 590 MAOP
            If dDiameter = 16 _
            And Range("EK" & CStr(lLastRowdDiameter)).Value > 590 _
            And (UnknownPurchDate = 0 And DateValue(vPurchaseYr) >= DateValue("1/1/1953") And DateValue(vPurchaseYr) <= DateValue("12/31/1954") _
            Or (UnknownPurchDate = 1 And UnknownInstDate = 0 And DateValue(vInstallYr) >= DateValue("1/1/1953") And DateValue(vInstallYr) <= DateValue("12/30/1964")) _
            Or (UnknownInstDate <> 0 And UnknownPurchDate <> 0)) Then
                If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.*") Then
                    Range("EV" & CStr(lLastRowdDiameter)).Select
                            With ActiveCell
                                .Characters(Len(.Value) + 1).Insert UCase("|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                            End With
                    btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|6| SME consider exception specs in PRUPF Table 3 which do not support MAOP-R.  ")
                End If
            End If
'??? Jim
            '18" Diameter
            'If dDiameter = 18 Then
            '    vSMYS = "Problem"
            '    sSuggestSeam = "Problem"
            '    vWT = "Problem"
            '    If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Exception for Diameter*" Then
            '        'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|6| Exception for Diameter " & dDiameter & ", See PRUPF Table 3 notes.  ")
            '        Range("EV" & CStr(lLastRowdDiameter)).Select
            '                With ActiveCell
            '                    .Characters(Len(.Value) + 1).Insert "|6| Exception for Diameter " & dDiameter & ", See PRUPF Table 2 notes.  "
            '                End With
            '        btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "|6| Exception for Diameter " & dDiameter & ", See PRUPF Table 2 notes.  ")
            '        'Selection.Font.Color = vbBlue
            '        'Selection.Font.Bold = True
            '    End If
            'End If
            'Exceptions for Steam Types (MAOP-R):
    Exceptions = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Exceptions = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Exceptions Function"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function Handles actions for the Reconditioned/Salvaged Column
' Returns : "Problem" in suggestion fields based on given criteria
' Created : 07/24/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function Salvaged() As Boolean
'If Salvaged column is anything but yes (No or Unknown) and Feature is Pipe, Bend (or Tee, Reducer, Mfg Bend & Date < 3/1/1940)
'Then macro should return Problem
    If (sSalvaged = "Yes" Or sSalvaged = "yes") _
    And ((fnxIsItIn(sFeature, "Pipe", "Field Bend") _
    Or (fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") And (DateValue(vInstallYr) < DateValue("3/1/1940"))))) Then
        Problem1_Salvaged = 1
        DoNotAutoPopSeam = 1
        DoNotAutoPopWT = 1
        DoNotAutoPopSMYS = 1
        DoNotAutoPopWT2 = 1
    'ElseIf (sSalvaged = "" Or sSalvaged = "unknown" Or sSalvaged = "Unknown" Or IsEmpty(sSalvaged) = True) _
    'And ((fnxIsItIn(sFeature, "Pipe", "Field Bend"))) Then
    '    If Not Range("EV" & CStr(lLastRowdDiameter)).value Like "*Reconditioned/Salvage is marked as Unknown, requires technical peer review. *" Then
    '        Range("EV" & CStr(lLastRowdDiameter)).value = (Range("EV" & CStr(lLastRowdDiameter)).value + "Reconditioned/Salvage is marked as Unknown, requires technical peer review.  ")
    '        Worksheets("Pipe Data").Range("EV" & CStr(lLastRowdDiameter)).Select
    '        Selection.Font.Color = vbBlue
    '        Selection.Font.Bold = True
    '    End If
    End If
    Salvaged = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Salvaged = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Salvaged"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function Handles actions for the Prior Purchased Column
' Returns : "Problem" in suggestion fields based on given criteria
' Created : 06/18/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function PriorPurchase() As Boolean
'If Salvaged column is anything but yes (No or Unknown) and Feature is Pipe, Bend (or Tee, Reducer, Mfg Bend & Date < 3/1/1940)
'Then macro should return Problem
    If (sPriorOperator = "Yes" Or sPriorOperator = "yes") Then
        Problem2_PriorPurchase = 1
        DoNotAutoPopSeam = 1
        DoNotAutoPopWT = 1
        DoNotAutoPopSMYS = 1
        DoNotAutoPopWT2 = 1
    End If
    PriorPurchase = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    PriorPurchase = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in PriorPurchase"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function Handles actions for tall distinct problems a user may run into
' Returns : List of what type of Problem is occuring with description
' Created : 06/13/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function Problems() As Boolean
'Default iProblemAutoPop to 0
iProblemAutoPop = 0
sProblemType = "Problems"
    If Problem1_Salvaged = 1 Then
    Problem1_Salvaged = 1
     'Insert FVE Comments Value in Blue
    If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Feature is marked as Reconditioned / Salvaged.*") Then
        'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|1| Feature is marked as Reconditioned / Salvaged. ")
        Range("EV" & CStr(lLastRowdDiameter)).Select
            With ActiveCell
                .Characters(Len(.Value) + 1).Insert UCase("|1| Feature is marked as Reconditioned / Salvaged. ")
            End With
        btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|1| Feature is marked as Reconditioned / Salvaged. ")
        'Selection.Font.Color = vbBlue
        'Selection.Font.Bold = True
    End If
    iProblemAutoPop = 1
    sProblemType = sProblemType & " |1|"
    End If
    If Problem2_PriorPurchase = 1 Then
         If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Feature is marked as a Prior Purchase.*") Then
            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|2| Feature is marked as a Prior Purchase.  ")
            Range("EV" & CStr(lLastRowdDiameter)).Select
            With ActiveCell
                .Characters(Len(.Value) + 1).Insert UCase("|2| Feature is marked as a Prior Purchase.  ")
            End With
            btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|2| Feature is marked as a Prior Purchase.  ")
            'Selection.Font.Color = vbBlue
            'Selection.Font.Bold = True
        End If
        iProblemAutoPop = 2
        sProblemType = sProblemType & " |2|"
    End If
    
    If Problem1_Salvaged <> 1 And Problem2_PriorPurchase <> 1 Then
        If Problem3_SMYS = 1 Then
             If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*SMYS not possible under given circumstances, refer to PRUPF.*") Then
                'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|3| SMYS not possible under given circumstances, refer to PRUPF.  ")
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert UCase("|3| SMYS not possible under given circumstances, refer to PRUPF.  ")
                End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|3| SMYS not possible under given circumstances, refer to PRUPF.  ")
                'Selection.Font.Color = vbBlue
                'Selection.Font.Bold = True
            End If
            iProblemAutoPop = 3
            sProblemType = sProblemType & " |3|"
        End If
        If Problem4_WT = 1 Then
             If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*W.T. suggestion not available, or not possible under given circumstances, refer to PRUPF.*") Then
                'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value) ' + "|4| W.T. not possible under given circumstances, refer to PRUPF.  ")
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert UCase("|4| W.T. suggestion not available, or not possible under given circumstances, refer to PRUPF.  ")
                End With
                Worksheets("Pipe Data").Range("EV" & CStr(lLastRowdDiameter)).Select
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "W.T. suggestion not available, or not possible under given circumstances, refer to PRUPF.  ")
                'Selection.Font.Color = vbBlue
                'Selection.Font.Bold = True
            End If
            iProblemAutoPop = 4
            sProblemType = sProblemType & " |4|"
        End If
        If Problem5_Seam = 1 Then
             If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Seam not possible under given circumstances, refer to PRUPF.*") Then
                'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|5| Seam not possible under given circumstances, refer to PRUPF.  ")
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert UCase("|5| Seam not possible under given circumstances, refer to PRUPF.  ")
                End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|5| Seam not possible under given circumstances, refer to PRUPF.  ")
                'Selection.Font.Color = vbBlue
                'Selection.Font.Bold = True
            End If
            iProblemAutoPop = 5
            sProblemType = sProblemType & " |5|"
        End If
        If Problem6_Exceptions = 1 Then
            If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Exception for Diameter*" Then
                'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|6| Exception for Diameter " & dDiameter & ", See PRUPF Table 2 & 3 notes.  ")
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert "|6| Exception for Diameter " & dDiameter & ", See PRUPF Table 2 & 3 notes.  "
                End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "|6| Exception for Diameter " & dDiameter & ", See PRUPF Table 2 & 3 notes.  ")
                'Selection.Font.Color = vbBlue
                'Selection.Font.Bold = True
            End If
            iProblemAutoPop = 6
            sProblemType = sProblemType & " |6|"
        End If
        If Problem7_SourceWT = 1 Then
             If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Could not find valid Source W.T.*") Then
                'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|7| Could not find valid Source W.T.  ")
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert UCase("|7| Could not find valid Source W.T.  ")
                End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|7| Could not find valid Source W.T.  ")
                'Selection.Font.Color = vbBlue
                'Selection.Font.Bold = True
            End If
            iProblemAutoPop = 7
            sProblemType = sProblemType & " |7|"
        End If
        If Problem8_NewSeam = 1 Then
             If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Seam type not yet in PRUPF*") Then
                'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "|8| Seam type not yet in PRUPF.  ")
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert UCase("|8| Seam type not yet in PRUPF.  ")
                End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "|8| Seam type not yet in PRUPF.  ")
                'Selection.Font.Color = vbBlue
                'Selection.Font.Bold = True
            End If
            iProblemAutoPop = 8
            sProblemType = sProblemType & " |8|"
        End If
        'Check if Mfg Bend has an Angle < 30 degrees (This is information only, no action!)
        If sFeature = "Mfg Bend" And Worksheets("Pipe Data").Range(AngleColumn & Trim(Str(lLastRowdDiameter))).Value < 30 Then
            If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Manufactured bend angle is small enough such that feature may be field bend instead - Ref PFL Build Guideline AKM-MAOP-404G.*" Then
                Range("EV" & CStr(lLastRowdDiameter)).Select
                With ActiveCell
                    .Characters(Len(.Value) + 1).Insert "Manufactured bend angle is small enough such that feature may be field bend instead - Ref PFL Build Guideline AKM-MAOP-404G.  "
                End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "Manufactured bend angle is small enough such that feature may be field bend instead - Ref PFL Build Guideline AKM-MAOP-404G.  ")
            End If
        End If
    End If
    Problems = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Problems = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Problems"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function calculates Source W.T. 1 and Source W.T. 2 to prep for comparison
' Returns : Fittings - Source W.T. 1 & 2
' Created : 02/27/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function Fittings_Source_WT_1_2(p_lRowIndex As Long) As Boolean
    Dim iCurrentRow As Long
    iCurrentRow = p_lRowIndex
    iRowsAway1 = 0
    iRowsAway2 = 0
    vSourceWT = 0
    vSourceWT1 = 0
    vSourceWT2 = 0
    'Choose Larger OD for Reducers and Tees by defaulting dDiameter to larger OD
    If sFeature = "Reducer" Or sFeature = "Tee" Then
        If dDiameter2 > dDiameter Then
            dDiameter = dDiameter2
            UseOD2 = 1
        End If
    End If
    'SourceWT1 with matching Source Diameter1 for Reducers with matching Diameters (Mfg Bend, Tee, Reducer)
    If sFeature = "Reducer" Or sFeature = "Tee" Or sFeature = "Mfg Bend" Or sFeature = "Cap" Then
        Do While vSourceWT1 = 0
            If p_lRowIndex = 3 Then ' (E)
                vSourceWT1 = SETTING_NOT_APPLICABLE
                sSourceFeature1 = "Last Row"
            Else
                If (Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow - 1))).Value = "Pipe" _
                Or Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow - 1))).Value = "Field Bend") _
                   And Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow - 1))).Value <> "N/A" _
                   And Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow - 1))).Value <> "" _
                   And dDiameter = Worksheets("Pipe Data").Range(sDiamterColumn & Trim(Str(iCurrentRow - 1))).Value Then
                    vSourceWT1 = Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow - 1))).Value
                    sSourceFeature1 = Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow - 1))).Value
                    iCurrentRow = iCurrentRow + 1
                ElseIf (iCurrentRow = 2) Then
                    vSourceWT1 = SETTING_NOT_APPLICABLE
                Else
                    iCurrentRow = iCurrentRow - 1
                    iRowsAway1 = iRowsAway1 + 1
                End If
            End If '(E)
        Loop
    End If
    'Establish Source2 WT1 for Mfg Bends and Tees
    'Reset Current Row
    iCurrentRow = p_lRowIndex
    'Establish 2st Source WT by going down in the spreadhseet (Mfg Bend, Tee)
    If sFeature = "Mfg Bend" Or sFeature = "Tee" Or sFeature = "Reducer" Or sFeature = "Cap" Then
        Do While vSourceWT2 = 0
            If p_lRowIndex = iPipeDataLastRow Then '(F)
                vSourceWT2 = SETTING_NOT_APPLICABLE
                sSourceFeature3 = "Last Row"
            Else
                If (Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow + 1))).Value = "Pipe" _
                Or Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow + 1))).Value = "Field Bend") _
                And Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow + 1))).Value <> "N/A" _
                And Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow + 1))).Value <> "" _
                And dDiameter = Worksheets("Pipe Data").Range(sDiamterColumn & Trim(Str(iCurrentRow + 1))).Value Then
                    vSourceWT2 = Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow + 1))).Value
                    sSourceFeature2 = Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow + 1))).Value
                    iCurrentRow = iCurrentRow + 1
                ElseIf iCurrentRow = iPipeDataLastRow Then
                    vSourceWT2 = SETTING_NOT_APPLICABLE
                Else
                    iCurrentRow = iCurrentRow + 1
                    iRowsAway2 = iRowsAway2 + 1
                End If
            End If '(F)
        Loop
    End If
    'Identify OD2 small OD
    If UseOD2 = 0 Then
        dDiameterSmall = dDiameter2
    Else
        dDiameterSmall = Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value 'PULL BACK OD1 FROM PFL HERE
    End If
    'Source WT for Reducer for smaller OD**************
            'For Tee OD, Just put nothing and default to Standard Wall as in notes:
    iRowsAway3 = 0
    vSource2WT = 0
    'Reset Current Row
    iCurrentRow = p_lRowIndex
    'Establish OD2 Source WT by going down in the spreadhseet (Reducer)
    If sFeature = "Reducer" Then
        Do While vSource2WT = 0
            If p_lRowIndex = iPipeDataLastRow Then '(F)
                vSource2WT = SETTING_NOT_APPLICABLE
                sSourceFeature2 = "Last Row"
            Else
                If (Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow + 1))).Value = "Pipe" _
                Or Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow + 1))).Value = "Field Bend") _
                And Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow + 1))).Value <> "N/A" _
                And Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow + 1))).Value <> "" _
                And dDiameterSmall = Worksheets("Pipe Data").Range(sDiamterColumn2 & Trim(Str(iCurrentRow + 1))).Value Then
                    vSource2WT = Worksheets("Pipe Data").Range(vWTUserColumn & Trim(Str(iCurrentRow + 1))).Value
                    sSourceFeature2 = Worksheets("Pipe Data").Range(FeatureColumn & Trim(Str(iCurrentRow + 1))).Value
                    iCurrentRow = iCurrentRow + 1
                ElseIf iCurrentRow = iPipeDataLastRow Then
                    vSource2WT = SETTING_NOT_APPLICABLE
                Else
                    iCurrentRow = iCurrentRow + 1
                    iRowsAway3 = iRowsAway3 + 1
                End If
            End If '(F)
        Loop
    End If
    'Determind if Source WT is valid or not
    If ((vSourceWT1 = Empty Or vSourceWT1 = "N/A") And (iRowsAway2 = 0 And vSourceWT2 <> Empty And vSourceWT2 <> "N/A") _
    Or (vSourceWT2 = Empty Or vSourceWT2 = "N/A") And (iRowsAway1 = 0 And vSourceWT1 <> Empty And vSourceWT1 <> "N/A")) _
    Or ((vSourceWT1 <> Empty And vSourceWT1 <> "N/A") And (vSourceWT2 <> Empty And vSourceWT2 <> "N/A")) _
    And (vSourceWT1 <> SETTING_NOT_APPLICABLE Or vSourceWT2 <> SETTING_NOT_APPLICABLE) Then
        sValidSourceWT = 1
    Else
        sValidSourceWT = 0
    End If
    'Determind if Source2 WT is valid or not
    If vSource2WT = Empty Or vSource2WT = "N/A" Or vSource2WT = 0 Or vSource2WT = SETTING_NOT_APPLICABLE Then
        sValidSource2WT = 0
    Else
        sValidSource2WT = 1
    End If
    Fittings_Source_WT_1_2 = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Fittings_Source_WT_1_2 = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in FITTINGS-Source W.T. 1&2"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function determines the source W.T. from comparing Source W.T. 1 and Source W.T. 2
' Returns : Fittings - Source W.T.
' Created : 02/27/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function Fittings_SourceWT() As Boolean
'Determine Source WT by comparing and choosing best of two Source WT
        If vSourceWT1 = SETTING_NOT_APPLICABLE And vSourceWT2 <> SETTING_NOT_APPLICABLE Then  '(H)
            vSourceWT = vSourceWT2
        ElseIf vSourceWT2 = SETTING_NOT_APPLICABLE And vSourceWT1 <> SETTING_NOT_APPLICABLE Then
            vSourceWT = vSourceWT1
        ElseIf vSourceWT2 = SETTING_NOT_APPLICABLE And vSourceWT1 <> SETTING_NOT_APPLICABLE Then
            vSourceWT = SETTING_NOT_APPLICABLE
        ElseIf vSourceWT1 <> SETTING_NOT_APPLICABLE And vSourceWT2 <> SETTING_NOT_APPLICABLE Then
            ' (I)
            If (sSourceFeature1 = "Pipe" Or sSourceFeature1 = "Field Bend") _
              And (sSourceFeature2 <> "Pipe" And sSourceFeature2 <> "Field Bend") Then
                vSourceWT = vSourceWT1
            ElseIf (sSourceFeature2 = "Pipe" Or sSourceFeature2 = "Field Bend") _
              And (sSourceFeature1 <> "Pipe" And sSourceFeature1 <> "Field Bend") Then
                vSourceWT = vSourceWT2
            ElseIf (sSourceFeature1 = "Pipe" Or sSourceFeature1 = "Field Bend") _
              And (sSourceFeature2 = "Pipe" Or sSourceFeature2 = "Field Bend") Then
                'Choose Source WT when both WT are Priorities:
                'source-fitting-source = select lowest of the two source WT to serve as the source.
                'source-other-fitting-other-other-other-source = select lowest of the two source WTs.
                'source-fitting-other-source = select the contiguous source.
                If (iRowsAway1 = iRowsAway2) Or (iRowsAway1 > 0 And iRowsAway2 > 0) Then
                    'Choose to take Maximum or Minimum Source Wall Thickness
                    If vSourceWT1 < vSourceWT2 Then
                        vSourceWT = vSourceWT2
                    ElseIf vSourceWT1 > vSourceWT2 Then
                        vSourceWT = vSourceWT1
                    ElseIf vSourceWT1 = vSourceWT2 Then
                        vSourceWT = vSourceWT1
                    End If
                ElseIf iRowsAway1 = 0 And iRowsAway2 > 0 Then
                    vSourceWT = vSourceWT1
                ElseIf iRowsAway1 > 0 And iRowsAway2 = 0 Then
                    vSourceWT = vSourceWT2
                End If
            End If '(I)
        End If '(H)
        Fittings_SourceWT = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Fittings_SourceWT = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Fittings_SourceWT"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function returns a suggested SMYS & Wall Thickness for Fittings
'           for Time Frames 1-3
' Returns : Fitting - SMYS & W.T. Time Frames 1-3
' Created : 08/21/2011 example
' --------------------------------------------------------------------------------------------------------------
Private Function Fittings_TimeFrame1_3() As Boolean
'Proccess 22):
    '1.) (Time Frame 1: Unknown Install Date)
    If (Range("DZ" & CStr(lLastRowdDiameter)).Value = Empty _
    Or Range("DZ" & CStr(lLastRowdDiameter)).Value Like "" _
    Or Range("DZ" & CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    Or Range("DZ" & CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And (Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value = Empty _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "" _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*") Then
        vWT = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
        vWT2 = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
        vSMYS = 24000
   ' End If
    '2.) (Time Frame 2: < 1/1/30)
    ElseIf ((Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value = Empty _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "" _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And DateValue(vInstallYr) < DateValue("1/1/1940")) _
    Or DateValue(vPurchaseYr) < DateValue("1/1/1930") Then
    '(DateValue(vPurchaseYr) < DateValue("1/1/1930")) Or (DateValue(vInstallYr) < DateValue("1/1/1940")) Then
        'Just choose Standard Wall Thickness!
        vWT = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
        vWT2 = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
        vSMYS = 24000
    'End If
    '3.) (Time Frame 3: 1/1/30 - 2/28/40)
    ElseIf ((Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value = Empty _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "" _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
    Or Range(vPurchaseDateColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And DateValue(vInstallYr) >= DateValue("1/1/1940") And DateValue(vInstallYr) <= DateValue("2/28/1940")) _
    Or DateValue(vPurchaseYr) >= DateValue("1/1/1930") And DateValue(vPurchaseYr) <= DateValue("2/28/1940") Then
    'ElseIf ((DateValue(vPurchaseYr) >= DateValue("1/1/1930") Or DateValue(vInstallYr) >= DateValue("1/1/1940")) _
    'And (DateValue(vPurchaseYr) <= DateValue("2/28/1940") Or DateValue(vInstallYr) <= DateValue("2/28/1940"))) _
    'And (Range("DZ" & CStr(lLastRowdDiameter)).Value <> Empty _
    'And Range("DZ" & CStr(lLastRowdDiameter)).Value <> "Unknown" _
    'And Range("DZ" & CStr(lLastRowdDiameter)).Value <> "unknown") Then
        'Just choose Standard Wall Thickness!
        vWT = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
        vWT2 = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
        vSMYS = 30000
    End If
    Fittings_TimeFrame1_3 = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Fittings_TimeFrame1_3 = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Fittings_TimeFrame1-3"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function returns a suggested SMYS & Wall Thickness for Fittings
'           for Time Frame 4: 3/1/40 - 4/11/63
' Returns : Fitting - SMYS & W.T. Time Frame 4
' Created : 08/21/2011 example
' --------------------------------------------------------------------------------------------------------------
Private Function Fittings_TimeFrame4() As Boolean
    'New vMAOP Calc (0 for first iteration):
    vnewvMAOP = 0
    '4.)Time Frame 4: 3/1/40 - 4/11/63
    'Starting Step 1
    If (DateValue(vInstallYr) >= DateValue("3/1/1940")) _
    And (DateValue(vInstallYr) <= DateValue("4/11/1963")) Then
    'Or (sFeature = "Tee" And DateValue(vInstallYr) >= DateValue("3/1/1940")) Then
    'OD LARGE
    'Step 1
        If (vSourceWT <= fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 30000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 2
        If ((vSourceWT > fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
        And vSourceWT <= fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
        And vSequenceIterator <> "Complete")) _
        Or vSequenceIterator = "InComplete" Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 30000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
     'Step 3
        If ((vSourceWT > fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) _
          And (vSequenceIterator <> "Complete")) _
          Or (vSequenceIterator = "InComplete") Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 30000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 4
        If vSequenceIterator = "InComplete" Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 35000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 5
        If vSequenceIterator = "InComplete" Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 35000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 6
        If vSequenceIterator = "InComplete" Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 35000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
         'Output comment if MAOP is never met
        If vSequenceIterator = "InComplete" And (UnknownSMYS = 1 Or UnknownWT = 1) Then
            If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" Then
                 'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                 Range("EV" & CStr(lLastRowdDiameter)).Select
                 With ActiveCell
                    .Characters(Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. "
                 End With
                 btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
             End If
        End If '
        
        'OD SMALL
        If dDiameterSmall <> 0 And (sFeature = "Reducer" Or sFeature = "Tee") Then 'And UnknownWT2 = 1 Then
             'Choose Smaller Diameter
             'If UseOD2 = 1 Then
             '    If Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value = "PipeOD+(2*SleeveWT)+0.25" Then
             '        dDiameter = 0
             '    Else
             '        dDiameter = Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value
             '    End If
             'Else
             '    dDiameter = dDiameterSmall
             'End If
            'Set SourceWT2 if unknown to Unkown Wt from table 7
            'If sValidSource2WT = 0 Then
            '    vSource2WT = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            'End If
               'Default vSequenceIterator to InComplete
               'vSequenceIterator = "InComplete"
               'Step 1
                   'If (vSource2WT <= fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
                   'And sFeature <> "Tee") Then
                       'Use Known WT if available
                   vWT2 = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Use Known SMYS from larger OD
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               'End If
           'Step 2
               'If ((vSource2WT > fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
               'And vSource2WT <= fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _

               If vSequenceIterator <> "Complete" Then
                   'Use either Standard Choice WT
                       vWT2 = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               End If
            'Step 3
               If vSequenceIterator <> "Complete" Then
                   'Use either Standard Choice WT
                       vWT2 = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               End If
                'Step 3
               If vSequenceIterator <> "Complete" Then
                   'Use either Standard Choice WT
                    vWT2 = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               End If
         'Stop here!
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               'End If
           'End If
           'Output comment if MAOP is never met
                If vSequenceIterator = "InComplete" And (UnknownSMYS = 1 Or UnknownWT = 1 Or UnknownWT2 = 1) Then
                    If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" Then
                        'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                         Range("EV" & CStr(lLastRowdDiameter)).Select
                         With ActiveCell
                            .Characters(Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. "
                         End With
                        btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
                    End If
                End If '
        End If
    End If
    Fittings_TimeFrame4 = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Fittings_TimeFrame4 = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Fittings_TimeFrame4"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function returns a suggested SMYS & Wall Thickness for Fittings
'           for Time Frame 5: 4/12/1963 - 10/31/1968
' Returns : Fitting - SMYS & W.T. Time Frame 5
' Created : 08/21/2011 example
' --------------------------------------------------------------------------------------------------------------
Private Function Fittings_TimeFrame5() As Boolean
    'Time Frame 5: 4/12/1963 - 10/31/1968
    vnewvMAOP = 0
    If (DateValue(vInstallYr) >= DateValue("4/12/1963")) Then
    'And sFeature <> "Tee") Then    '(J)
            'Step 1
        If (vSourceWT <= fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 35000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 2
        If ((vSourceWT > fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
        And vSourceWT <= fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") And vSequenceIterator <> "Complete") _
        Or vSequenceIterator = "InComplete") Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 35000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
     'Step 3
        If ((vSourceWT > fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")) _
          And (vSequenceIterator <> "Complete")) _
          Or (vSequenceIterator = "InComplete") Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 35000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 4
        If vSequenceIterator = "InComplete" Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                'Choose Source WT if larger than Extra Heavy WT
                If vSourceWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") Then
                    vWT = vSourceWT
                Else
                    vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
                End If
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 35000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 5
        If vSequenceIterator = "InComplete" Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                'Choose Source WT if larger than Extra Heavy WT
                If vSourceWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") Then
                    vWT = vSourceWT
                Else
                    vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
                End If
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 42000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
    'Step 6
        If vSequenceIterator = "InComplete" Then
            'Use Known WT if available
            If UnknownWT = 0 Then
                vWT = vUserWT
            Else
                'Choose Source WT if larger than Extra Heavy WT
                If vSourceWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") Then
                    vWT = vSourceWT
                Else
                    vWT = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
                End If
            End If
            'Use Known SMYS if available
            If UnknownSMYS = 0 Then
                vSMYS = vUserSMYS
            Else
                vSMYS = 52000
            End If
            'Define MAOP
            vnewvMAOP = CInt((2 * vSMYS * vWT * vDesignFactor * vLsFactor) / (dDiameter))
            'Test MAOP
            If vnewvMAOP >= vMAOP Then  '(M)
                vSequenceIterator = "Complete"
            Else
                vSequenceIterator = "InComplete"
            End If
        End If
        'Output comment if MAOP is never met
        If vSequenceIterator = "InComplete" And (UnknownSMYS = 1 Or UnknownWT = 1) Then
            If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" Then
                'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                  Range("EV" & CStr(lLastRowdDiameter)).Select
                  With ActiveCell
                     .Characters(Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. "
                  End With
                  btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
            End If
        End If '
        'Reset vSequenceIterator
        vSequenceIterator = "Start"
            'OD SMALL
         If dDiameter2 <> 0 And (sFeature = "Reducer" Or sFeature = "Tee") Then ' and UnknownWT2 = 1 Then
             'Choose Smaller Diameter
             'If UseOD2 = 1 Then
             '    If Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value = "PipeOD+(2*SleeveWT)+0.25" Then
             '        dDiameter = 0
             '    Else
             '        dDiameter = Range(sDiamterColumn & CStr(lLastRowdDiameter)).Value
             '    End If
             'Else
             '    dDiameter = dDiameterSmall
             'End If
        'Set SourceWT2 if unknown to Unkown Wt from table 7
        'If sValidSource2WT = 0 Then
        '    vSource2WT = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B")
        'End If
           'Default vSequenceIterator to InComplete
           'vSequenceIterator = "InComplete"
           'Step 1
               'If (vSource2WT <= fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
               'And sFeature <> "Tee") Then
                   'Use Known WT if available
                   vWT2 = fnxDetermineThickness1("Logic", iUnknownTbl, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Use Known SMYS from larger OD
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               'End If
           'Step 2
               If vSequenceIterator <> "Complete" Then
                   'Use either Standard Choice WT
                       vWT2 = fnxDetermineThickness1("Logic", iFirstChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               End If
            'Step 3
               If vSequenceIterator <> "Complete" Then
                   'Use either Standard Choice WT
                       vWT2 = fnxDetermineThickness1("Logic", iSecond_StandardChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               End If
                'Step 3
               If vSequenceIterator <> "Complete" Then
                   'Use either Standard Choice WT
                    vWT2 = fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameterSmall, sSeam, vPurchaseYr, "A", "B")
                   'Define MAOP
                   vnewvMAOP = CInt((2 * vSMYS * vWT2 * vDesignFactor * vLsFactor) / (dDiameterSmall))
                   'Test MAOP
                   If vnewvMAOP >= vMAOP Then  '(M)
                       vSequenceIterator = "Complete"
                   Else
                       vSequenceIterator = "InComplete"
                   End If
               End If
        'Stop here!
          'End If
          'Output comment if MAOP is never met
          If vSequenceIterator = "InComplete" And (UnknownSMYS = 1 Or UnknownWT = 1 Or UnknownWT2 = 1) Then
              If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*Suggestion/s are based on maximum values and MAOP did not pass.*" Then
                  'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "Suggestion/s are based on maximum values and MAOP did not pass. ")
                 Range("EV" & CStr(lLastRowdDiameter)).Select
                 With ActiveCell
                    .Characters(Len(.Value) + 1).Insert "Suggestion/s are based on maximum values and MAOP did not pass. "
                 End With
                btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "Suggestion/s are based on maximum values and MAOP did not pass. ")
              End If
          End If
        End If
    End If 'Date
    Fittings_TimeFrame5 = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    Fittings_TimeFrame5 = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Fittings_TimeFrame5"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function will populate a suggestion in the appropriate field if there is no value there
'           currently.  It will also put a 1 in the rational, fill out the FVE comments and make it blue.
' Returns : Suggestion, Rat'le and Comments
' Created : 6/18/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function SourceWT_NA() As Boolean
   'Error Handling (Instances where source W.T. is Blank or N/A) (G)
                    'Set SMYS when W.T is Invalid
                    ' Time Frame 1
                    If (DateValue(vInstallYr) >= DateValue("4/12/1963")) Then
                        vSMYS = 35000
                    ' Time Frame 2
                    ElseIf (DateValue(vInstallYr) >= DateValue("3/1/1940")) _
                      And (DateValue(vInstallYr) <= DateValue("4/11/1963")) Then
                        vSMYS = 30000
                    ' Time Frame 3
                    ElseIf (DateValue(vInstallYr) >= DateValue("1/1/1930")) _
                      And (DateValue(vInstallYr) <= DateValue("2/28/1940")) Then
                        vSMYS = 30000
                    ' Time Frame 4
                    ElseIf (DateValue(vInstallYr) < DateValue("1/1/1930")) Then
                        vSMYS = 24000
                    End If
                    ' Unknown Install Date
                    If Range("F" & CStr(lLastRowdDiameter)).Value = Empty _
                             Or Range("F" & CStr(lLastRowdDiameter)).Value = "" _
                             Or Range("F" & CStr(lLastRowdDiameter)).Value Like "*Unknown*" _
                             Or Range("F" & CStr(lLastRowdDiameter)).Value Like "*unknown*" Then
                        vSMYS = 24000
                    End If
                    'Set Wall Thickness to invalid
                    vWT = "Invld Src W.T."
                    Problem7_SourceWT = 1
    SourceWT_NA = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    SourceWT_NA = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in SourceWT_NA"
    Resume PROC_EXIT
End Function
' --------------------------------------------------------------------------------------------------------------
' Comments: This function will populate a suggestion in the appropriate field if there is no value there
'           currently.  It will also put a 1 in the rational, fill out the FVE comments and make it blue.
' Returns : Suggestion, Rat'le and Comments
' Created : 5/23/2012 example
' --------------------------------------------------------------------------------------------------------------
Private Function AutoPopulate() As Boolean
    ' Handle Auto-populate for instances where Seam, SMYMS or W.T. are blank
    ' If Seam is Blank (Auto Pop)
    If ((Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildSeamTypeColumn & CStr(lLastRowdDiameter)).Value = "N/A") _
    And (Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value = 0 Or Range("DK" & CStr(lLastRowdDiameter)).Value = Empty Or Range("DK" & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range("DK" & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range("DK" & CStr(lLastRowdDiameter)).Value = "N/A")) _
    And Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value <> 0 _
    And Not Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value Like "*error*" _
    And Not Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And Not Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value <> Empty _
    And DoNotAutoPopSeam = 0 Then
    'And Problem1_Salvaged <> 1 _
    'And Problem2_PriorPurchase <> 1 _
    'And Problem5_Seam <> 1 _
    'And Problem6_Exceptions <> 1
        Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Value = Range(vSuggestedSeamColumn & CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(sUserSeamColumn & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("DN" & CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DN" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
         'Update Catagory and Tier/Category Columns
        Range("EX" & CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("EY" & CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Pipe30Inch = 1
        ' Make inserted Suggestion, Rational & FVE Comments Value Blue
        If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*SEAM was calculated by PRUPF Logic*" Then
            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "SEAM was calculated by PRUPF Logic. ")
            Range("EV" & CStr(lLastRowdDiameter)).Select
            With ActiveCell
               .Characters(Len(.Value) + 1).Insert "SEAM was calculated by PRUPF Logic. "
            End With
            btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "SEAM was calculated by PRUPF Logic. ")
            'Selection.Font.Color = vbBlue
            'Selection.Font.Bold = True
        End If
    End If
    ' If SMYS is Blank (Auto Pop)
    If ((Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildSMYSColumn & CStr(lLastRowdDiameter)).Value = "N/A") _
    And (Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value = 0 Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value = "N/A")) _
    And Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value <> "Field Bend" _
    And Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value <> 0 _
    And Not Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value Like "*error*" _
    And Not Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And Not Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value <> Empty _
    And DoNotAutoPopSMYS = 0 Then
    'And Problem1_Salvaged <> 1 _
    'And Problem2_PriorPurchase <> 1 _
    'And Problem3_SMYS <> 1 _
    'And Problem6_Exceptions <> 1 Then
        Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Value = Range(vSuggestedSMYSColumn & CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(vSMYSUserColumn & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("DB" & CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DB" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        'Update Catagory and Tier/Category Columns
        Range("EX" & CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("EY" & CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        SMYS35K = 1
        Pipe30Inch = 1
        ' Make inserted Suggestion, Rational & FVE Comments Value Blue
        If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*SMYS was calculated by PRUPF Logic*" Then
            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "SMYS was calculated by PRUPF Logic. ")
            Range("EV" & CStr(lLastRowdDiameter)).Select
            With ActiveCell
               .Characters(Len(.Value) + 1).Insert "SMYS was calculated by PRUPF Logic. "
            End With
            btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "SMYS was calculated by PRUPF Logic. ")
            'Selection.Font.Color = vbBlue
            'Selection.Font.Bold = True
        End If
    End If
    ' If W.T. is Blank (Auto Pop)
    If ((Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildWTColumn & CStr(lLastRowdDiameter)).Value = "N/A") _
    And (Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value = Empty) Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value = "N/A" Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value <> 0 _
    And Not Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value Like "*error*" _
    And Not Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And Not Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value <> Empty _
    And DoNotAutoPopWT = 0 Then
    'And Problem1_Salvaged <> 1 _
    'And Problem2_PriorPurchase <> 1 _
    'And Problem4_WT <> 1 _
    'And Problem6_Exceptions <> 1 Then
        Range(vWTUserColumn & CStr(lLastRowdDiameter)).Value = Range(vSuggestedWTColumn & CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(vWTUserColumn & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("DG" & CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DG" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        'Update Catagory and Tier/Category Columns
        Range("EX" & CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("EY" & CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Pipe30Inch = 1
        ' Make inserted Suggestion, Rational & FVE Comments Value Blue
        If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*W.T. 1 was calculated by PRUPF Logic*" Then
            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "W.T. 1 was calculated by PRUPF Logic. ")
            Range("EV" & CStr(lLastRowdDiameter)).Select
            With ActiveCell
               .Characters(Len(.Value) + 1).Insert "W.T. 1 was calculated by PRUPF Logic. "
            End With
            btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "W.T. 1 was calculated by PRUPF Logic. ")
            'Selection.Font.Color = vbBlue
            'Selection.Font.Bold = True
        End If
    End If
       ' If W.T.2 is Blank (Auto Pop)
    If ((Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value = Empty Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value = "" Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value Like "*unknown*" Or Range(vBuildWT2Column & CStr(lLastRowdDiameter)).Value = "N/A") _
    And (Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value = "" Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value = Empty) Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value = "N/A" Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value Like "*Unknown*" Or Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value Like "*unknown*") _
    And Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value <> 0 _
    And Not Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value Like "*error*" _
    And Not Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value Like "*problem*" _
    And Not Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value Like "*Problem*" _
    And Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value <> Empty _
    And DoNotAutoPopWT2 = 0 Then
    'And Problem1_Salvaged <> 1 _
    'And Problem2_PriorPurchase <> 1 _
    'And Problem4_WT <> 1 _
    'And Problem6_Exceptions <> 1 Then
        Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Value = Range(vSuggestedWT2Column & CStr(lLastRowdDiameter)).Value
        Worksheets("Pipe Data").Range(vWT2UserColumn & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("DG" & CStr(lLastRowdDiameter)).Value = 1
        Worksheets("Pipe Data").Range("DG" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        'Update Catagory and Tier/Category Columns
        Range("EX" & CStr(lLastRowdDiameter)).Value = "HRD"
        Worksheets("Pipe Data").Range("EX" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Range("EY" & CStr(lLastRowdDiameter)).Value = "1"
        Worksheets("Pipe Data").Range("EY" & CStr(lLastRowdDiameter)).Select
        Selection.Font.Color = vbBlue
        Selection.Font.Bold = True
        Pipe30Inch = 1
        ' Make inserted Suggestion, Rational & FVE Comments Value Blue
        If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*W.T. 2 was calculated by PRUPF Logic*" Then
            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "W.T. 2 was calculated by PRUPF Logic. ")
            Range("EV" & CStr(lLastRowdDiameter)).Select
            With ActiveCell
               .Characters(Len(.Value) + 1).Insert "W.T. 2 was calculated by PRUPF Logic. "
            End With
            btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "W.T. 2 was calculated by PRUPF Logic. ")
            'Selection.Font.Color = vbBlue
            'Selection.Font.Bold = True
        End If
    End If
    ' If MAOP doesn't pass, fill in SME comments
    If iSMEComments = 1 And Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*W.T. Suggestion is based on maximum values and MAOP did not pass.*" Then
        'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "W.T. Suggestion is based on maximum values and MAOP did not pass. ")
        Range("EV" & CStr(lLastRowdDiameter)).Select
        With ActiveCell
           .Characters(Len(.Value) + 1).Insert "W.T. Suggestion is based on maximum values and MAOP did not pass. "
        End With
        btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "W.T. Suggestion is based on maximum values and MAOP did not pass. ")
        'Selection.Font.Color = vbBlue
        'Selection.Font.Bold = True
    End If
    ' If W.T. Exceeds E.H. W.T., then fill in WT > EH Comments
    If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like "*This W.T. exceeds Extra Heavy W.T. for this diameter.*" _
    And IsNumeric(vWT) = True _
    And vWT > fnxDetermineThickness1("Logic", iThirdChoice, 3, dDiameter, sSeam, vPurchaseYr, "A", "B") _
    And (Worksheets("Pipe Data").Range(FeatureColumn & Trim(CStr(lLastRowdDiameter))).Value = "Tee" _
          Or (Worksheets("Pipe Data").Range(FeatureColumn & Trim(CStr(lLastRowdDiameter))).Value = "Mfg Bend") _
          Or (Worksheets("Pipe Data").Range(FeatureColumn & Trim(CStr(lLastRowdDiameter))).Value = "Reducer") _
          Or (Worksheets("Pipe Data").Range(FeatureColumn & Trim(CStr(lLastRowdDiameter))).Value = "Cap")) Then
        'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "This W.T. exceeds Extra Heavy W.T. for this diameter. ")
        Range("EV" & CStr(lLastRowdDiameter)).Select
        With ActiveCell
           .Characters(Len(.Value) + 1).Insert "This W.T. exceeds Extra Heavy W.T. for this diameter. "
        End With
        btest = ColorText("EV" & CStr(lLastRowdDiameter), "DarkBlue", "This W.T. exceeds Extra Heavy W.T. for this diameter. ")
        'Selection.Font.Color = vbBlue
        'Selection.Font.Bold = True
    End If
    'If SMYS > 35K, "Requires Technical Pier Review", Only for fittings
    If fnxIsItIn(sFeature, "Tee", "Mfg Bend", "Reducer", "Cap") _
    And vSMYS > 35000 _
    And SMYS35K = 1 Then
        If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*SMYS > 35, Requires Technical Peer Review*") Then
            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "SMYS > 35, Requires Technical Peer Review. ")
        Range("EV" & CStr(lLastRowdDiameter)).Select
        With ActiveCell
           .Characters(Len(.Value) + 1).Insert UCase("SMYS > 35, Requires Technical Peer Review. ")
        End With
        btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "SMYS > 35, Requires Technical Peer Review. ")
            'Selection.Font.Color = vbBlue
            'Selection.Font.Bold = True
        End If
    End If
'Process 23:
    'If Macro auto populates and the Diameter = 30.5, Requires Tech Peer review
    If Pipe30Inch = 1 And dDiameter = 30.5 Then
        If Not Range("EV" & CStr(lLastRowdDiameter)).Value Like UCase("*Diameter = 30.5, Requires Technical Peer Review*") Then
            'Range("EV" & CStr(lLastRowdDiameter)).Value = (Range("EV" & CStr(lLastRowdDiameter)).Value + "Diameter = 30.5, Requires Technical Peer Review. ")
            Range("EV" & CStr(lLastRowdDiameter)).Select
            With ActiveCell
               .Characters(Len(.Value) + 1).Insert UCase("Diameter = 30.5, Requires Technical Peer Review. ")
            End With
            btest = ColorText("EV" & CStr(lLastRowdDiameter), "purple", "Diameter = 30.5, Requires Technical Peer Review. ")
            'Selection.Font.Color = vbBlue
            'Selection.Font.Bold = True
        End If
    End If
        
    'Enter all non-logical scenerios here:
    If Worksheets("Pipe Data").Range(FeatureColumn & CStr(lLastRowdDiameter)).Value < 0 Then
        MsgBox "No Feature in FVE Section"
        AutoPopulate = False
    Else
        AutoPopulate = True
    End If
PROC_EXIT:
    Exit Function
PROC_ERR:
    AutoPopulate = False
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in Auto Populate Function"
    Resume PROC_EXIT
End Function

Private Sub Worksheet_Change(ByVal Target As Range)

Dim errMsg As String

If validationStatus = True And Not blankRow(Target.row) Then
    errMsg = setValidation(Target)
    If errMsg <> "" Then
        MsgBox errMsg
    End If
End If


End Sub


