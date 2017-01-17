Attribute VB_Name = "FVEvalidation"
Sub turnRangeRed(rng As Range)
    rng.Font.ColorIndex = 3
End Sub

Function setValidation(rng As Range) As String

Dim fieldname As String
Dim subrng As Range
Dim errorMsg As String
Dim fieldArray() As Variant
Dim rangeName As String
Dim resultRange As Range
Dim subAddress As String
Dim installDate As Variant
Dim fittingMAOPrange As Range
Dim buildArray() As Variant
Dim MaxWorkingPressure_low As Variant
Dim MaxWorkingPressure_high As Variant


errorMsg = ""

fieldArray = Array("SMYS", "OD 1", "OD 2", "LS Factor", "Seam Type", "Fitting Rating", _
                "Installed CL", "Installed CL Design Factor", "Today's CL", "Today's CL Design Factor", _
                "Fitting MAOP", "Design Factor", "WT 1", "WT 2", _
                "Remove From MAOP Report ""R"" or ""D""", "Component")
                
MaxWorkingPressure_low = Range("=MaxWorkingPressure_low").Value
MaxWorkingPressure_high = Range("=MaxWorkingPressure_high").Value

                
For Each subrng In rng

    subAddress = subrng.AddressLocal(rowabsolute:=False, columnabsolute:=False)
    fieldname = Sheets("Pipe Data").Cells(2, subrng.Column)
    If (subrng.row > 2) And ((subrng.Column >= 102) Or (subrng.Column = 8)) Then
        If inArray(fieldname, fieldArray) And hasValidation(subrng) Then
            subrng.Validation.Delete
        End If
        Select Case fieldname
            Case "Component"
                If featureIsOther(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=ComponentFeatureType_FVE"
                Else
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=ComponentFeature_FVE"
                End If
            Case "SMYS"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=SMYS_FVE"
            Case "OD 1", "OD 2", "Feature"
                ODvalidation fieldname, subrng
            Case "LS Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=LSFactor_FVE"
            Case "Seam Type"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=SeamType_FVE"
            Case "Fitting Rating"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=FittingRating_FVE"
            Case "Installed CL"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=ClassLocation_FVE"
            Case "Installed CL Design Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=DesignFactor_FVE"
            Case "Today's CL"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=ClassLocationNoBlank_FVE"
            Case "Today's CL Design Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=DesignFactor_FVE"
            
            Case "Fitting MAOP"
            '    subrng.Validation.Add Type:=xlValidateList, Formula1:="=FittingMAOP_FVE"
                On Error Resume Next
                If isSkidMount(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=modelDynamic"
                ElseIf isHPR(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=HPRdynamic"
                ElseIf hasMaxPressure(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateDecimal, Operator:=xlBetween, Formula1:="=MaxWorkingPressure_low", Formula2:="=MaxWorkingPressure_high"
                Else
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=fittingDynamic"
                End If
                On Error GoTo 0
            Case "Design Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=DesignFactor_FVE"
            Case "Remove From MAOP Report ""R"" or ""D"""
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=RemoveFromMAOPReport_FVE"
            Case "WT 1"
                subrng.Validation.Add Type:=xlValidateDecimal, Formula1:=0.1, Formula2:=1.5, Operator:=xlBetween
            Case "WT 2"
                subrng.Validation.Add Type:=xlValidateDecimal, Formula1:=0.1, Formula2:=1.5, Operator:=xlBetween
        End Select
    End If
    
    'set validation for Fitting MAOP if change was made to Feature Type or Figure/Model#
    If (fieldname = "Type" Or fieldname = "Figure - Model #" Or fieldname = "Max Working Pressure") And (subrng.row > 2) Then
        Set fittingMAOPrange = subrng.Parent.Cells(subrng.row, fittingMAOPColumn)
        If hasValidation(fittingMAOPrange) Then
            fittingMAOPrange.Validation.Delete
        End If
        'On Error Resume Next
        If isSkidMount(subrng.row) Then
            fittingMAOPrange.Validation.Add Type:=xlValidateList, Formula1:="=modelDynamic"
        ElseIf isHPR(subrng.row) Then
            fittingMAOPrange.Validation.Add Type:=xlValidateList, Formula1:="=HPRdynamic"
        ElseIf hasMaxPressure(subrng.row) Then
            fittingMAOPrange.Validation.Add Type:=xlValidateDecimal, Operator:=xlBetween, Formula1:="=MaxWorkingPressure_low", Formula2:="=MaxWorkingPressure_high"
        Else
            fittingMAOPrange.Validation.Add Type:=xlValidateList, Formula1:="=fittingDynamic"
        End If
        'On Error GoTo 0
    End If
    
    
    If inArray(fieldname, fieldArray) And (subrng.row > 2) And (subrng.Column >= 102) Then
        'turnRangeRed subrng
        Select Case fieldname
            Case "WT 1"
                If IsNumeric(subrng.Value) Then
                    If subrng.Value < 0.1 Or subrng.Value > 1.5 Then
                        errorMsg = addError(errorMsg, fieldname, subrng)
                    End If
                ElseIf subrng.Value <> "N/A" Then
                    errorMsg = addError(errorMsg, fieldname, subrng)
                End If
            Case "WT 2"
                If IsNumeric(subrng.Value) Then
                    If subrng.Value < 0.1 Or subrng.Value > 1.5 Then
                        errorMsg = addError(errorMsg, fieldname, subrng)
                    End If
                ElseIf subrng.Value <> "N/A" Then
                    errorMsg = addError(errorMsg, fieldname, subrng)
                End If
            Case "Fitting MAOP"
                If hasMaxPressure(subrng.row) Then
                    If IsNumeric(subrng.Value) Then
                        If subrng.Value < MaxWorkingPressure_low Or subrng.Value > MaxWorkingPressure_high Then
                            errorMsg = addError(errorMsg, fieldname, subrng)
                        End If
                    ElseIf subrng.Value <> "N/A" Then
                        errorMsg = addError(errorMsg, fieldname, subrng)
                    End If
                Else
                    rangeName = Replace(subrng.Validation.Formula1, "=", "")
                    Set fittingMAOPrange = fittingMAOPcheck(subrng, rangeName)
                    If fittingMAOPrange Is Nothing Then
                        If Not subrng.Value = "" Then
                            errorMsg = addError(errorMsg, fieldname, subrng)
                        End If
                    Else
                        Set resultRange = fittingMAOPrange.Find(subrng.Value, lookat:=xlWhole)
                        If resultRange Is Nothing Then
                            errorMsg = addError(errorMsg, fieldname, subrng)
                        End If
                    End If
                End If
            Case Else
                rangeName = Replace(subrng.Validation.Formula1, "=", "")
                Set resultRange = Range(rangeName).Find(subrng.Value, lookat:=xlWhole)
                If resultRange Is Nothing Then
                    errorMsg = addError(errorMsg, fieldname, subrng)
                    If fieldname = "Remove From MAOP Report ""R"" or ""D""" Then
                        Application.EnableEvents = False
                        subrng.ClearContents
                        Application.EnableEvents = True
                        errorMsg = errorMsg & "Value deleted." & vbNewLine
                    End If
                End If
        End Select
        
        If errorMsg <> "" Then
            Application.EnableEvents = False
            subrng.ClearContents
            Application.EnableEvents = True
        End If
    End If
Next

setValidation = errorMsg

End Function


Function addError(errorMsg As String, fieldname As String, rng As Range)

If errorMsg = "" Then
    errorMsg = "Invalid data deleted from the following cells: " & vbNewLine
End If

addError = errorMsg & vbNewLine & "Field: " & fieldname & vbNewLine
addError = addError & "Value: " & rng.Value & vbNewLine
addError = addError & "Address: " & rng.Address & vbNewLine


End Function

Function inArray(searchVal As String, searchSet() As Variant) As Boolean

Dim i As Long
inArray = False
For i = 0 To UBound(searchSet)
    If UCase(searchVal) = UCase(searchSet(i)) Then
        inArray = True
        Exit Function
    End If
Next

End Function


Function hasValidation(cellobj As Range) As Boolean

    On Error Resume Next
        If cellobj.SpecialCells(xlCellTypeSameValidation).Cells.Count < 1 Then
            hasValidation = False
        Else
            hasValidation = True
        End If
    On Error GoTo 0
End Function

Sub initializeValidation()

Dim lastRow As Long
Dim sht As Worksheet
Dim validationRange As Range
Dim subrng As Range
Dim fieldname As String
Dim fieldArray() As Variant
Dim fittingMAOPrange As Range

'sample formula for dynamic validation:
'=OFFSET(INDIRECT(ADDRESS(MATCH(H2,OFFSET(A1,0,0,COUNTA(A:A),1),0),1)),0,1,COUNTIF(A:A,"="&H2),1)
'where H2 is the source cell, A1 and A:A are the source range
'can make H2 depend on current cell using =OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),yOffset,xOffset,1,1)
'backref with =fittingRatingRelative
'so:=OFFSET(INDIRECT(ADDRESS(MATCH(fittingRatingRelative,'FVE Validation'!$L$3:$L$50,0)+2,13,1,true,"FVE Validation")),0,0,COUNTIF('FVE Validation'!$L$3:$L$50,"="&fittingRatingRelative),1)


Set sht = Sheets("pipe data")
lastRow = sht.Range(addressForLastRow).End(xlUp).row
Set validationRange = sht.Range("cx3:fc" & lastRow)

fieldArray = Array("SMYS", "OD 1", "OD 2", "LS Factor", "Seam Type", "Fitting Rating", _
                "Installed CL", "Installed CL Design Factor", "Today's CL", "Today's CL Design Factor", _
                "Fitting MAOP", "Design Factor", "WT 1", "WT 2", _
                "Remove From MAOP Report ""R"" or ""D""", "Component")

For Each subrng In validationRange
    fieldname = sht.Cells(2, subrng.Column).Value
    If inArray(fieldname, fieldArray) Then
        subrng.Validation.Delete
    End If
    If (Not hasValidation(subrng)) And ((subrng.Column >= 102) Or (subrng.Column = 8)) Then
        Select Case fieldname
            Case "Component"
                If featureIsOther(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=ComponentFeatureType_FVE"
                Else
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=ComponentFeature_FVE"
                End If
            Case "SMYS"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=SMYS_FVE"
            Case "OD 1", "OD 2", "Feature"
                ODvalidation fieldname, subrng
            Case "LS Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=LSFactor_FVE"
            Case "Seam Type"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=SeamType_FVE"
            Case "Fitting Rating"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=FittingRating_FVE"
            Case "Installed CL"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=ClassLocation_FVE"
            Case "Installed CL Design Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=DesignFactor_FVE"
            Case "Today's CL"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=ClassLocationNoBlank_FVE"
            Case "Today's CL Design Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=DesignFactor_FVE"
            Case "Fitting MAOP"
            '    subrng.Validation.Add Type:=xlValidateList, Formula1:="=FittingMAOP_FVE"
                On Error Resume Next
                If isSkidMount(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=modelDynamic"
                ElseIf isHPR(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=HPRdynamic"
                ElseIf hasMaxPressure(subrng.row) Then
                    subrng.Validation.Add Type:=xlValidateDecimal, Operator:=xlBetween, Formula1:="=MaxWorkingPressure_low", Formula2:="=MaxWorkingPressure_high"
                Else
                    subrng.Validation.Add Type:=xlValidateList, Formula1:="=fittingDynamic"
                End If
                On Error GoTo 0
                
            Case "Design Factor"
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=DesignFactor_FVE"
            Case "Remove From MAOP Report ""R"" or ""D"""
                subrng.Validation.Add Type:=xlValidateList, Formula1:="=RemoveFromMAOPReport_FVE"
            Case "WT 1"
                subrng.Validation.Add Type:=xlValidateDecimal, Formula1:=0.1, Formula2:=1.5, Operator:=xlBetween
            Case "WT 2"
                subrng.Validation.Add Type:=xlValidateDecimal, Formula1:=0.1, Formula2:=1.5, Operator:=xlBetween
        End Select
    End If
Next


End Sub


Sub deleteValidation()

Dim lastRow As Long
Dim sht As Worksheet
Dim validationRange As Range
Dim subrng As Range
Dim fieldname As String
Dim fieldArray() As Variant



Set sht = Sheets("pipe data")
lastRow = sht.Range(addressForLastRow).End(xlUp).row
Set validationRange = sht.Range("cx3:ez" & lastRow)

fieldArray = Array("SMYS", "OD 1", "OD 2", "LS Factor", "Seam Type", "Fitting Rating", _
                "Installed CL", "Installed CL Design Factor", "Today's CL", "Today's CL Design Factor", _
                "Fitting MAOP", "Design Factor", "WT 1", "WT 2", _
                "Remove From MAOP Report ""R"" or ""D""")
                
                
For Each subrng In validationRange
    fieldname = sht.Cells(2, subrng.Column).Value
    If inArray(fieldname, fieldArray) Then
        subrng.Validation.Delete
    End If
Next


End Sub
Sub testprop()
'setValidationStatus (False)
Debug.Print validationStatus

End Sub

Function validationStatus() As Boolean
Dim sht As Worksheet
Dim cprop As CustomProperty

Set sht = Sheets("fve validation")
For Each cprop In sht.CustomProperties
    If cprop.Name = "validationStatus" Then
        validationStatus = cprop.Value
        Exit Function
    End If
Next


End Function

Sub setValidationStatus(Value As Boolean)
Dim sht As Worksheet
Dim cprop As CustomProperty

Set sht = Sheets("fve validation")
For Each cprop In sht.CustomProperties
    If cprop.Name = "validationStatus" Then
        cprop.Value = Value
        Exit Sub
    End If
Next

End Sub

'need to figure out how to trim @crlf from isHPR result
Function isHPR(row As Long) As Boolean
Dim sht As Worksheet
Dim feature As Variant
Dim featureType As Variant
Dim HPRval As Variant


Set sht = Sheets("Pipe Data")
feature = sht.Cells(row, buildFeatureColumn)
featureType = sht.Cells(row, buildFeatureTypeColumn)

If feature = "FarmTapRegSet" Then
    HPRval = featureType
Else
    HPRval = ""
End If
isHPR = Len(HPRval) > 0

End Function

Function installDate(row As Long) As Variant

Dim sht As Worksheet

Set sht = Sheets("Pipe Data")

installDate = sht.Cells(row, installDateCol)


End Function

Function isSkidMount(row As Long) As Boolean
Dim sht As Worksheet
Dim feature As Variant
Dim featureType As Variant

Set sht = Sheets("Pipe Data")
feature = sht.Cells(row, buildFeatureColumn)
featureType = sht.Cells(row, buildFeatureTypeColumn)

isSkidMount = ((feature = "Meter") And ((featureType = "Orifice Skid Mnt - Gas Well") Or (featureType = "OrificeSkidMntGasWell")))

End Function

Function featureIsOther(row As Long) As Boolean
Dim sht As Worksheet
Dim feature As Variant

Set sht = Sheets("Pipe Data")
feature = sht.Cells(row, buildFeatureColumn)
featureIsOther = (feature = "Other")

End Function
Function hasMaxPressure(row As Long) As Boolean
Dim sht As Worksheet
Dim maxWorkingPressure As Variant

Set sht = Sheets("Pipe Data")
maxWorkingPressure = sht.Cells(row, maxWorkingPressureColumn)
hasMaxPressure = maxWorkingPressure <> ""

End Function

Sub testInstall()
Dim datevar As Variant
datevar = installDate(3)
Debug.Print datevar
Debug.Print datevar < #1/2/2013#
End Sub

Function lastColumn(sht As Worksheet)
    lastColumn = sht.Range(addressForLastColumn).End(xlToLeft).Column
End Function

Function blankRow(row As Long) As Boolean
    Dim sht As Worksheet
    Dim rng As Range
    Dim subrng As Range
    Set sht = Sheets("Pipe Data")
    Set rng = sht.Range(sht.Cells(row, 1), sht.Cells(row, lastColumn(sht)))
    For Each subrng In rng
        If VarType(subrng.Value) = vbError Then
            blankRow = False
            Exit Function
        End If
        If subrng.Value <> "" Then
            blankRow = False
            Exit Function
        End If
    Next
    blankRow = True
End Function

Function fittingMAOPcheck(rng As Range, namedRange As String) As Range

Dim fittingRatingRange As Range
Dim fittingRatingSubRange As Range
Dim startRow As Long
Dim endRow As Long
Dim rowOffset As Long
Dim startMAOP As Range
Dim endMAOP As Range
Dim valueOffset As Long
Dim validationTable As String


Select Case namedRange
    Case "modelDynamic"
        valueOffset = -102
        validationTable = "SMMS_FVE"
    Case "HPRdynamic"
        valueOffset = -120
        validationTable = "HPR_FVE"
    Case "fittingDynamic"
        valueOffset = -10
        validationTable = "FittingRating_FVE"
End Select


rowOffset = 0


Set fittingRatingRange = Range(validationTable)
Set fittingRatingSubRange = fittingRatingRange.Find(rng.Offset(0, valueOffset).Value)
If Not fittingRatingSubRange Is Nothing Then
    startRow = fittingRatingSubRange.row
    While fittingRatingSubRange.Offset(rowOffset, 0).Value = fittingRatingSubRange.Value
        rowOffset = rowOffset + 1
    Wend
    With ThisWorkbook.Sheets("FVE Validation")
        Set startMAOP = .Cells(fittingRatingSubRange.row, fittingRatingSubRange.Column + 1)
        Set endMAOP = .Cells(fittingRatingSubRange.row + rowOffset - 1, fittingRatingSubRange.Column + 1)
        Set fittingMAOPcheck = .Range(startMAOP, endMAOP)
    End With
Else
    Set fittingMAOPcheck = Nothing
End If

End Function

Function fittingMAOPresultRange(subrng As Range) As Range
    Dim validationRange As Range
    Set validationRange = fittingMAOPcheck(subrng)
    Set resultRange = subrng.Find(tempValue, lookat:=xlWhole)
End Function


Sub ODvalidation(fieldname As Variant, rng As Range)

Select Case fieldname
    Case "Feature"
        Select Case rng.Value
            Case "Sleeve"
                rng.Offset(columnoffset:=feature_OD1_offset).Validation.Delete
                rng.Offset(columnoffset:=feature_OD1_offset).Validation.Add Type:=xlValidateList, Formula1:="=ODlong_FVE"
                rng.Offset(columnoffset:=feature_OD2_offset).Validation.Delete
                rng.Offset(columnoffset:=feature_OD2_offset).Validation.Add Type:=xlValidateList, Formula1:="=ODlong_FVE"
            Case Else
                rng.Offset(columnoffset:=feature_OD1_offset).Validation.Delete
                rng.Offset(columnoffset:=feature_OD1_offset).Validation.Add Type:=xlValidateList, Formula1:="=ODshort_FVE"
                rng.Offset(columnoffset:=feature_OD2_offset).Validation.Delete
                rng.Offset(columnoffset:=feature_OD2_offset).Validation.Add Type:=xlValidateList, Formula1:="=ODshort_FVE"
        End Select
    Case "OD 1"
        Select Case rng.Offset(columnoffset:=(-1 * feature_OD1_offset)).Value
            Case "Sleeve"
                rng.Validation.Add Type:=xlValidateList, Formula1:="=ODlong_FVE"
            Case Else
                rng.Validation.Add Type:=xlValidateList, Formula1:="=ODshort_FVE"
        End Select
    Case "OD 2"
        Select Case rng.Offset(columnoffset:=(-1 * feature_OD2_offset)).Value
            Case "Sleeve"
                rng.Validation.Add Type:=xlValidateList, Formula1:="=ODlong_FVE"
            Case Else
                rng.Validation.Add Type:=xlValidateList, Formula1:="=ODshort_FVE"
        End Select

End Select
      
End Sub
