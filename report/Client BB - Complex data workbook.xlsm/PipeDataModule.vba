Attribute VB_Name = "PipeDataModule"
Option Explicit  ' Causes error if variable is not defined with Dim
Private Const ERROR_YIELD As Long = 99999999
' Use to highlight rows a different color
' see function subHighlightRow()
Public Enum ROW_HIGHLIGHT
    Highlight_Off = 1
    Highlight_RED = 2
End Enum
Global Const MAX_SAVE_ARRAY As Integer = 3000
Global vSaveSTPRArray(MAX_SAVE_ARRAY, 15) As Variant
Global lSaveSTRPArrayHighNumber As Long
Global vSaveJobArray(MAX_SAVE_ARRAY) As Variant
Global lSaveJobArrayHighNumber As Long
Public Function ColorText(CellToColor As String, Color As String, Value As String) As Boolean

    Dim CellLength As Integer
    Dim StartPosition As Integer
    Dim ColorLength As Integer
    
    CellLength = Len(Range(CellToColor).Value)
    ColorLength = Len(Value)
    StartPosition = CellLength - ColorLength
    
    Range(CellToColor).Select
    With ActiveCell.Characters(Start:=StartPosition, Length:=ColorLength).Font
                            '.Name = "Arial"
                            '.FontStyle = "Bold"
                            '.Size = 8
                            '.Strikethrough = False
                            '.Superscript = False
                            '.Subscript = False
                            '.OutlineFont = False
                            '.Shadow = False
                            '.Underline = xlUnderlineStyleNone
                            If Color = "Red" Or Color = "red" Then
                                .Color = -16776961
                            ElseIf Color = "Blue" Or Color = "blue" Then
                                .Color = -100520
                            ElseIf Color = "DarkBlue" Or Color = "Darkblue" Or Color = "darkblue" Then
                                .Color = RGB(0, 0, 255)
                            ElseIf Color = "Purple" Or Color = "purple" Then
                                .Color = -6279056
                            ElseIf Color = "DarkPurple" Or "darkpurple" Or "Darkpurple" Then
                                .Color = -9174924
                            End If
                            '.TintAndShade = 0
                            '.ThemeFont = xlThemeFontNone
    End With
    ColorText = True
End Function

Public Sub SaveSTPRRowNumber(p_bInitialize As Boolean, p_lSaveRowNumber As Long, p_sTabName As String)
    Dim k As Integer
    If p_bInitialize Then
        ' Intializeing Array ONLY
        lSaveSTRPArrayHighNumber = 0
        Debug.Print "Initialize Array"
    Else
        ' Save Row number that was modified
        lSaveSTRPArrayHighNumber = lSaveSTRPArrayHighNumber + 1
        For k = 1 To 11
            vSaveSTPRArray(lSaveSTRPArrayHighNumber, k) = Worksheets(p_sTabName).Range("E" & Chr(64 + k) & p_lSaveRowNumber).Value
        Next k
    End If
End Sub
Public Sub SaveJobowNumber(p_bInitialize As Boolean, p_lSaveRowNumber As Long, p_sTabName As String)
    Dim k As Integer
    If p_bInitialize Then
        ' Intializeing Array ONLY
        lSaveSTRPArrayHighNumber = 0
        Debug.Print "Initialize Array"
    Else
        ' Save Row number that was modified
        lSaveSTRPArrayHighNumber = lSaveSTRPArrayHighNumber + 1
        For k = 1 To 4
            vSaveSTPRArray(lSaveSTRPArrayHighNumber, k) = Worksheets(p_sTabName).Range("E" & Chr(75 + k) & p_lSaveRowNumber).Value
        Next k
    End If
End Sub ' --------------------------------------------------
' Comments: Get Yield from Tab 'Tb2a2 DATA'
' --------------------------------------------------
Sub CalcYield_2A2()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 12        ' Table starts at this row
   iNumColumns = 12  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("O12").Value
   iPipeSize = Worksheets(sTabName).Range("O13").Value
   sSeamType = Worksheets(sTabName).Range("O14").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 12, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("O15").Value = lYield
End Sub
' --------------------------------------------------
' Comments: Get Yield from Tab 'Tbl_2b3 DATA'
' --------------------------------------------------
Sub CalcYield_2b3()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 34        ' Table starts at this row
   iNumColumns = 11  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("N35").Value
   iPipeSize = Worksheets(sTabName).Range("N36").Value
   sSeamType = Worksheets(sTabName).Range("N37").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 11, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("N38").Value = lYield
End Sub
' --------------------------------------------------
' Comments: Get Yield from Tab 'Tbl2b4_1 DATA'
' --------------------------------------------------
Sub CalcYield_2b4_1()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 51        ' Table starts at this row
   iNumColumns = 6  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("I50").Value
   iPipeSize = Worksheets(sTabName).Range("I51").Value
   sSeamType = Worksheets(sTabName).Range("I52").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 6, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("I53").Value = lYield
End Sub
' --------------------------------------------------
' Comments: Get Yield from Tab 'Tbl2b4_2 DATA'
' --------------------------------------------------
Sub CalcYield_2b4_2()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 58        ' Table starts at this row
   iNumColumns = 4  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("G57").Value
   iPipeSize = Worksheets(sTabName).Range("G58").Value
   sSeamType = Worksheets(sTabName).Range("G59").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 4, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("G60").Value = lYield
End Sub
' --------------------------------------------------
' Comments: Get Yield from Tab '2b4_Unknown_1 DATA'
' --------------------------------------------------
Sub CalcYield_2b4_Unknown_1()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 67        ' Table starts at this row
   iNumColumns = 11  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("N67").Value
   iPipeSize = Worksheets(sTabName).Range("N68").Value
   sSeamType = Worksheets(sTabName).Range("N69").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 11, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("N70").Value = lYield
End Sub

' --------------------------------------------------
' Comments: Get Yield from Tab '2b4_Unknown_2 DATA'
' --------------------------------------------------
Sub CalcYield_2b4_Unknown_2()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 76        ' Table starts at this row
   iNumColumns = 12  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("O76").Value
   iPipeSize = Worksheets(sTabName).Range("O77").Value
   sSeamType = Worksheets(sTabName).Range("O78").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 12, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("O79").Value = lYield
End Sub

' --------------------------------------------------
' Comments: Get Yield from Tab 'Tbl2b4_3 DATA'
' --------------------------------------------------
Sub CalcYield_2b4_3()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 85        ' Table starts at this row
   iNumColumns = 4  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("G84").Value
   iPipeSize = Worksheets(sTabName).Range("G85").Value
   sSeamType = Worksheets(sTabName).Range("G86").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 4, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("G87").Value = lYield
End Sub
' --------------------------------------------------
' Comments: Get Yield from Tab 'Tbl2b4_4 DATA'
' --------------------------------------------------
Sub CalcYield_2b4_4()
   Dim lYield As Long, iPipeSize As Double, sSeamType As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   Dim sOutput As String
   iStartingRow = 101        ' Table starts at this row
   iNumColumns = 10  ' Number of columns in table
   sTabName = "Logic" 'Name of sheet here
   sOutput = "Pipe Data"
   vPurchaseDate = Worksheets(sTabName).Range("M102").Value
   iPipeSize = Worksheets(sTabName).Range("M103").Value
   sSeamType = Worksheets(sTabName).Range("M104").Value
   lYield = fnxDetermineYield(sTabName, iStartingRow, 10, iPipeSize, sSeamType, vPurchaseDate, 1)
   Worksheets(sTabName).Range("M105").Value = lYield
End Sub
' --------------------------------------------------
' Comments: Get Thickness from Tab 'Tbl3 PipeWTCurrentLogic'
' --------------------------------------------------
Sub CalcThickness_3_2()
   Dim vThickness As Variant, iDiameter As Double, sLogic As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   iStartingRow = 145       ' Table starts at this row
   iNumColumns = 8         ' Number of columns in table
   sTabName = "Logic"
   vPurchaseDate = Worksheets(sTabName).Range("K145").Value
   iDiameter = Worksheets(sTabName).Range("K146").Value
   sLogic = Worksheets(sTabName).Range("K147").Value
   ' Only 8 columns
   vThickness = fnxDetermineThickness(sTabName, iStartingRow, 8, iDiameter, sLogic, vPurchaseDate, 1)
   Worksheets(sTabName).Range("K148").Value = vThickness
End Sub
' --------------------------------------------------
' Comments: Get Thickness from Tab 'Tbl3 PipeWTCurrentLogic'
' --------------------------------------------------
Sub CalcThickness_5B_SEAM()
   Dim vThickness As Variant, iDiameter As Double, sLogic As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer
   iStartingRow = 187        ' Table starts at this row
   iNumColumns = 5         ' Number of columns in table
   sTabName = "Logic"
   vPurchaseDate = Worksheets(sTabName).Range("H187").Value
   iDiameter = Worksheets(sTabName).Range("H188").Value
   sLogic = Worksheets(sTabName).Range("H190").Value
   ' Only 8 columns
   vThickness = fnxDetermineThickness(sTabName, iStartingRow, 5, iDiameter, sLogic, vPurchaseDate, 1)
   Worksheets(sTabName).Range("H189").Value = vThickness

End Sub
' --------------------------------------------------
' Comments: Get Thickness from Tab 'Tbl3 PipeWTCurrentLogic'
' --------------------------------------------------
' START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
Sub CalcThickness_UnkDate()
   Dim vThickness As Variant, iDiameter As Double, sLogic As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer, sStartColumn As String, sEndColumn As String
   iStartingRow = 206        ' Table starts at this row
   iNumColumns = 3         ' Number of columns in table
   sTabName = "Logic"
   sStartColumn = "A"
   sEndColumn = "B"
   vPurchaseDate = Worksheets(sTabName).Range("F207").Value
   iDiameter = Worksheets(sTabName).Range("F208").Value
   sLogic = Worksheets(sTabName).Range("F209").Value
   ' Only 8 columns
   vThickness = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn)
   Worksheets(sTabName).Range("F210").Value = vThickness

End Sub
' --------------------------------------------------
' Comments: Get Thickness from Tab 'Tbl3 PipeWTCurrentLogic'
' --------------------------------------------------
' START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
Sub CalcThickness_1st()
   Dim vThickness As Variant, iDiameter As Double, sLogic As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer, sStartColumn As String, sEndColumn As String
   iStartingRow = 233        ' Table starts at this row
   iNumColumns = 3         ' Number of columns in table
   sTabName = "Logic"
   sStartColumn = "A"
   sEndColumn = "B"
   vPurchaseDate = Worksheets(sTabName).Range("F234").Value
   iDiameter = Worksheets(sTabName).Range("F235").Value
   sLogic = Worksheets(sTabName).Range("F236").Value
   ' Only 8 columns
   vThickness = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn)
   Worksheets(sTabName).Range("F237").Value = vThickness

End Sub
' --------------------------------------------------
' Comments: Get Thickness from Tab 'Tbl3 PipeWTCurrentLogic'
' --------------------------------------------------
' START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
Sub CalcThickness_2nd()
   Dim vThickness As Variant, iDiameter As Double, sLogic As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer, sStartColumn As String, sEndColumn As String
   iStartingRow = 260        ' Table starts at this row
   iNumColumns = 3         ' Number of columns in table
   sTabName = "Logic"
   sStartColumn = "A"
   sEndColumn = "B"
   vPurchaseDate = Worksheets(sTabName).Range("F261").Value
   iDiameter = Worksheets(sTabName).Range("F262").Value
   sLogic = Worksheets(sTabName).Range("F263").Value
   ' Only 8 columns
   vThickness = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn)
   Worksheets(sTabName).Range("F264").Value = vThickness

End Sub
' --------------------------------------------------
' Comments: Get Thickness from Tab 'Tbl3 PipeWTCurrentLogic'
' --------------------------------------------------
' START THE iStartingRow 1 Row above the Titles! (2 rows above the data)
Sub CalcThickness_3rd()
   Dim vThickness As Variant, iDiameter As Double, sLogic As String, vPurchaseDate As Variant
   Dim sTabName As String, iStartingRow As Integer, iNumColumns As Integer, sStartColumn As String, sEndColumn As String
   iStartingRow = 287        ' Table starts at this row
   iNumColumns = 3         ' Number of columns in table
   sTabName = "Logic"
   sStartColumn = "A"
   sEndColumn = "B"
   vPurchaseDate = Worksheets(sTabName).Range("F288").Value
   iDiameter = Worksheets(sTabName).Range("F289").Value
   sLogic = Worksheets(sTabName).Range("F290").Value
   ' Only 8 columns
   vThickness = fnxDetermineThickness1(sTabName, iStartingRow, 3, iDiameter, sLogic, vPurchaseDate, sStartColumn, sEndColumn)
   Worksheets(sTabName).Range("F291").Value = vThickness
End Sub
' --------------------------------------------------
' Comments: fnxDetermineYield -
' Params  : p_sWorksheet name of worksheet (i.e. Pipe Data);
'
' Returns : long (-1 implies error, 0 implies cannot be determined)
' Created : 10/25/2012 example
' --------------------------------------------------
Public Function fnxDetermineYield(p_sWorksheet As String, p_iStartingRow As Integer, p_iNumColumns As Integer, _
        p_iPipeSize As Double, p_sSeamType As String, p_vPurchasDate As Variant, i_UseDateShift As Integer) As Long
    On Error GoTo PROC_ERR
    Dim PossibleYields(100) As Long, iYieldIndex As Integer, lBestYield As Long
    Dim iCurrentRow As Integer, iRowFound As Integer, sPipeCellValue As String, iCurColumn As Integer
    Dim dMinDate As Date, dMaxDate As Date, vDateVal As Variant, dParamDate As Date, idx As Integer
    ' default answer for Yield is 0 .. which means cannot be determined
    lBestYield = 0
    iRowFound = 0
    ' The first row of data starts 2 rows after the header
    iCurrentRow = p_iStartingRow + 2
    ' First find mathcing Pipe Size / Seam Type Row
    Do While Len(Trim(Worksheets(p_sWorksheet).Range("A" & Trim(Str(iCurrentRow))).Value)) > 0
        sPipeCellValue = Worksheets(p_sWorksheet).Range("A" & Trim(Str(iCurrentRow))).Value
        ' Check to ensure we're looking at a numeric value
        If IsNumeric(sPipeCellValue) Then
            ' See if paramater matches spreadsheet
            If Val(sPipeCellValue) = p_iPipeSize Then
                ' Now see if Seam Type matches
                If p_sSeamType = Trim(Worksheets(p_sWorksheet).Range("B" & Trim(Str(iCurrentRow))).Value) Then
                    iRowFound = iCurrentRow
                    Exit Do
                End If
            End If
        End If
        iCurrentRow = iCurrentRow + 1
    Loop
    ' Found matching Pipe Size and Seam Type
    If iRowFound > 0 Then
        ' Now that we found the right row .. time to get the yield
        ' .. get all possible yields and store in array PossibleYields
        ' .. then get minimum value of all Yields
        iYieldIndex = 0
        ' Unknown date is translated into 1/1/1900
        If IsDate(p_vPurchasDate) Then
            dParamDate = p_vPurchasDate
        Else
            dParamDate = #1/1/1900#
        End If
        ' See if Purchase Date is within Range of Install Dates
        For iCurColumn = 3 To p_iNumColumns
            ' Chr(65) = "A"
            ' Get Min Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(p_iStartingRow))).Value
            If IsDate(vDateVal) Then
                dMinDate = DateValue(vDateVal)
            Else
                dMinDate = #1/1/1900#
            End If
            ' Get Max Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(p_iStartingRow + 1))).Value
            If IsDate(vDateVal) Then
                ' Add 10 years to compensate for Install Datee
                ' dMaxDate = DateAdd("YYYY", 10, DateValue(vDateVal))
                dMaxDate = DateValue(vDateVal)
            Else
                dMaxDate = #1/1/2100#
            End If
            ' Failing here
            If (dParamDate >= dMinDate And dParamDate <= dMaxDate) Or _
                DateAdd("d", -1, DateAdd("YYYY", 10, DateValue(dParamDate))) >= dMinDate And dParamDate <= dMaxDate Then
                'DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate Then
                ' This cell is in range
                iYieldIndex = iYieldIndex + 1
                PossibleYields(iYieldIndex) = Val(Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(iRowFound))).Value)
            End If
        Next iCurColumn
        ' Find lowest value in Array .. Use 1st selection as starting point
        lBestYield = IIf(PossibleYields(1) = 0, ERROR_YIELD, PossibleYields(1))
        ' Loop through all possible Yiels values .. select lowest value
        If i_UseDateShift = 1 Then
            For idx = 1 To iYieldIndex
                If IIf(PossibleYields(idx) = 0, ERROR_YIELD, PossibleYields(idx)) < lBestYield Then
                    lBestYield = PossibleYields(idx)
                End If
            Next idx
            fnxDetermineYield = IIf(lBestYield = ERROR_YIELD, 0, lBestYield)
        Else
            fnxDetermineYield = IIf(lBestYield = ERROR_YIELD, 0, lBestYield)
        End If
    Else
        fnxDetermineYield = 0
    End If
PROC_EXIT:
    Exit Function
PROC_ERR:
    fnxDetermineYield = -1
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in fnxDetermineYield"
    Resume PROC_EXIT
End Function
Public Function xTestYield() As Long
  xTestYield = fnxDetermineYield("Tbl2 PipeYSCurrentLogic_Mas", 37, 10, 24, "DSAW", #1/1/1980#, 1)
End Function
' --------------------------------------------------
' Comments: fnxDetermineThickness -
' Params  : p_sWorksheet name of worksheet (i.e. Tbl3 PipeWTCurrentLogic);
'
' Returns : long (-1 implies error, 0 implies cannot be determined)
' Created : 10/25/2012 example
' --------------------------------------------------
Public Function fnxDetermineThickness(p_sWorksheet As String, p_iStartingRow As Integer, p_iNumColumns As Integer, _
        p_iDiameter As Double, p_sLogic As Variant, p_vPurchasDate As Variant, i_UseDateShift As Integer) As Variant
    On Error GoTo PROC_ERR
    Dim PossibleThickness(100) As Variant, iDiamIndex As Integer, vBestDiameter As Variant
    Dim iCurrentRow As Integer, iRowFound As Integer, sDiamCellValue As String, iCurColumn As Integer
    Dim dMinDate As Date, dMaxDate As Date, vDateVal As Variant, dParamDate As Date, idx As Integer
    Dim dThickness As Double
    Dim Test As String
    ' default answer for Yield is 0 .. which means cannot be determined
    vBestDiameter = 0
    iRowFound = 0
    ' The first row of data starts 2 rows after the header
    iCurrentRow = p_iStartingRow + 2
    ' First find mathcing Diameter / Seam Type Row
    Do While Len(Trim(Worksheets(p_sWorksheet).Range("A" & Trim(Str(iCurrentRow))).Value)) > 0
        sDiamCellValue = Worksheets(p_sWorksheet).Range("A" & Trim(Str(iCurrentRow))).Value
        ' Check to ensure we're looking at a numeric value
        If IsNumeric(sDiamCellValue) Then
            ' See if paramater matches spreadsheet
            If Val(sDiamCellValue) = p_iDiameter Then
                ' Now see if Seam Type matches (p_sLogic is the Seam)
                If p_sLogic = Trim(Worksheets(p_sWorksheet).Range("B" & Trim(Str(iCurrentRow))).Value) Or _
                        Len(Trim(Worksheets(p_sWorksheet).Range("B" & Trim(Str(iCurrentRow))).Value)) = 0 Then
                    iRowFound = iCurrentRow
                    Exit Do
                End If
            End If
        End If
        iCurrentRow = iCurrentRow + 1
    Loop
    ' Found matchin Pipe Size and Seam Type
    If iRowFound > 0 Then
        ' Now that we found the right row .. time to get the yield
        ' .. get all possible yields and store in array PossibleThickness
        ' .. then get minimum value of all Yields
        iDiamIndex = 0
        ' Unknown date is translated into 1/1/1900
        If IsDate(p_vPurchasDate) Then
            dParamDate = p_vPurchasDate
        Else
            dParamDate = #1/1/1800#
        End If
        
        ' See if Purchase Date is within Range of Install Dates
        For iCurColumn = 3 To p_iNumColumns
            '''''
            ' Get Min/Max values in Columns
            '''''
            ' Chr(65) = "A"
            ' Get Min Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(p_iStartingRow))).Value
            If IsDate(vDateVal) Then
                dMinDate = DateValue(vDateVal)
            Else
                dMinDate = #1/1/1800#
            End If
            ' Get Max Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(p_iStartingRow + 1))).Value
            If IsDate(vDateVal) Then
                ' Add 10 years to compensate for Install Datee
                dMaxDate = DateValue(vDateVal)
            Else
                dMaxDate = #1/1/2100#
            End If
            '''''''
            ' Determine if Purchase date is within bounds of min/max dates
            '''''''
            If (dParamDate >= dMinDate And dParamDate <= dMaxDate) Or _
                DateAdd("d", -1, DateAdd("YYYY", 10, DateValue(dParamDate))) >= dMinDate And dParamDate <= dMaxDate Then
                'DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate Then
                'DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate Then
                ' This cell is in range
                iDiamIndex = iDiamIndex + 1
                PossibleThickness(iDiamIndex) = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(iRowFound))).Value
            End If
        Next iCurColumn
        ' Find lowest value in Array .. Use 1st selection as starting point
        vBestDiameter = PossibleThickness(1)
        ' Look for miniumum .. if alpha selected .. then always take alpha .. no minimum needed
        'If i_UseDateShift = 1 then choose lowest value within range, otherwise just choose single value
        If i_UseDateShift = 1 Then
            If IsNumeric(vBestDiameter) Then
                ' Loop through all possible Thickness values
                ' .. field investig will always win
                For idx = 1 To iDiamIndex
                    ' Always choose non numeric thichkness
                    If IsNumeric(PossibleThickness(idx)) Then
                        If CDec(PossibleThickness(idx)) < vBestDiameter Then
                            vBestDiameter = PossibleThickness(idx)
                        End If
                    Else
                        ' field investig always selected
                        ' vBestDiameter = PossibleThickness(idx)
                        '
                        ' Exit For
                    End If
                Next idx
            End If
        End If
        fnxDetermineThickness = vBestDiameter
    Else
        fnxDetermineThickness = 0
    End If
PROC_EXIT:
    Exit Function
PROC_ERR:
    fnxDetermineThickness = -1
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in fnxDetermineThickness"
    Resume PROC_EXIT
End Function
' --------------------------------------------------
' Comments: fnxDetermineThickness -
' Params  : p_sWorksheet name of worksheet (i.e. Tbl3 PipeWTCurrentLogic);
'
' Returns : long (-1 implies error, 0 implies cannot be determined)
' Created : 08/15/2011 example
' --------------------------------------------------
Public Function fnxDetermineThickness1(p_sWorksheet As String, p_iStartingRow As Integer, p_iNumColumns As Integer, _
        p_iDiameter As Double, p_sLogic As Variant, p_vPurchasDate As Variant, sStartColumn As String, sEndColumn As String) As Variant
    On Error GoTo PROC_ERR
    Dim PossibleThickness(100) As Variant, iDiamIndex As Integer, vBestDiameter As Variant
    Dim iCurrentRow As Integer, iRowFound As Integer, sDiamCellValue As String, iCurColumn As Integer
    Dim dMinDate As Date, dMaxDate As Date, vDateVal As Variant, dParamDate As Date, idx As Integer
    Dim dThickness As Double
    Dim Test As String
    ' default answer for Yield is 0 .. which means cannot be determined
    vBestDiameter = 0
    iRowFound = 0
    ' The first row of data starts 2 rows after the header
    iCurrentRow = p_iStartingRow + 2
    ' First find mathcing Diameter / Seam Type Row
    Do While Len(Trim(Worksheets(p_sWorksheet).Range("A" & Trim(Str(iCurrentRow))).Value)) > 0
        sDiamCellValue = Worksheets(p_sWorksheet).Range("A" & Trim(Str(iCurrentRow))).Value
        ' Check to ensure we're looking at a numeric value
        If IsNumeric(sDiamCellValue) Then
            ' See if paramater matches spreadsheet
            If Val(sDiamCellValue) = p_iDiameter Then
                ' Now see if Seam Type matches (p_sLogic is the Seam)
                If p_sLogic = Trim(Worksheets(p_sWorksheet).Range("B" & Trim(Str(iCurrentRow))).Value) Or _
                        Len(Trim(Worksheets(p_sWorksheet).Range("B" & Trim(Str(iCurrentRow))).Value)) = 0 Then
                    iRowFound = iCurrentRow
                    Exit Do
                End If
            End If
        End If
        iCurrentRow = iCurrentRow + 1
    Loop
    ' Found matchin Pipe Size and Seam Type
    If iRowFound > 0 Then
        ' Now that we found the right row .. time to get the yield
        ' .. get all possible yields and store in array PossibleThickness
        ' .. then get minimum value of all Yields
        iDiamIndex = 0
        ' Unknown date is translated into 1/1/1900
        If IsDate(p_vPurchasDate) Then
            dParamDate = p_vPurchasDate
        Else
            dParamDate = #1/1/1800#
        End If
        ' See if Purchase Date is within Range of Install Dates
        For iCurColumn = 3 To p_iNumColumns
            '''''
            ' Get Min/Max values in Columns
            '''''
            ' Chr(65) = "A"
            ' Get Min Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(p_iStartingRow))).Value
            If IsDate(vDateVal) Then
                dMinDate = DateValue(vDateVal)
            Else
                dMinDate = #1/1/1800#
            End If
            ' Get Max Date
            vDateVal = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(p_iStartingRow + 1))).Value
            If IsDate(vDateVal) Then
                ' Add 10 years to compensate for Install Datee
                dMaxDate = DateValue(vDateVal)
            Else
                dMaxDate = #1/1/2100#
            End If
            '''''''
            ' Determine if Purchase date is within bounds of min/max dates
            '''''''
            If (dParamDate >= dMinDate And dParamDate <= dMaxDate) Or _
                DateAdd("d", -1, DateAdd("YYYY", 10, DateValue(dParamDate))) >= dMinDate And dParamDate <= dMaxDate Then
                'DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate Then
                'DateAdd("YYYY", 10, DateValue(dParamDate)) >= dMinDate And dParamDate <= dMaxDate Then
                ' This cell is in range
                iDiamIndex = iDiamIndex + 1
                PossibleThickness(iDiamIndex) = Worksheets(p_sWorksheet).Range(Chr(64 + iCurColumn) & Trim(Str(iRowFound))).Value
            End If
        Next iCurColumn
        ' Find lowest value in Array .. Use 1st selection as starting point
        vBestDiameter = PossibleThickness(1)
        ' Look for miniumum .. if alpha selected .. then always take alpha .. no minimum needed
        If IsNumeric(vBestDiameter) Then
            ' Loop through all possible Thickness values
            ' .. field investig will always win
            For idx = 1 To iDiamIndex
                ' Always choose non numeric thichkness
                If IsNumeric(PossibleThickness(idx)) Then
                    If CDec(PossibleThickness(idx)) < vBestDiameter Then
                        vBestDiameter = PossibleThickness(idx)
                    End If
                Else
                    ' field investig always selected
                    ' vBestDiameter = PossibleThickness(idx)
                    '
                    ' Exit For
                End If
            Next idx
        End If
        fnxDetermineThickness1 = vBestDiameter
    Else
        fnxDetermineThickness1 = 0
    End If
PROC_EXIT:
    Exit Function
PROC_ERR:
    fnxDetermineThickness1 = -1
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in fnxDetermineThickness1"
    Resume PROC_EXIT
End Function

Public Function xTestThick() As Variant
' Public Function fnxDetermineThickness(p_sWorksheet As String, p_iStartingRow As Integer, p_iNumColumns As Integer, _
'        p_iDiameter As Integer, p_sLogic As String, p_vPurchasDate As Variant) As Variant
  xTestThick = fnxDetermineThickness("Tbl3 PipeWTCurrentLogic_Mas", 46, 8, 18, "", #1/1/1950#, 1)
End Function

'Option Explicit     ' Causes error if variable is not defined with Dim
''''''''''''''''''''''Basic Functions'''''''''''''''''''''''
' --------------------------------------------------
' Comments: fnxIsItIn - Return boolean if 1st paramater is contained in next set of paramaters
' Params  : p_lTest is the var to be tested,
'           ParamArray vars is the test set
' Returns : Boolean
' Created : 08/15/2011 example
' --------------------------------------------------
Public Function fnxIsItIn(p_lTest As Variant, ParamArray a_Test()) As Boolean
    On Error GoTo PROC_ERR
    Dim lIncrement As Variant
    fnxIsItIn = False   ' Default that p_lTest is not in ParamArray
    ' Cycle through ALL lon integers in paramater a_Test
    For Each lIncrement In a_Test
        ' If test va p_lTest is contained in any of the elements of a_Test, return true
        If p_lTest = lIncrement Then
            fnxIsItIn = True
            Exit Function
        End If
    Next lIncrement
    ' All variables have been tested
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in fnxIsItIn"
    Resume PROC_EXIT
End Function
Public Function fnxTestForNull() As Boolean
    Dim a
    a = Worksheets("Pipe Data").Range("O20").Value
    Debug.Print a
    If Len(Trim(a)) = 0 Then
        fnxTestForNull = True
    Else
        fnxTestForNull = False
    End If
End Function
' --------------------------------------------------
' Comments: fnxFindLastRow - Return Last Row, example: fnxFindLastRow("Pipe Data", "GG", "OD 1")
' Params  : p_sWorksheet name of worksheet (i.e. Pipe Data);
'           p_sColumn is the column to search (i.e. "GG"), p_sColumnHeader is the column header (i.e. "OD 1")
' Returns : long (1st blank row after column header)
' Created : 08/15/2011 example
' --------------------------------------------------
Public Function fnxFindLastRow(p_sWorksheet As String, p_sColumn As String, p_sColumnHeader As String) As Long
    On Error GoTo PROC_ERR
    Dim sHeader As String
    Dim lLastRow As Long
    sHeader = Worksheets(p_sWorksheet).Range(p_sColumn & "3").Value
    If sHeader <> p_sColumnHeader Then
        fnxFindLastRow = 0
        MsgBox "Error (Cannot Find Header Column" & p_sColumnHeader & ")", vbExclamation + vbOKOnly, _
                "Error in fnxFindLastRow"
    Else
        lLastRow = 4
        Do While Len(Trim(Worksheets(p_sWorksheet).Range(Trim(p_sColumn) & Trim(Str(lLastRow))).Value)) > 0
            lLastRow = lLastRow + 1
        Loop
    End If
    fnxFindLastRow = lLastRow - 1
PROC_EXIT:
    Exit Function
PROC_ERR:
    fnxFindLastRow = 0
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in fnxFindLastRow"
    Resume PROC_EXIT
End Function
' --------------------------------------------------
' Comments: subHighlightRow -
' Params  :
' Created : 08/14/2011 10:08 AM
' Modified:
' --------------------------------------------------
Public Sub subHighlightRow(p_sWorksheet As String, p_lRowNumber As Long, p_eHighlight As ROW_HIGHLIGHT)
    On Error GoTo PROC_ERR
    If p_lRowNumber > 3 Then
        Worksheets(p_sWorksheet).Range(Trim(Str(p_lRowNumber)) & ":" & Trim(Str(p_lRowNumber))).Select
        Select Case p_eHighlight
            Case ROW_HIGHLIGHT.Highlight_Off
                Selection.Font.Color = vbBlack
                Selection.Font.Bold = False
            Case ROW_HIGHLIGHT.Highlight_RED
                Selection.Font.Color = vbRed
                Selection.Font.Bold = True
            Case Else
                Selection.Font.Color = vbRed
                Selection.Font.Bold = True
        End Select
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Error (" & Err.Description & ")", vbExclamation + vbOKOnly, "Error in subHighlightRow"
    Resume PROC_EXIT
End Sub
Public Function a() As Boolean
    subHighlightRow "Pipe Data", 4, Highlight_Off
End Function

