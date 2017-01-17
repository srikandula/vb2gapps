Attribute VB_Name = "readMe"


'
'Version 24
'-Updated data validation to check Unnamed Shorts name scheme
'-Updated data validation to check the X-Ray column
'-Updated image checker to run faster and be more robust
'-Locked VBA with password
'
'Version 23
'Updated rule for Class Location column
'
'Version 22
'Integrated data validation macro
'   -rules were changed to account for version 22 changes
'   -rules were added
'   -rules were removed
'   -does not substitute the image check macro for internal use at example
'Implemented image checking macro
'   -database calls were removed
'   -comparison and percentage integrated into code
'Naming macro integrated
Private Sub ClearUsedRange()
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Worksheets
        Sh.Activate
        ActiveSheet.UsedRange
    Next Sh
    ThisWorkbook.Worksheets("Pipe Data").Activate
End Sub
Private Sub RemoveTrailingSpaces()
    Dim Sh As Worksheet
    Dim c As Long
    For Each Sh In ThisWorkbook.Worksheets
        With Sh
            .Activate
            ActiveSheet.UsedRange
            For i = 1 To .UsedRange.Rows.Count
                For j = 1 To .UsedRange.Columns.Count
                    If Trim(CStr(.Cells(i, j))) <> CStr(.Cells(i, j)) Then
                        c = c + 1
                        .Cells(i, j).Value = Trim(CStr(.Cells(i, j).Value))
                        Debug.Print Sh.Name, Cells(i, j).Address
                    End If
                Next j
            Next i
        End With
    Next Sh
    ThisWorkbook.Worksheets("Pipe Data").Activate
    MsgBox c & " cells were fixed."
End Sub
Private Sub DataValidationCleanUp()
    Application.ScreenUpdating = False
    With ThisWorkbook.Worksheets("Pipe Data")
        Dim r As Range
        For Each r In .Range("A3:CW3")
            Application.StatusBar = "Column: " & r.Column
            r.Copy
            .Range(.Cells(4, r.Column), .Cells(.Rows.Count, r.Column)).PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Next r
        Application.StatusBar = ""
    End With
End Sub
