Attribute VB_Name = "ImageLinker"
Sub AddImageLinks()
    Dim found As Boolean
    Dim ImageArray() As String
    Dim ImageArraySize As Long
    Dim ImageCounter As Long
    Dim InitialFilePath As String
    Dim ImageColumns() As String
    Dim ImageColList As String
    Dim ImageDirectory As String
    Dim iCol As Long
    Dim iRow As Long
    Dim lastRow As Long
    Dim r As Range
    Dim Response As String
    Dim SuccessfulExecution As Boolean
    Dim tImage As String
    Dim wb As Workbook
    
    On Error GoTo Closeout
    SuccessfulExecution = False

    ImageColList = "BA,BC,BE,BG,BI,BK,CA,CB,CC,CD,CE,CM,CN,CO,CP,CQ"
    ImageColumns = Split(ImageColList, ",")

    'For Each wb In Workbooks
    '    Response = vbNo
    '    If Not wb Is ThisWorkbook Then
    '        wb.Activate
    '        Response = MsgBox("Is this the workbook to add links to?", vbYesNo, "Identify the Workbook")
    '        If Response = vbYes Then
    '            Exit For
    '        End If
    '    End If
    'Next wb

    'If Response = vbNo Then
    '    Exit Sub
    'End If

    Set wb = ThisWorkbook

    'If ThisWorkbook.Worksheets("ImageLinker").DefaultCustom.Value = True Then
    '    InitialFilePath = ThisWorkbook.Worksheets("ImageLinker").Range("F6").Value
    'Else
    '    If InStr(1, wb.Path, "Admin Shared", vbBinaryCompare) > 0 Then
    '        InitialFilePath = wb.Path
    '    End If
    
    InitialFilePath = wb.Path
    
    'End If
    
    'If InitialFilePath = "" Then
    '    InitialFilePath = "U:\"
    'End If

    If Right(InitialFilePath, 1) <> "\" Then
        InitialFilePath = InitialFilePath & "\"
    End If
    
    'Ask for directory to look in
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = InitialFilePath
        .Title = "Select the Image Directory"
        .Show
        On Error Resume Next
        ImageDirectory = .SelectedItems(1)
        On Error GoTo Closeout
    End With

    If ImageDirectory = "" Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ReDim ImageArray(1 To 2, 1 To 1)
    ImageArraySize = ImageArraySize + 1

    'cycle through image ranges
    With wb.Worksheets("Pipe Data")
        For iCol = LBound(ImageColumns) To UBound(ImageColumns)
            Application.StatusBar = Format(iCol / UBound(ImageColumns), "0%")
            lastRow = .Range(ImageColumns(iCol) & .Rows.Count).End(xlUp).row
            If lastRow > 2 Then
                For Each r In .Range(ImageColumns(iCol) & "3:" & ImageColumns(iCol) & lastRow)
                    If Not IsError(r.Value) Then
                        If Not r.Value = "" Then
                            found = False
                            For ImageCounter = LBound(ImageArray, 2) To UBound(ImageArray, 2)
                                If ImageArray(1, ImageCounter) = r.Value Then
                                    CreateHyperlink r, ImageArray(2, ImageCounter)
                                    found = True
                                End If
                            Next ImageCounter
                            If Not found Then
                                tImage = GetImageName(r.Value)
                                If isFile(ImageDirectory & "\" & tImage) Then
                                    CreateHyperlink r, ImageDirectory & "\" & tImage
                                    ImageArraySize = ImageArraySize + 1
                                    ReDim Preserve ImageArray(1 To 2, 1 To ImageArraySize)
                                    ImageArray(1, ImageArraySize) = r.Value
                                    ImageArray(2, ImageArraySize) = ImageDirectory & "\" & tImage
                                Else
                                    tImage = "B-" & tImage
                                    If isFile(ImageDirectory & "\" & tImage) Then
                                        CreateHyperlink r, ImageDirectory & "\" & tImage
                                        ImageArraySize = ImageArraySize + 1
                                        ReDim Preserve ImageArray(1 To 2, 1 To ImageArraySize)
                                        ImageArray(1, ImageArraySize) = r.Value
                                        ImageArray(2, ImageArraySize) = ImageDirectory & "\" & tImage
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next r
            End If
        Next iCol
        lastRow = .Range("H" & .Rows.Count).End(xlUp).row
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).Font.Name = "Arial"
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).Font.Size = 8
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).HorizontalAlignment = xlCenter
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).VerticalAlignment = xlCenter
    End With

    SuccessfulExecution = True
Closeout:
    Application.CutCopyMode = False
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    If Not SuccessfulExecution Then
        MsgBox "An error occurred and not all links could be added."
    End If
End Sub
Private Function isImage(ImageName As String) As Boolean
    If ImageName = "" Then
        isImage = False
        Exit Function
    End If
    Dim Ext$
    Ext$ = Right$(UCase$(ImageName), 4)
    If Ext$ = ".JPG" Or Ext$ = "JPEG" Then
        isImage = True
    End If
    If Ext$ = ".TIF" Or Ext$ = "TIFF" Then
        isImage = True
    End If
    If Ext$ = ".PDF" Then
        isImage = True
    End If
    If Ext$ = ".BMP" Then
        isImage = True
    End If
End Function
Private Function isFile(FilePath As String) As Boolean
    'check the existence of a file
    Dim FSO As Object
    Set FSO = New FileSystemObject
    isFile = FSO.FileExists(FilePath)
    Set FSO = Nothing
End Function
Sub CreateHyperlink(MyRange As Range, InputAddress As String)
    Dim c() As Long
    Dim i As Long
    Dim fColor As Long
    Dim fBold As Boolean
    Dim MultiColored As Boolean
    Dim rLen As Long

    With MyRange
        
        fBold = .Cells.Font.Bold
        rLen = Len(.Value)

        If Not IsNumeric(.Font.Color) Then
            MultiColored = True
            ReDim c(1 To rLen)
            For i = 1 To rLen
                c(i) = .Characters(i, 1).Font.Color
            Next i
        Else
            fColor = .Cells.Font.Color
        End If
                
        .Hyperlinks.Add Anchor:=MyRange, Address:=InputAddress, _
        subAddress:="", TextToDisplay:=CStr(.Value)
        
        If MultiColored Then
            For i = 1 To rLen
                .Characters(i, 1).Font.Color = c(i)
            Next i
        Else
            .Font.Color = fColor
        End If

        If fBold Then
            .Font.Bold = fBold
        End If
    End With
End Sub
Sub RemoveHyperlink(MyRange As Range)
    Dim c() As Long
    Dim fColor As Long
    Dim fBold As Boolean
    Dim MultiColored As Boolean
    Dim rLen As Long
    Dim i As Long

    With MyRange
        rLen = Len(.Value)
        If Not IsNumeric(.Font.Color) Then
            MultiColored = True
            ReDim c(1 To rLen)
            For i = 1 To rLen
                c(i) = .Characters(i, 1).Font.Color
            Next i
        Else
            fColor = .Cells.Font.Color
        End If
        fBold = .Cells.Font.Bold
        
        .Hyperlinks.Delete

        If MultiColored Then
            For i = 1 To rLen
                .Characters(i, 1).Font.Color = c(i)
            Next i
        Else
            .Font.Color = fColor
        End If
    
        If fBold Then
            .Font.Bold = fBold
        End If
    End With
End Sub
Public Function GetImageName(ImageName As String) As String
    Dim Ext$
    Ext$ = Right$(UCase$(ImageName), 4)
    If Ext$ = ".TIF" Or Ext$ = ".JPG" Then
        GetImageName = Left(ImageName, Len(ImageName) - 4) & ".PDF"
    End If
    If Ext$ = "TIFF" Or Ext$ = "JPEG" Then
        GetImageName = Left(ImageName, Len(ImageName) - 5) & ".PDF"
    End If
    If Ext$ = ".PDF" Then
        GetImageName = ImageName
    End If
    If Ext$ = ".BMP" Then
        GetImageName = Left(ImageName, Len(ImageName) - 4) & ".PDF"
    End If
    If Not InStr(1, ImageName, ".", vbBinaryCompare) > 0 Then
        GetImageName = ImageName & ".PDF"
    End If
End Function
Sub RemoveImageLinks()
    Dim ImageColumns() As String
    Dim ImageColList As String
    Dim iCol As Long
    Dim iRow As Long
    Dim lastRow As Long
    Dim r As Range
    Dim Response As String
    Dim wb As Workbook

    ImageColList = "BA,BC,BE,BG,BI,BK,CA,CB,CC,CD,CE,CM,CN,CO,CP,CQ"
    ImageColumns = Split(ImageColList, ",")

    Set wb = ThisWorkbook
    
    'For Each wb In Workbooks
    '    Response = vbNo
    '    If Not wb Is ThisWorkbook Then
    '        wb.Activate
    '        Response = MsgBox("Is this the workbook to remove links from?", vbYesNo, "Identify the Workbook")
    '        If Response = vbYes Then
    '            Exit For
    '        End If
    '    End If
    'Next wb

    'If Response = vbNo Then
    '    Exit Sub
    'End If

    Application.ScreenUpdating = False
    'cycle through image ranges
    With wb.Worksheets("Pipe Data")
        For iCol = LBound(ImageColumns) To UBound(ImageColumns)
            Application.StatusBar = Format(iCol / UBound(ImageColumns), "0%")
            lastRow = .Range(ImageColumns(iCol) & .Rows.Count).End(xlUp).row
            If lastRow > 2 Then
                For Each r In .Range(ImageColumns(iCol) & "3:" & ImageColumns(iCol) & lastRow)
                    If r.Hyperlinks.Count > 0 Then
                        RemoveHyperlink r
                    End If
                Next r
            End If
        Next iCol
        lastRow = .Range("H" & .Rows.Count).End(xlUp).row
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).Font.Name = "Arial"
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).Font.Size = 8
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).HorizontalAlignment = xlCenter
        .Range(ImageColumns(LBound(ImageColumns)) & "3:" & ImageColumns(UBound(ImageColumns)) & lastRow).VerticalAlignment = xlCenter
    End With

    Application.StatusBar = ""
    Application.ScreenUpdating = True
End Sub
Sub testcolors()
    Dim i As Long
    Dim found As Boolean
    Dim j As Long
    For j = 1 To 100000
    If Not IsNumeric(Selection.Font.Color) Then
        Dim c() As Long
        ReDim c(1 To Len(Selection.Value))
        For i = 1 To Len(Selection.Value)
            If IsNumeric(Selection.Characters(i, 1).Font.Color) Then
                c(i) = Selection.Characters(i, 1).Font.Color
            End If
        Next i
        Selection.Font.Color = 0
        For i = 1 To Len(Selection.Value)
            If IsNumeric(c(i)) Then
                Selection.Characters(i, 1).Font.Color = c(i)
            End If
        Next i
    
    End If
    Next j
End Sub


