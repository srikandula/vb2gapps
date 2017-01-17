Attribute VB_Name = "ImageChecker"
Option Explicit
Global Const firstRow21 = 3
Global Const imageColStr22 = "53,55,57,59,61,63,79,80,81,82,83,91,92,93,94,95"
Global Const reportFirstRow = 20
Global Const PFLimageColumn = 3
Global Const fileImageColumn = 9
Global Const PFLimageAddress = "C20"
Global Const fileImageAddress = "I20"
Global Const featureColumn21 = 7
Global Const resultAddress = "B16"
Public Sub ClearFile_Click()
    Application.ScreenUpdating = False
    With ThisWorkbook.Worksheets("Tools")
        .Range("C20:I" & .Rows.Count).ClearContents
        .Range("C20:I" & .Rows.Count).NumberFormat = "@"
        .Range(resultAddress).ClearContents
    End With
End Sub
Public Sub ImageCheck_Click()
    Application.ScreenUpdating = False

    'Clear the report
    ClearFile_Click

    'Add image references from the Pipe Data tab
    PopulateImages

    'Add image file names from user selected directory
    PopulateFilenames

    'Compare the lists
    Dim fr As Long, lr As Long, lr2 As Long, MatchCount As Long, MissingCount As Long
    Dim r As Range, s As Range, iRng As Range, fRng As Range
    Dim ImageMatch As Boolean

    fr = reportFirstRow
    With ThisWorkbook.Worksheets("Tools")
        lr = .Cells(Rows.Count, PFLimageColumn).End(xlUp).row
        lr2 = .Cells(Rows.Count, fileImageColumn).End(xlUp).row
        Set iRng = .Range("C" & fr & ":C" & lr)
        Set fRng = .Range("I" & fr & ":I" & lr2)
    End With

    For Each r In iRng
        ImageMatch = False
        For Each s In fRng
            If StrComp(TrimImgStr$(UCase$(r.Cells.Value)), TrimImgStr$(UCase$(s.Cells.Value)), vbBinaryCompare) = 0 Then
                ImageMatch = True
            End If
        Next
        If ImageMatch And r.row >= reportFirstRow And r.Value <> "" Then
            r.Offset(0, 3).Value = "Present"
            MatchCount = MatchCount + 1
        Else
            If r.row >= reportFirstRow And r.Value <> "" Then
                r.Offset(0, 3).Value = "Missing"
                MissingCount = MissingCount + 1
            End If
        End If
    Next

    If (MatchCount + MissingCount) <> 0 Then
        ThisWorkbook.Worksheets("Tools").Range(resultAddress).Value = (MatchCount / (MatchCount + MissingCount))
    End If
    Application.ScreenUpdating = True
End Sub
Private Sub PopulateImages()
    Application.ScreenUpdating = False
    Dim FeatureColumn As Long
    Dim FirstRow As Long
    Dim iCol As Long
    Dim ImageColumns() As String
    Dim ImageName As String
    Dim iRow As Long
    Dim iRowImage As Long
    Dim lastRow As Long
    Dim r As Range
    Dim ReportColumn As Long
    Dim ReportSht As Worksheet
    Dim sht As Worksheet

    Set sht = ThisWorkbook.Worksheets("Pipe Data")
    Set ReportSht = ThisWorkbook.Sheets("Tools")
    ImageColumns = Split(imageColStr22, ",")
    FirstRow = firstRow21
    FeatureColumn = featureColumn21
    lastRow = sht.Cells(Rows.Count, FeatureColumn).End(xlUp).row
    iRowImage = reportFirstRow
    ReportColumn = PFLimageColumn

    'Pull all image names into the report 'PFL Image Name' column
    For iCol = LBound(ImageColumns) To UBound(ImageColumns)
        For iRow = FirstRow To (lastRow + 1)
            If Trim(sht.Cells(iRow, CLng(ImageColumns(iCol))).Value) <> "" Then
                With ReportSht.Cells(iRowImage, ReportColumn)
                    .NumberFormat = "@"
                    .Value = sht.Cells(iRow, CLng(ImageColumns(iCol))).Value
                End With
                iRowImage = iRowImage + 1
            End If
        Next iRow
    Next iCol
    
    'Remove duplicates from Image Reference list
    Set r = ReportSht.Range(PFLimageAddress & ":C" & iRowImage)
    RemoveRngDups r
    
    Set ReportSht = Nothing
    Set sht = Nothing
End Sub
Private Sub RemoveRngDups(rng As Range)
    Application.ScreenUpdating = False
    If rng.Count > 1 Then
        rng.removeDuplicates (1)
    End If
    rng.Font.Color = 1
End Sub
Private Sub PopulateFilenames()
    Application.ScreenUpdating = False
    Dim i As Long
    Dim ReportSht As Worksheet
    Dim Cursor As Long
    Dim UserInputDirectory As String

    Set ReportSht = ThisWorkbook.Worksheets("Tools")
    
    'Ask the user for the directory
    UserInputDirectory = GetFolder()
    If UserInputDirectory = "" Then
        Exit Sub
    End If

    Dim FileArray() As Variant
    Dim FileCount As Integer
    Dim FileName As String

    FileCount = 0
    On Error GoTo InvalidDir   'this doesn't seem to be necessary
    FileName = Dir(UserInputDirectory & Application.PathSeparator & "*.*")
    If FileName = "" Then Exit Sub

    'Loop through all files to build the array
    Do While FileName <> ""
        FileCount = FileCount + 1
        ReDim Preserve FileArray(1 To FileCount)
        FileArray(FileCount) = FileName
        FileName = Dir()
    Loop

    With ReportSht
        'Populate Directory File Name with array contents
        For i = LBound(FileArray) To UBound(FileArray)
            Cursor = i + reportFirstRow - 1
            .Cells(Cursor, fileImageColumn).NumberFormat = "@"
            .Cells(Cursor, fileImageColumn).Value = FileArray(i)
        Next i
        'Remove duplicates
        RemoveRngDups .Range("I" & reportFirstRow & ":I" & Cursor + 1)
    End With

InvalidDir:
    Set ReportSht = Nothing
End Sub
Private Function TrimImgStr(ByVal Str As String) As String
    Str = Replace$(Str, " ", "")
    Str = Replace$(Str, ".TIFF", "")
    Str = Replace$(Str, ".TIF", "")
    Str = Replace$(Str, ".JPG", "")
    Str = Replace$(Str, ".BMP", "")
    Str = Replace$(Str, ".PDF", "")
    Str = Replace$(Str, ".JPEG", "")
    Str = Replace$(Str, "B-", "")
    Str = Replace$(Str, "B_", "")
    TrimImgStr = Str
End Function
Private Function GetFolder() As String
    Dim Fldr As FileDialog
    Dim sItem As String
    Set Fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With Fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        If .Show <> -1 Then GoTo ExitCode
        sItem = .SelectedItems(1)
    End With
ExitCode:
    GetFolder = sItem
    Set Fldr = Nothing
End Function
