Attribute VB_Name( = "ImageChecker")
Option Explicit()
Global var firstRow21 = 3
Global var imageColStr22 = "53,55,57,59,61,63,79,80,81,82,83,91,92,93,94,95"
Global var reportFirstRow = 20
Global var PFLimageColumn = 3
Global var fileImageColumn = 9
Global var PFLimageAddress = "C20"
Global var fileImageAddress = "I20"
Global var featureColumn21 = 7
Global var resultAddress = "B16"
 function ClearFile_Click(){
    Application.ScreenUpdating = false
    with( ThisWorkbook.Worksheets("Tools")) {
        .Range("C20I" + .Rows.Count).ClearContents
        .Range("C20I" + .Rows.Count).NumberFormat = "@"
        .Range((resultAddress).ClearContents)
     }
 }
 function ImageCheck_Click(){
    Application.ScreenUpdating = false

    //Clear the report
    ClearFile_Click()

    //Add image references from the Pipe Data tab
    PopulateImages()

    //Add image file names from user selected directory
    PopulateFilenames()

    //Compare the lists
    var fr  , lr  , lr2  , MatchCount  , MissingCount  
    var r  Range, s  Range, iRng  Range, fRng  Range
    var ImageMatch  Boolean

    fr( = reportFirstRow)
    with( ThisWorkbook.Worksheets("Tools")) {
        lr = .Cells(Rows.Count, PFLimageColumn).(xlUp).row
        lr2 = .Cells(Rows.Count, fileImageColumn).(xlUp).row
         iRng = .Range("C" + fr + "C" + lr)
         fRng = .Range("I" + fr + "I" + lr2)
     }

    for ( var r in iRng) {
        ImageMatch = false
        for ( var s in fRng) {
            if( StrComp(TrimImgStr$(UCase$(r.Cells.Value)), TrimImgStr$(UCase$(s.Cells.Value)), vbBinaryCompare) == 0 ) {
                ImageMatch = true
             }
        }
        if( ImageMatch And r.row >== reportFirstRow And r.Value != "" ) {
            r(.Offset(0, 3).Value = "Present")
            MatchCount( = MatchCount + 1)
        }else{
            if( r.row >== reportFirstRow And r.Value != "" ) {
                r(.Offset(0, 3).Value = "Missing")
                MissingCount( = MissingCount + 1)
             }
         }
    }

    if( (MatchCount + MissingCount) != 0 ) {
        ThisWorkbook(.Worksheets("Tools").Range(resultAddress).Value = (MatchCount / (MatchCount + MissingCount)))
     }
    Application.ScreenUpdating = true
 }
 function PopulateImages(){
    Application.ScreenUpdating = false
    var FeatureColumn  
    var FirstRow  
    var iCol  
    var ImageColumns()  
    var ImageName  
    var iRow  
    var iRowImage  
    var lastRow  
    var r  Range
    var ReportColumn  
    var ReportSht  Worksheet
    var sht  Worksheet

     sht = ThisWorkbook.Worksheets("Pipe Data")
     ReportSht = ThisWorkbook.Sheets("Tools")
    ImageColumns( = Split(imageColStr22, ","))
    FirstRow( = firstRow21)
    FeatureColumn( = featureColumn21)
    lastRow = sht.Cells(Rows.Count, FeatureColumn).(xlUp).row
    iRowImage( = reportFirstRow)
    ReportColumn( = PFLimageColumn)

    //Pull all image names into the report //PFL Image Name// column
    for( iCol = LBound(ImageColumns) To UBound(ImageColumns)) {
        for( iRow = FirstRow To (lastRow + 1)) {
            if( Trim(sht.Cells(iRow, CLng(ImageColumns(iCol))).Value) != "" ) {
                with( ReportSht.Cells(iRowImage, ReportColumn)) {
                    .NumberFormat( = "@")
                    .Value( = sht.Cells(iRow, CLng(ImageColumns(iCol))).Value)
                 }
                iRowImage( = iRowImage + 1)
             }
        } iRow
    } iCol
    
    //Remove duplicates from Image Reference list
     r = ReportSht.Range(PFLimageAddress + "C" + iRowImage)
    RemoveRngDups( r)
    
     ReportSht = Nothing
     sht = Nothing
 }
 function RemoveRngDups(rng  Range){
    Application.ScreenUpdating = false
    if( rng.Count > 1 ) {
        rng(.removeDuplicates (1))
     }
    rng(.Font.Color = 1)
 }
 function PopulateFilenames(){
    Application.ScreenUpdating = false
    var i  
    var ReportSht  Worksheet
    var Cursor  
    var UserInputDirectory  

     ReportSht = ThisWorkbook.Worksheets("Tools")
    
    //Ask the user for the directory
    UserInputDirectory( = GetFolder)()
    if( UserInputDirectory == "" ) {
         break}
     }

    var FileArray()  Variant
    var FileCount  Integer
    var FileName  

    FileCount( = 0)
    On Error GoTo InvalidDir   //this doesn//t seem to be necessary
    FileName = Dir(UserInputDirectory + Application.PathSeparator + "*.*")
    if( FileName == "" ) {  break}

    //Loop through all files to build the array
    Do while FileName != "") {
        FileCount( = FileCount + 1)
        ReDim( Preserve FileArray(1 To FileCount))
        FileArray((FileCount) = FileName)
        FileName( = Dir)()
    Loop()

    with( ReportSht) {
        //Populate Directory File Name with array contents
        for( i = LBound(FileArray) To UBound(FileArray)) {
            Cursor( = i + reportFirstRow - 1)
            .Cells((Cursor, fileImageColumn).NumberFormat = "@")
            .Cells((Cursor, fileImageColumn).Value = FileArray(i))
        } i
        //Remove duplicates
        RemoveRngDups .Range("I" + reportFirstRow + "I" + Cursor + 1)
     }

InvalidDir()
     ReportSht = Nothing
 }
 function TrimImgStr(ByVal Str  )  {
    Str( = Replace$(Str, " ", ""))
    Str( = Replace$(Str, ".TIFF", ""))
    Str( = Replace$(Str, ".TIF", ""))
    Str( = Replace$(Str, ".JPG", ""))
    Str( = Replace$(Str, ".BMP", ""))
    Str( = Replace$(Str, ".PDF", ""))
    Str( = Replace$(Str, ".JPEG", ""))
    Str( = Replace$(Str, "B-", ""))
    Str( = Replace$(Str, "B_", ""))
    TrimImgStr( = Str)
 }
 function GetFolder()  {
    var Fldr  FileDialog
    var sItem  
     Fldr = Application.FileDialog(msoFileDialogFolderPicker)
    with( Fldr) {
        .Title( = "Select a Folder")
        .AllowMultiSelect = false
        .InitialFileName = ThisWorkbook.Path + Application.PathSeparator
        if( .Show != -1 ) { GoTo ExitCode
        sItem( = .SelectedItems(1))
     }
ExitCode()
    GetFolder( = sItem)
     Fldr = Nothing
 }
