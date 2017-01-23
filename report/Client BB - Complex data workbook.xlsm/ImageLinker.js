Attribute VB_Name( = "ImageLinker")
function AddImageLinks(){
    var found  Boolean
    var ImageArray()  
    var ImageArraySize  
    var ImageCounter  
    var InitialFilePath  
    var ImageColumns()  
    var ImageColList  
    var ImageDirectory  
    var iCol  
    var iRow  
    var lastRow  
    var r  Range
    var Response  
    var SuccessfulExecution  Boolean
    var tImage  
    var wb  Workbook
    
    On( Error GoTo Closeout)
    SuccessfulExecution = false

    ImageColList( = "BA,BC,BE,BG,BI,BK,CA,CB,CC,CD,CE,CM,CN,CO,CP,CQ")
    ImageColumns( = Split(ImageColList, ","))

    //for ( var wb in Workbooks) {
    //    Response = vbNo
    //    if( ! wb Is ThisWorkbook ) {
    //        wb.Activate
    //        Response = MsgBox("Is this the workbook to add links to?", vbYesNo, "Identify the Workbook")
    //        if( Response == vbYes ) {
    //             break
    //         }
    //     }
    //} wb

    //if( Response == vbNo ) {
    //     break}
    // }

     wb = ThisWorkbook

    //if( ThisWorkbook.Worksheets("ImageLinker").DefaultCustom.Value == true ) {
    //    InitialFilePath = ThisWorkbook.Worksheets("ImageLinker").Range("F6").Value
    //}else{
    //    if( InStr(1, wb.Path, "Admin Shared", vbBinaryCompare) > 0 ) {
    //        InitialFilePath = wb.Path
    //     }
    
    InitialFilePath( = wb.Path)
    
    // }
    
    //if( InitialFilePath == "" ) {
    //    InitialFilePath = "U\"
    // }

    if( Right(InitialFilePath, 1) != "\" ) {
        InitialFilePath = InitialFilePath + "\"
     }
    
    //Ask for directory to look in
    with( Application.FileDialog(msoFileDialogFolderPicker)) {
        .InitialFileName( = InitialFilePath)
        .Title( = "Select the Image Directory")
        .Show()
        On Error Resume }
        ImageDirectory( = .SelectedItems(1))
        On( Error GoTo Closeout)
     }

    if( ImageDirectory == "" ) {
         break}
     }

    Application.ScreenUpdating = false

    ReDim( ImageArray(1 To 2, 1 To 1))
    ImageArraySize( = ImageArraySize + 1)

    //cycle through image ranges
    with( wb.Worksheets("Pipe Data")) {
        for( iCol = LBound(ImageColumns) To UBound(ImageColumns)) {
            Application(.StatusBar = Format(iCol / UBound(ImageColumns), "0%"))
            lastRow = .Range(ImageColumns(iCol) + .Rows.Count).(xlUp).row
            if( lastRow > 2 ) {
                for ( var r in .Range(ImageColumns(iCol) + "3" + ImageColumns(iCol) + lastRow)) {
                    if( ! IsError(r.Value) ) {
                        if( ! r.Value == "" ) {
                            found = false
                            for( ImageCounter = LBound(ImageArray, 2) To UBound(ImageArray, 2)) {
                                if( ImageArray(1, ImageCounter) == r.Value ) {
                                    CreateHyperlink( r, ImageArray(2, ImageCounter))
                                    found = true
                                 }
                            } ImageCounter
                            if( ! found ) {
                                tImage( = GetImageName(r.Value))
                                if( isFile(ImageDirectory + "\" + tImage) ) {
                                    CreateHyperlink r, ImageDirectory + "\" + tImage
                                    ImageArraySize( = ImageArraySize + 1)
                                    ReDim( Preserve ImageArray(1 To 2, 1 To ImageArraySize))
                                    ImageArray((1, ImageArraySize) = r.Value)
                                    ImageArray(2, ImageArraySize) = ImageDirectory + "\" + tImage
                                }else{
                                    tImage = "B-" + tImage
                                    if( isFile(ImageDirectory + "\" + tImage) ) {
                                        CreateHyperlink r, ImageDirectory + "\" + tImage
                                        ImageArraySize( = ImageArraySize + 1)
                                        ReDim( Preserve ImageArray(1 To 2, 1 To ImageArraySize))
                                        ImageArray((1, ImageArraySize) = r.Value)
                                        ImageArray(2, ImageArraySize) = ImageDirectory + "\" + tImage
                                     }
                                 }
                             }
                         }
                     }
                } r
             }
        } iCol
        lastRow = .Range("H" + .Rows.Count).(xlUp).row
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).Font.Name = "Arial"
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).Font.Size = 8
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).HorizontalAlignment = xlCenter
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).VerticalAlignment = xlCenter
     }

    SuccessfulExecution = true
Closeout()
    Application.CutCopyMode = false
    Application(.StatusBar = "")
    Application.ScreenUpdating = true
    if( ! SuccessfulExecution ) {
        MsgBox( "An error occurred and not all links could be added.")
     }
 }
 function isImage(ImageName  )  Boolean{
    if( ImageName == "" ) {
        isImage = false
         function{
     }
    var Ext$
    Ext$( = Right$(UCase$(ImageName), 4))
    if( Ext$ == ".JPG" || Ext$ == "JPEG" ) {
        isImage = true
     }
    if( Ext$ == ".TIF" || Ext$ == "TIFF" ) {
        isImage = true
     }
    if( Ext$ == ".PDF" ) {
        isImage = true
     }
    if( Ext$ == ".BMP" ) {
        isImage = true
     }
 }
 function isFile(FilePath  )  Boolean{
    //check the existence of a file
    var FSO  Object
     FSO = New FileSystemObject
    isFile( = FSO.FileExists(FilePath))
     FSO = Nothing
 }
function CreateHyperlink(MyRange  Range, InputAddress  ){
    var c()  
    var i  
    var fColor  
    var fBold  Boolean
    var MultiColored  Boolean
    var rLen  

    with( MyRange) {
        
        fBold( = .Cells.Font.Bold)
        rLen( = Len(.Value))

        if( ! IsNumeric(.Font.Color) ) {
            MultiColored = true
            ReDim( c(1 To rLen))
            for( i = 1 To rLen) {
                c((i) = .Characters(i, 1).Font.Color)
            } i
        }else{
            fColor( = .Cells.Font.Color)
         }
                
        .Hyperlinks(.Add Anchor=MyRange, Address=InputAddress, _)
        subAddress(="", TextToDisplay=CStr(.Value))
        
        if( MultiColored ) {
            for( i = 1 To rLen) {
                .Characters((i, 1).Font.Color = c(i))
            } i
        }else{
            .Font(.Color = fColor)
         }

        if( fBold ) {
            .Font(.Bold = fBold)
         }
     }
 }
function RemoveHyperlink(MyRange  Range){
    var c()  
    var fColor  
    var fBold  Boolean
    var MultiColored  Boolean
    var rLen  
    var i  

    with( MyRange) {
        rLen( = Len(.Value))
        if( ! IsNumeric(.Font.Color) ) {
            MultiColored = true
            ReDim( c(1 To rLen))
            for( i = 1 To rLen) {
                c((i) = .Characters(i, 1).Font.Color)
            } i
        }else{
            fColor( = .Cells.Font.Color)
         }
        fBold( = .Cells.Font.Bold)
        
        .Hyperlinks(.Delete)

        if( MultiColored ) {
            for( i = 1 To rLen) {
                .Characters((i, 1).Font.Color = c(i))
            } i
        }else{
            .Font(.Color = fColor)
         }
    
        if( fBold ) {
            .Font(.Bold = fBold)
         }
     }
 }
 function GetImageName(ImageName  )  {
    var Ext$
    Ext$( = Right$(UCase$(ImageName), 4))
    if( Ext$ == ".TIF" || Ext$ == ".JPG" ) {
        GetImageName = Left(ImageName, Len(ImageName) - 4) + ".PDF"
     }
    if( Ext$ == "TIFF" || Ext$ == "JPEG" ) {
        GetImageName = Left(ImageName, Len(ImageName) - 5) + ".PDF"
     }
    if( Ext$ == ".PDF" ) {
        GetImageName( = ImageName)
     }
    if( Ext$ == ".BMP" ) {
        GetImageName = Left(ImageName, Len(ImageName) - 4) + ".PDF"
     }
    if( ! InStr(1, ImageName, ".", vbBinaryCompare) > 0 ) {
        GetImageName = ImageName + ".PDF"
     }
 }
function RemoveImageLinks(){
    var ImageColumns()  
    var ImageColList  
    var iCol  
    var iRow  
    var lastRow  
    var r  Range
    var Response  
    var wb  Workbook

    ImageColList( = "BA,BC,BE,BG,BI,BK,CA,CB,CC,CD,CE,CM,CN,CO,CP,CQ")
    ImageColumns( = Split(ImageColList, ","))

     wb = ThisWorkbook
    
    //for ( var wb in Workbooks) {
    //    Response = vbNo
    //    if( ! wb Is ThisWorkbook ) {
    //        wb.Activate
    //        Response = MsgBox("Is this the workbook to remove links from?", vbYesNo, "Identify the Workbook")
    //        if( Response == vbYes ) {
    //             break
    //         }
    //     }
    //} wb

    //if( Response == vbNo ) {
    //     break}
    // }

    Application.ScreenUpdating = false
    //cycle through image ranges
    with( wb.Worksheets("Pipe Data")) {
        for( iCol = LBound(ImageColumns) To UBound(ImageColumns)) {
            Application(.StatusBar = Format(iCol / UBound(ImageColumns), "0%"))
            lastRow = .Range(ImageColumns(iCol) + .Rows.Count).(xlUp).row
            if( lastRow > 2 ) {
                for ( var r in .Range(ImageColumns(iCol) + "3" + ImageColumns(iCol) + lastRow)) {
                    if( r.Hyperlinks.Count > 0 ) {
                        RemoveHyperlink( r)
                     }
                } r
             }
        } iCol
        lastRow = .Range("H" + .Rows.Count).(xlUp).row
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).Font.Name = "Arial"
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).Font.Size = 8
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).HorizontalAlignment = xlCenter
        .Range(ImageColumns(LBound(ImageColumns)) + "3" + ImageColumns(UBound(ImageColumns)) + lastRow).VerticalAlignment = xlCenter
     }

    Application(.StatusBar = "")
    Application.ScreenUpdating = true
 }
function testcolors(){
    var i  
    var found  Boolean
    var j  
    for( j = 1 To 100000) {
    if( ! IsNumeric(Selection.Font.Color) ) {
        var c()  
        ReDim( c(1 To Len(Selection.Value)))
        for( i = 1 To Len(Selection.Value)) {
            if( IsNumeric(Selection.Characters(i, 1).Font.Color) ) {
                c((i) = Selection.Characters(i, 1).Font.Color)
             }
        } i
        Selection(.Font.Color = 0)
        for( i = 1 To Len(Selection.Value)) {
            if( IsNumeric(c(i)) ) {
                Selection(.Characters(i, 1).Font.Color = c(i))
             }
        } i
    
     }
    } j
 }


