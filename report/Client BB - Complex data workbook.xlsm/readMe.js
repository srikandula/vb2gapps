Attribute VB_Name( = "readMe")


//
//Version 24
//-Updated data validation to check Unnamed Shorts name scheme
//-Updated data validation to check the X-Ray column
//-Updated image checker to run faster and be more robust
//-Locked VBA with password
//
//Version 23
//Updated rule for Class Location column
//
//Version 22
//Integrated data validation macro
//   -rules were changed to account for version 22 changes
//   -rules were added
//   -rules were removed
//   -does not substitute the image check macro for internal use at example
//Implemented image checking macro
//   -database calls were removed
//   -comparison and percentage integrated into code
//Naming macro integrated
 function ClearUsedRange(){
    var Sh  Worksheet
    for ( var Sh in ThisWorkbook.Worksheets) {
        Sh(.Activate)
        ActiveSheet(.UsedRange)
    } Sh
    ThisWorkbook(.Worksheets("Pipe Data").Activate)
 }
 function RemoveTrailingSpaces(){
    var Sh  Worksheet
    var c  
    for ( var Sh in ThisWorkbook.Worksheets) {
        with( Sh) {
            .Activate()
            ActiveSheet(.UsedRange)
            for( i = 1 To .UsedRange.Rows.Count) {
                for( j = 1 To .UsedRange.Columns.Count) {
                    if( Trim(CStr(.Cells(i, j))) != CStr(.Cells(i, j)) ) {
                        c( = c + 1)
                        .Cells((i, j).Value = Trim(CStr(.Cells(i, j).Value)))
                        Debug(.Print Sh.Name, Cells(i, j).Address)
                     }
                } j
            } i
         }
    } Sh
    ThisWorkbook(.Worksheets("Pipe Data").Activate)
    MsgBox c + " cells were fixed."
 }
 function DataValidationCleanUp(){
    Application.ScreenUpdating = false
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        var r  Range
        for ( var r in .Range("A3CW3")) {
            Application.StatusBar = "Column " + r.Column
            r(.Copy)
            .Range((.Cells(4, r.Column), .Cells(.Rows.Count, r.Column)).PasteSpecial Paste=xlPasteValidation, Operation=xlNone, _)
        SkipBlanks=false, Transpose=false
        } r
        Application(.StatusBar = "")
     }
 }
