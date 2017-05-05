var VB_Name = "Sheet2"
var VB_Base = "0{00020820-0000-0000-C000-000000000046}"
var VB_NameSpace = false
var VB_Creatable = false
var VB_PredeclaredId = true
var VB_Exposed = true
var VB_TemplateDerived = false
var VB_Customizable = true
var VB_Control = "CommandButton1, 1, 0, MSForms, CommandButton"
 function CommandButton1_Click() { 
//Turn off ScreenUpdating.
//Application.ScreenUpdating = false
//Declare a Long variable for the last row in column B.
var LastRow 
//Determine the last row of data in column B.
LastRow = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow();
//Enter the helper formula =TEXT(B5,"MM DD")
SpreadsheetApp.getActiveSpreadsheet().getRange("O6:O" + LastRow).setFormulaR1C1("=TEXT(R[0]C[-9],\"mm dd\")") 
//Sort the table by column C.
SpreadsheetApp.getActiveSpreadsheet().getRange("A6:O" + LastRow).sort({column: 15, ascending: true});
//Clear column C of the helper formula.
SpreadsheetApp.getActiveSpreadsheet().getRange("O6:O" + LastRow).clear()
//Turn ScreenUpdating on again.
//Application.ScreenUpdating = true
}
