Attribute VB_Name( = "Sheet10")
Attribute VB_Base( = "0{00020820-0000-0000-C000-000000000046}")
Attribute VB_GlobalNameSpace = false
Attribute VB_Creatable = false
Attribute VB_PredeclaredId = true
Attribute VB_Exposed = true
Attribute VB_TemplateDerived = false
Attribute VB_Customizable = true
Attribute VB_Control( = "validationSwitch, 1, 0, MSForms, CommandButton")

 function validationSwitch_Click(){
setValidationStatus (! validationStatus)
with( Me.validationSwitch) {
    Select( Case .Caption)
        Case( "On")
            .Caption( = "Off")
            deleteValidation()
             break}
        Case( "Off")
            .Caption( = "On")
            initializeValidation()
             break}
     Select
 }

 }

function addProperty(){

var sht  Worksheet

 sht = Sheets("FVE Validation")
sht.CustomProperties.Add "validationStatus", true

 }

function initializeSwitch(){

Me.validationSwitch(.Caption = "On")


 }

function checkprop(){

var cprop  CustomProperty
for ( var cprop in Sheets("FVE Validation").CustomProperties) {
    Debug.Print cprop.Name + " " + cprop.Value
}

 }

