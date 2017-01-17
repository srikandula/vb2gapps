Attribute VB_Name = "Sheet10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "validationSwitch, 1, 0, MSForms, CommandButton"

Private Sub validationSwitch_Click()
setValidationStatus (Not validationStatus)
With Me.validationSwitch
    Select Case .Caption
        Case "On"
            .Caption = "Off"
            deleteValidation
            Exit Sub
        Case "Off"
            .Caption = "On"
            initializeValidation
            Exit Sub
    End Select
End With

End Sub

Sub addProperty()

Dim sht As Worksheet

Set sht = Sheets("FVE Validation")
sht.CustomProperties.Add "validationStatus", True

End Sub

Sub initializeSwitch()

Me.validationSwitch.Caption = "On"


End Sub

Sub checkprop()

Dim cprop As CustomProperty
For Each cprop In Sheets("FVE Validation").CustomProperties
    Debug.Print cprop.Name & " " & cprop.Value
Next

End Sub

