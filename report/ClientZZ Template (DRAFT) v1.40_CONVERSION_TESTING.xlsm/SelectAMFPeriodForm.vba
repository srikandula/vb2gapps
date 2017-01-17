Attribute VB_Name = "SelectAMFPeriodForm"
Attribute VB_Base = "0{AC8B4E76-DED6-4DCE-BB5A-E1239B7AE4A2}{227F21EF-0DB8-47BC-87CA-1D76D33E9C77}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub UserForm_Initialize()
    Load SelectAMFPeriodForm
    SelectAMFPeriodForm.PreCloseAMFButton = False
    SelectAMFPeriodForm.CurrentMonthAMFButton = False
End Sub

Private Sub OKButton_Click()
'    • Make sure that a radio button is selected
'    • If pre-close:
'        ? If overwriting, delete pre-close data
'        ? If not overwriting, call import pre-close
'    • If not pre-close
'        ? If overwriting, delete current data
'        ? If not overwriting, call import current
    
    ' If nothing selected then prompt
    ' If pre-close then check for overwrite. If overwrite then
    
    Call InitializeTemplateGlobals
    
    If Not (SelectAMFPeriodForm.PreCloseAMFButton.Value Or SelectAMFPeriodForm.CurrentMonthAMFButton) Then
        MsgBox prompt:="Please select a period for the All Mapping File or hit 'Cancel'"
    ElseIf SelectAMFPeriodForm.PreCloseAMFButton Then 'if pre-close
        SelectAMFPeriodForm.Hide
        PreCloseAMFButton = False
        If Not (IsEmpty(ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange)) Then
            If MsgBox("Importing another pre-close All Mapping File will completely overwrite (DELETE) the current data.  Are you sure you want to do this?", vbYesNo) = vbNo Then
                MsgBox ("Pre-close All Mapping File import CANCELLED!")
                Exit Sub
            Else
                Call ResetPreCloseData 'if overwriting then reset pre-close data
            End If
        End If
        Call ImportExternalAMF(bPreCloseAMF:=True) 'import pre-close AMF
    Else 'if not pre-close
        SelectAMFPeriodForm.Hide
        CurrentMonthAMFButton = False
        If Not (IsEmpty(ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange)) Then
            If MsgBox("Importing another current month All Mapping File will completely overwrite (DELETE) the current data.  Are you sure you want to do this?", vbYesNo) = vbNo Then
                MsgBox ("Current month All Mapping File import CANCELLED!")
                Exit Sub
            Else
                Call ResetCurrentMonthData 'if overwriting then reset current month data
            End If
        End If
        Call ImportExternalAMF(bPreCloseAMF:=False) 'import current month AMF
    End If
End Sub

Private Sub CancelButton_Click()
    SelectAMFPeriodForm.Hide
End Sub
