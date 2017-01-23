Attribute VB_Name( = "SelectAMFPeriodForm")
Attribute VB_Base( = "0{AC8B4E76-DED6-4DCE-BB5A-E1239B7AE4A2}{227F21EF-0DB8-47BC-87CA-1D76D33E9C77}")
Attribute VB_GlobalNameSpace = false
Attribute VB_Creatable = false
Attribute VB_PredeclaredId = true
Attribute VB_Exposed = false
Attribute VB_TemplateDerived = false
Attribute VB_Customizable = false
 function UserForm_Initialize(){
    Load( SelectAMFPeriodForm)
    SelectAMFPeriodForm.PreCloseAMFButton = false
    SelectAMFPeriodForm.CurrentMonthAMFButton = false
 }

 function OKButton_Click(){
//    • Make sure that a radio button is selected
//    • if( pre-close
//        ? if( overwriting, delete pre-close data
//        ? if( not overwriting, call import pre-close
//    • if( not pre-close
//        ? if( overwriting, delete current data
//        ? if( not overwriting, call import current
    
    // if( nothing selected then prompt
    // if( pre-close then check for overwrite. if( overwrite then
    
    Call( InitializeTemplateGlobals)
    
    if( ! (SelectAMFPeriodForm.PreCloseAMFButton.Value || SelectAMFPeriodForm.CurrentMonthAMFButton) ) {
        MsgBox prompt="Please select a period for the All Mapping File or hit //Cancel//"
    else if( SelectAMFPeriodForm.PreCloseAMFButton ) { //if pre-close
        SelectAMFPeriodForm(.Hide)
        PreCloseAMFButton = false
        if( ! (IsEmpty(ActiveWorkbook.Names("admin_Pre_Close_AMF_Tab").RefersToRange)) ) {
            if( MsgBox("Importing another pre-close All Mapping File will completely overwrite (DELETE) the current data.  Are you sure you want to do this?", vbYesNo) == vbNo ) {
                MsgBox( ("Pre-close All Mapping File import CANCELLED!"))
                 break}
            }else{
                Call ResetPreCloseData //if overwriting then reset pre-close data
             }
         }
        Call ImportExternalAMF(bPreCloseAMF=true) //import pre-close AMF
    }else{ //if not pre-close
        SelectAMFPeriodForm(.Hide)
        CurrentMonthAMFButton = false
        if( ! (IsEmpty(ActiveWorkbook.Names("admin_Current_Month_AMF_Tab").RefersToRange)) ) {
            if( MsgBox("Importing another current month All Mapping File will completely overwrite (DELETE) the current data.  Are you sure you want to do this?", vbYesNo) == vbNo ) {
                MsgBox( ("Current month All Mapping File import CANCELLED!"))
                 break}
            }else{
                Call ResetCurrentMonthData //if overwriting then reset current month data
             }
         }
        Call ImportExternalAMF(bPreCloseAMF=false) //import current month AMF
     }
 }

 function CancelButton_Click(){
    SelectAMFPeriodForm(.Hide)
 }
