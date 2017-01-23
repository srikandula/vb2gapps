Attribute VB_Name( = "DataValidation")
Option Explicit()
 Version
 var FirstRow = 3
 var ActSizeColumn = "W"
 var AngleColumn = "AD"
 var ANSIColumn = "V"
 var BarredColumn = "AG"
 var BegStaColumn = "A"
 var BLineColumn = "CG"
 var BuilderColumn = "CI"
 var CasTypeColumn = "AN"
 var CheckerColumn = "CJ"
 var ClassLocColumn = "D"
 var CoatDateColumn = "AR"
 var CoatTypeColumn = "AP"
 var DiscrepancyColumn = "CL"
 var DPlatColumn = "BN"
 var Dwg1Column = "BA"
 var Dwg1QColumn = "BB"
 var Dwg2Column = "BC"
 var Dwg2QColumn = "BD"
 var EndConnectColumn = "AB"
 var EndStaColumn = "B"
 var FabAssColumn = "AY"
 var FeatureColumn = "H"
 var FeatureNoColumn = "G"
 var FieldSTColumn = "M"
 var GISPipeSegIDColumn = "C"
 var Img10Column = "CM"
 var Img11Column = "CN"
 var Img12Column = "CO"
 var Img13Column = "CP"
 var Img14Column = "CQ"
 var Img1Column = "BE"
 var Img1QColumn = "BF"
 var Img2Column = "BG"
 var Img2QColumn = "BH"
 var Img3Column = "BI"
 var Img3QColumn = "BJ"
 var Img4Column = "BK"
 var Img4QColumn = "BL"
 var InsertionColumn = "AI"
 var InstallDateColumn = "F"
 var InsulatorColumn = "AO"
 var JobNoColumn = "E"
 var LengthColumn = "J"
 var MaterialColumn = "AC"
 var MaxPressColumn = "X"
 var MethodColumn = "AH"
 var MPColumn = "L"
 var OD1Column = "O"
 var OD2Column = "Q"
 var OperatorColumn = "AK"
 var OrientationColumn = "AF"
 var PercentSMYSColumn = "CW"
 var PipeSTColumn = "N"
 var PreFabColumn = "AX"
 var PurDateColumn = "AV"
 var PurOtherColumn = "AW"
 var RadiusColumn = "AE"
 var ReconSalvColumn = "AZ"
 var SealColumn = "AM"
 var SeamTypeColumn = "S"
 var ShellPressColumn = "AJ"
 var SMYSColumn = "U"
 var SpecsColumn = "T"
 var STPRAdjPressColumn = "BT"
 var STPRBChartColumn = "CC"
 var STPRDateColumn = "BU"
 var STPRDatePriorColumn = "BX"
 var STPRDeadLogColumn = "CD"
 var STPRDurColumn = "BS"
 var STPRFChartColumn = "CB"
 var STPRImgColumn = "CA"
 var STPRMAOPColumn = "BZ"
 var STPRMediaColumn = "BQ"
 var STPRPressColumn = "BR"
 var STPRPressPriorColumn = "BY"
 var STPRQColumn = "CF"
 var STPRSketchColumn = "CE"
 var STPRTypeColumn = "BO"
 var TypeColumn = "I"
 var VentedColumn = "AL"
 var WT1Column = "P"
 var WT2Column = "R"
 var XRayColumn = "CH"
 lastRow  
 function DataValidation_Click(){
    ValidateData()
 }
 function ValidClean_Click(){
    ValidationCleanup()
    ThisWorkbook(.Worksheets("Pipe Data").Activate)
 }
 function ValidateData(){
    var ForCounter  
    var ListName  Name
    var PipeDataSht  Worksheet
    var RowCtr  
    Application.ScreenUpdating = false
     PipeDataSht = ThisWorkbook.Worksheets("Pipe Data")

    ValidationCleanup()

    //Count all the rows
    lastRow = PipeDataSht.Range("H" + PipeDataSht.Rows.Count).(xlUp).row
    if( lastRow < FirstRow ) {  break}

    //Remove corrupt List Names
    for ( var ListName in ActiveWorkbook.Names) {
        if( UCase$(ListName.RefersTo) Like ("*.XL*") ) {
            ListName(.Delete)
         }
    } ListName

    with( PipeDataSht) {
        //Check standard PFL data
        if( IsEmpty(.Range("G1")) ) {
            .Range(("G1").Interior.ColorIndex = 4)
         }
        if( ! IsEmpty(.Range("J1")) ) {
            if( IsEmpty(.Range("L1")) ) {
                .Range(("L1").Interior.ColorIndex = 4)
             }
        }else{
            if( ! IsEmpty(.Range("L1")) ) {
                .Range(("J1").Interior.ColorIndex = 4)
             }
         }

        //Begin and  Mile Point
        if( IsEmpty(.Range(MPColumn + FirstRow)) ) { MarkWrong FirstRow, MPColumn
        if( IsEmpty(.Range(MPColumn + lastRow)) ) { MarkWrong lastRow, MPColumn

        //Iterate through each row
        for( ForCounter = FirstRow To lastRow) {
            //Check % SMYS
            if( IsNumeric(.Range(PercentSMYSColumn + ForCounter)) ) {
                if( .Range(PercentSMYSColumn + ForCounter).Value >== 1 ) {
                    .Range(FeatureNoColumn + ForCounter).Interior.Color = 65535
                 }
             }
            //Beg Station
            if( ! IsEmpty(.Range(BegStaColumn + ForCounter)) ) {
                if( ! IsNumeric(.Range(BegStaColumn + ForCounter)) ) {
                    MarkWrong( ForCounter, BegStaColumn)
                 }
            }else{
                MarkWrong( ForCounter, BegStaColumn)
             }
            // Station
            if( ! IsEmpty(.Range(EndStaColumn + ForCounter)) ) {
                if( ! IsNumeric(.Range(EndStaColumn + ForCounter)) ) {
                    MarkWrong( ForCounter, EndStaColumn)
                 }
            }else{
                MarkWrong( ForCounter, EndStaColumn)
             }
            //Pipe Segment ID
            if( .Range(GISPipeSegIDColumn + ForCounter) == Empty ) {
                MarkWrong( ForCounter, GISPipeSegIDColumn)
             }
            //Job Number
            if( IsEmpty(.Range(JobNoColumn + ForCounter)) _
            And .Range(FeatureColumn + ForCounter).Value != "Tap" ) {
                MarkWrong( ForCounter, JobNoColumn)
             }
            //Check Feature Number column
            if( IsNumeric(.Range(FeatureNoColumn + ForCounter).Value) ) {
                if( ForCounter > FirstRow ) {
                    if( .Range(FeatureNoColumn + ForCounter).Value _
                    <= .Range(FeatureNoColumn + ForCounter - 1).Value ) {
                        MarkWrong( ForCounter, FeatureNoColumn)
                     }
                 }
                if( ForCounter < lastRow ) {
                    if( .Range(FeatureNoColumn + ForCounter).Value _
                    >= .Range(FeatureNoColumn + ForCounter + 1).Value ) {
                        MarkWrong( ForCounter, FeatureNoColumn)
                     }
                 }
            }else{
                MarkWrong( ForCounter, FeatureNoColumn)
             }
            //Check Required Fields
            if( IsEmpty(.Range(FeatureColumn + ForCounter)) ) {
                MarkWrong( ForCounter, FeatureColumn)
             }
            if( IsEmpty(.Range(TypeColumn + ForCounter)) ) {
                MarkWrong( ForCounter, TypeColumn)
             }
            
            //Check length field
            //if( .Range(LengthColumn + ForCounter).Value _
            != .Range(EndStaColumn + ForCounter).Value _
            - .Range(BegStaColumn + ForCounter).Value ) {
            //    MarkWrong ForCounter, LengthColumn
            // }
            
            //Check Current MAOP field
            if( IsEmpty(.Range(STPRMAOPColumn + ForCounter)) ) {
                MarkWrong( ForCounter, STPRMAOPColumn)
            }else{
                if( ! IsNumeric(.Range(STPRMAOPColumn + ForCounter)) ) {
                    MarkWrong( ForCounter, STPRMAOPColumn)
                 }
             }
            
            //Check that seam type and install date line up
            if( UCase$(.Range(SeamTypeColumn + ForCounter)) == UCase$("Unknown > 4 - Modern") ) {
                if( ! IsEmpty(.Range(InstallDateColumn + ForCounter)) ) {
                    if( IsDate(.Range(InstallDateColumn + ForCounter).Value) ) {
                        if( Year(CDate(.Range(InstallDateColumn + ForCounter).Value)) < 1960 ) {
                            if( IsEmpty(.Range(OD1Column + ForCounter)) ) {
                                MarkWrong( ForCounter, SeamTypeColumn)
                            }else{
                                if( .Range(OD1Column + ForCounter).Value < 28 ) {
                                    MarkWrong( ForCounter, SeamTypeColumn)
                                 }
                             }
                         }
                     }
                 }
             }
            
        } ForCounter
     }

    //Testing Class Loc.
    ValidateWithCodeLists( ClassLocColumn, "ClassLocation")

    //Testing Install Date
    ValidateDate( InstallDateColumn, CDate("1/1/1910"), CDate("12/31/2015"))
    ValidateRequiredField( InstallDateColumn)

    //Testing Feature Type
    ValidateWithCodeLists( FeatureColumn, "EventType")

    //Testing Type, Type List varies based on feature type
    ValidateWithCodeLists( TypeColumn, "CheckType")

    //Testing Mile Point
    ValidateNumber( MPColumn, Min=0)

    //Testing Field Station
    ValidateNumber( FieldSTColumn, Min=0)

    //Testing Pipe Station
    ValidateNumber( PipeSTColumn, Min=0)
       
    //******* MAOP Specifications *******
    //Testing O.D.1
    ValidateRequiredField( OD1Column, FeatureVal="Cap, Regulator, Pipe, Tap, Field Bend, Mfg Bend, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateRequiredEmptyField( OD1Column, FeatureVal="Appurtenance")
    ValidateWithCodeLists( OD1Column, "CheckOD")
    //Testing W.T.1
    ValidateRequiredField( WT1Column, FeatureVal="Pipe, Field Bend, Mfg Bend, Tee, Reducer, Sleeve, Other")
    ValidateRequiredField( WT1Column, FeatureVal="Flange", TypeVal="Lap Joint, Threaded, Reducing, Socket Weld, Insulating")
    ValidateRequiredEmptyField( WT1Column, FeatureVal="PCF, Valve, Appurtenance, Meter, Pig Trap, Regulator, Relief Valve")
    ValidateRequiredEmptyField( WT1Column, FeatureVal="Flange", TypeVal="Weld Neck, Blind, Slip On")
    ValidateNumber( WT1Column, Min=0.01, Max=1.5)
    //Testing O.D.2
    ValidateRequiredField( OD2Column, FeatureVal="Tee, Reducer")
    ValidateRequiredField( OD2Column, FeatureVal="Mfg Bend, Flange", TypeVal="Reducing")
    ValidateRequiredField( OD2Column, FeatureVal="Pipe", TypeVal="Casing")
    ValidateRequiredEmptyField( OD2Column, FeatureVal="Field Bend, Flange, PCF, Tap, Valve, Sleeve, Appurtenance, Meter, Pig Trap, Regulator, Relief Valve, Cap, Other")
    ValidateRequiredEmptyField( OD2Column, FeatureVal="Mfg Bend", TypeVal="Forged, Miter, Bell Bell Chill Ring, Wrinkle, Rolled Plate, Socket Weld")
    ValidateRequiredEmptyField( OD2Column, FeatureVal="Pipe", TypeVal="No Casing,Pipe Bridge,Pipe Span,Pipe Liner, Pipe Encased")
    ValidateRequiredEmptyField( OD2Column, FeatureVal="Flange", TypeVal="Weld Neck, Blind, Slip On, Lap Joint, Threaded, Socket Weld, Insulating")
    ValidateWithCodeLists( OD2Column, "OutsideDiameter")
    //Testing W.T.2
    ValidateRequiredField( WT2Column, FeatureVal="Tee, Reducer")
    ValidateRequiredField( WT2Column, FeatureVal="Mfg Bend, Flange", TypeVal="Reducing")
    ValidateRequiredField( WT2Column, FeatureVal="Pipe", TypeVal="Casing")
    ValidateRequiredEmptyField( WT2Column, FeatureVal="Pipe", TypeVal="No Casing,Pipe Bridge,Pipe Span,Pipe Liner, Pipe Encased")
    ValidateRequiredEmptyField( WT2Column, FeatureVal="Field Bend, Flange, PCF, Tap, Valve, Sleeve, Appurtenance, Meter, Pig Trap, Regulator, Relief Valve, Cap, Other")
    ValidateRequiredEmptyField( WT2Column, FeatureVal="Pipe", TypeVal="No Casing,Pipe Bridge,Pipe Span,Pipe Liner, Pipe Encased")
    ValidateRequiredEmptyField( WT2Column, FeatureVal="Mfg Bend", TypeVal="Forged, Miter, Bell Bell Chill Ring, Wrinkle, Rolled Plate, Socket Weld")
    ValidateRequiredEmptyField( WT2Column, FeatureVal="Flange", TypeVal="Weld Neck, Blind, Slip On, Lap Joint, Threaded, Socket Weld, Insulating")
    ValidateNumber( WT2Column, 0.01, 1.5)
    //Testing Seam Type
    ValidateRequiredField( SeamTypeColumn, FeatureVal="Pipe, Field Bend, Mfg Bend, Tee, Reducer")
    ValidateRequiredEmptyField( SeamTypeColumn, FeatureVal="Flange, Appurtenance, Pig Trap")
    ValidateWithCodeLists( SeamTypeColumn, "SeamType")
    //Testing Specification / Rating
    ValidateRequiredField( SpecsColumn, FeatureVal="Cap, Pipe, Tap, Field Bend, Mfg Bend, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other")
    ValidateRequiredField( SpecsColumn, FeatureVal="Relief Valve")
    ValidateRequiredEmptyField( SpecsColumn, FeatureVal="Valve, Appurtenance, Regulator")
    ValidateWithCodeLists( SpecsColumn, "CheckSpecs")
    //Testing SMYS Column
    ValidateRequiredField( SMYSColumn, FeatureVal="Cap, Pipe, Tap, Field Bend, Mfg Bend, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other")
    ValidateRequiredEmptyField( SMYSColumn, FeatureVal="Valve, Appurtenance, Regulator")
    ValidateWithCodeLists( SMYSColumn, "SMYS")
    //Testing ANSI - WOG Rating
    ValidateRequiredField( ANSIColumn, FeatureVal="Regulator, Meter, Cap, Tap, Sleeve, PCF, Flange, Relief Valve")
    ValidateRequiredEmptyField( ANSIColumn, FeatureVal="Field Bend, Pipe, Appurtenance, Pig Trap")
    ValidateRequiredEmptyField( ANSIColumn, FeatureVal="Other", TypeVal="Collar")
    ValidateRequiredField( ANSIColumn, FeatureVal="Other", TypeVal="Closure, Insulating Joint, Filter/Strainer, Separator, Pig Signal, Pressure Transmitter, Dresser Mechanical, Internal Baffle, Threaded Weldolet, Weldolet or variations, Plug, Union, Repair Can, Other, Unknown")
    ValidateMutuallyExclusive( ANSIColumn, SMYSColumn, FeatureVal="Reducer,Tee,Mfg Bend,Flange")
    ValidateWithCodeLists( ANSIColumn, "Ansi")
    //Testing Actual Size or Opening
    ValidateRequiredField( ActSizeColumn, FeatureVal="Tap, Valve, Sleeve, Regulator, Relief Valve, Other")
    ValidateRequiredEmptyField( ActSizeColumn, FeatureVal="Cap, Pipe, Field Bend, Mfg Bend, Tee, Reducer, Appurtenance, PCF, Flange, Meter, Pig Trap")
    ValidateNumber( ActSizeColumn, 0.1, 125)
    //Testing Max. Working Pressure
    ValidateNumber( MaxPressColumn, 100, 5000)
    ValidateRequiredEmptyField( MaxPressColumn, FeatureVal="Field Bend, Pipe, Appurtenance, Meter, Pig Trap, Cap")
    ValidateRequiredField( MaxPressColumn, FeatureVal="Tap, Regulator, Sleeve, PCF, Flange, Relief Valve")
    ValidateRequiredField( MaxPressColumn, FeatureVal="Other", TypeVal="Closure, Insulating Joint, Filter/Strainer, Separator, Pig Signal, Pressure Transmitter, Dresser Mechanical, Internal Baffle, Threaded Weldolet, Weldolet or variations, Plug, Union, Repair Can, Other, Unknown")
    //******* Feature Specs *******
    //Testing  Connect
    ValidateRequiredField( EndConnectColumn, FeatureVal="Pipe, Valve, Flange, Regulator, Other")
    ValidateWithCodeLists( EndConnectColumn, "EndConnect")
    //Testing Material Type
    ValidateWithCodeLists( MaterialColumn, "MaterialType")
    ValidateRequiredField( MaterialColumn, FeatureVal="Pipe, Mfg Bend, Sleeve")
    //******* Bend Data *******
    //Testing Angle
    ValidateRequiredField( AngleColumn, FeatureVal="Field Bend, Mfg Bend")
    ValidateRequiredEmptyField( AngleColumn, FeatureVal="Appurtenance, Cap, Regulator, Pipe, Tap, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateNumber( AngleColumn, 0, 90)
    //Testing Radius
    ValidateRequiredField( RadiusColumn, FeatureVal="Field Bend, Mfg Bend", TypeVal="Forged, Bell Bell Chill Ring, Wrinkle, Rolled Plate, Reducing, Socket Weld, Unknown, Bending Machine, Int Mandrel-Bending Machine")
    ValidateRequiredEmptyField( RadiusColumn, FeatureVal="Appurtenance, Cap, Regulator, Pipe, Tap, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithCodeLists( RadiusColumn, "CheckRadius")
    //Testing Orient
    ValidateRequiredField( OrientationColumn, FeatureVal="Field Bend, Mfg Bend")
    ValidateRequiredEmptyField( OrientationColumn, FeatureVal="Appurtenance, Cap, Regulator, Pipe, Tap, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithCodeLists( OrientationColumn, "Orientation")
    
    //******* Tee Data *******
    //Testing Barred?
    ValidateRequiredField( BarredColumn, FeatureVal="Tee")
    ValidateRequiredEmptyField( BarredColumn, FeatureVal="Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tap, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithString( BarredColumn, "Yes, No, Unknown")
    
    //******* Tap Data *******
    //Testing Method
    ValidateRequiredField( MethodColumn, FeatureVal="Tap")
    ValidateRequiredEmptyField( MethodColumn, FeatureVal="Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithCodeLists( MethodColumn, "HotCold")
    //Testing Insertion
    ValidateRequiredField( InsertionColumn, FeatureVal="Tap")
    ValidateRequiredEmptyField( InsertionColumn, FeatureVal="Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithString( InsertionColumn, "Yes, No, Unknown")
    
    //******* Valves *******
    //Testing Shell Test Pressure
    ValidateRequiredField( ShellPressColumn, FeatureVal="Valve")
    ValidateRequiredEmptyField( ShellPressColumn, FeatureVal="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other")
    ValidateNumber( ShellPressColumn, 100, 5000)
    //Testing Operator Type
    ValidateRequiredField( OperatorColumn, FeatureVal="Valve")
    ValidateRequiredEmptyField( OperatorColumn, FeatureVal="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other")
    ValidateWithCodeLists( OperatorColumn, "ValveOperatorType")

    //******* Casing Data *******
    //Testing Vented?
    ValidateRequiredField( VentedColumn, FeatureVal="Pipe", TypeVal="Casing")
    ValidateRequiredEmptyField( VentedColumn, FeatureVal="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithCodeLists( VentedColumn, "YesNoUnknown")
    //Testing Seal Type
    ValidateRequiredField( SealColumn, FeatureVal="Pipe", TypeVal="Casing")
    ValidateRequiredEmptyField( SealColumn, FeatureVal="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithCodeLists( SealColumn, "CasingSealType")
    //Testing Type
    ValidateRequiredField( CasTypeColumn, FeatureVal="Pipe", TypeVal="Casing")
    ValidateRequiredEmptyField( CasTypeColumn, FeatureVal="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithCodeLists( CasTypeColumn, "CasingType")
    //Testing Insulator Type
    ValidateRequiredField( InsulatorColumn, FeatureVal="Pipe", TypeVal="Casing")
    ValidateRequiredEmptyField( InsulatorColumn, FeatureVal="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other")
    ValidateWithCodeLists( InsulatorColumn, "CasingInsulateType")

    //******* Coating *******
    //Testing Coating Type
    ValidateWithCodeLists( CoatTypeColumn, "CoatingType")
    //Testing Coating Install Date
    ValidateDate( CoatDateColumn, MinDate=CDate("1/1/1915"), MaxDate=CDate("12/31/2015"))

    //******* Purchase and Install Information *******
    //Testing Purchase Date of Feature
    ValidateDate( PurDateColumn, MinDate="1/1/1915", MaxDate="1/1/2020")
    //Testing Purchased from Other Company?
    ValidateWithCodeLists( PurOtherColumn, "YesNoUnknown")
    //Testing Pre-Fabricated Feature?
    ValidateWithCodeLists( PreFabColumn, "YesNoUnknown")
    //Testing Fabricated Assembly?
    ValidateWithCodeLists( FabAssColumn, "YesNoUnknown")
    //Testing Reconditioned / Salvaged?
    ValidateWithCodeLists( ReconSalvColumn, "YesNoUnknown")

    //******* Reference Document Images [ELS-EDMS] *******
    //Testing Drawing 1 Quality
    ValidateMutuallyInclusive( Dwg1Column, Dwg1QColumn)
    ValidateMutuallyInclusive( Dwg1QColumn, Dwg1Column)
    //Testing Drawing 2 Quality
    ValidateMutuallyInclusive( Dwg2Column, Dwg2QColumn)
    ValidateMutuallyInclusive( Dwg2QColumn, Dwg2Column)

    //******* Reference Document Images [ECTS] ******
    //Testing Image 1 Quality
    ValidateMutuallyInclusive( Img1Column, Img1QColumn)
    ValidateMutuallyInclusive( Img1QColumn, Img1Column)
    //Testing Image 2 Quality
    ValidateMutuallyInclusive( Img2Column, Img2QColumn)
    ValidateMutuallyInclusive( Img2QColumn, Img2Column)
    //Testing Image 3 Quality
    ValidateMutuallyInclusive( Img3Column, Img3QColumn)
    ValidateMutuallyInclusive( Img3QColumn, Img3Column)
    //Testing Image 4 Quality
    ValidateMutuallyInclusive( Img4Column, Img4QColumn)
    ValidateMutuallyInclusive( Img4QColumn, Img4Column)

    //******* Reference Maps *******
    //Testing Distribution Wall Map and Plat Sheet
    //******* Strength Test Information *******
    //Testing Type
    ValidateWithCodeLists( STPRTypeColumn, "PipeTestType")
    //Testing Media
    ValidateMutuallyInclusive( STPRMediaColumn, STPRTypeColumn)
    ValidateWithCodeLists( STPRMediaColumn, "TestMedium")
    //Testing Test Pressure
    ValidateMutuallyInclusive( STPRPressColumn, STPRTypeColumn)
    ValidateNumber( STPRPressColumn, 100, 5000)
    //Testing Duration (hrs)
    ValidateMutuallyInclusive( STPRDurColumn, STPRTypeColumn)
    ValidateNumber( STPRDurColumn, 0.5, 90)
    //Testing Adj Test Pressure
    ValidateNumber( STPRAdjPressColumn, 100, 5000)
    //Testing Test Date
    ValidateMutuallyInclusive( STPRDateColumn, STPRTypeColumn)
    ValidateDate( STPRDateColumn, MinDate=CDate("1/1/1901"), MaxDate=CDate("12/31/2020"))
    //Testing Date Prior Test
    ValidateDate( STPRDatePriorColumn, MinDate=CDate("1/1/1901"), MaxDate=CDate("12/31/2020"))
    //Testing Pressure Prior Test
    ValidateMutuallyInclusive( STPRPressPriorColumn, STPRDatePriorColumn)
    ValidateNumber( STPRPressPriorColumn, 100, 5000)
    //Testing Current MAOP
    ValidateNumber( STPRMAOPColumn, 35, 3000)
    
    //******* STPR Reference Images [ECTS] *******
    //Testing STPR Quality Validation
    with( PipeDataSht) {
        for( ForCounter = FirstRow To lastRow) {
            if( (.Cells(ForCounter, STPRImgColumn) != Empty || .Cells(ForCounter, STPRFChartColumn) != Empty || .Cells(ForCounter, STPRBChartColumn) != Empty || .Cells(ForCounter, STPRDeadLogColumn) != Empty || .Cells(ForCounter, STPRSketchColumn) != Empty) And .Cells(ForCounter, STPRQColumn) == Empty ) {
                MarkWrong( ForCounter, STPRQColumn)
            else if( .Cells(ForCounter, STPRImgColumn) = Empty And .Cells(ForCounter, STPRFChartColumn) = Empty And .Cells(ForCounter, STPRBChartColumn) = Empty And .Cells(ForCounter, STPRDeadLogColumn) = Empty And .Cells(ForCounter, STPRSketchColumn) = Empty And .Cells(ForCounter, STPRQColumn) != Empty ) {
                MarkWrong( ForCounter, STPRQColumn)
             }
        }
     }

    //Testing STPR Quality
    ValidateWithCodeLists( STPRQColumn, "QualityCodeSTPR")

    //******* Branch *******
    ValidateRequiredField( BLineColumn, FeatureVal="Tee")
    ValidateRequiredField( BLineColumn, FeatureVal="Tap", TypeVal="Drip-External,Drip-Internal,Insertion Probe,Thermowell,Take-off,Take-off Service Tee,Take-off Pin off-Valve Tee,Take-off Curb Valve Tee w/ nut,Other,Unknown")
    ValidateRequiredField( BLineColumn, FeatureVal="PCF", TypeVal="Unknown,Other,TDW,Bottom Tap,Side Tap,Top Tap")
    ValidateUnnamedBranches( BLineColumn)

    //******* X-Ray *******
    ValidateWithCodeLists( XRayColumn, "PercentXRay")

    //******* Responsible example Lan Id ********
    ValidateRequiredField( BuilderColumn)
    ValidateRequiredField( CheckerColumn)
    ValidateWithCodeLists( DiscrepancyColumn, "YesNo")

    PipeDataSht(.Activate)

    Application.ScreenUpdating = true
 }
 function ValidationCleanup(){
    Application.ScreenUpdating = false
    var BeginColumn  
    var EndColumn  
    //Setting Initial values
    BeginColumn( = "A")
    EndColumn( = "CW")
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        lastRow( = .Rows.Count)
        //Setting Name fields to no color
        .Range((.Cells(1, FeatureNoColumn), .Cells(1, "L")).Interior.Color = RGB(255, 255, 255))
        //Setting Feature Number cells background to no color
        .Range((.Cells(FirstRow, FeatureNoColumn), .Cells(lastRow, FeatureNoColumn)).Interior.Color = RGB(255, 255, 255))
        //Changing cells border
        with( .Range(BeginColumn + FirstRow + "" + EndColumn + lastRow)) {
            .Borders((xlDiagonalDown).LineStyle = xlNone)
            .Borders((xlDiagonalUp).LineStyle = xlNone)
            with( .Borders(xlEdgeLeft)) {
                .LineStyle( = xlContinuous)
                .ColorIndex( = 0)
                .Weight( = xlThin)
             }
            with( .Borders(xlEdgeTop)) {
                .LineStyle( = xlContinuous)
                .ColorIndex( = 0)
                .Weight( = xlThin)
             }
            with( .Borders(xlEdgeBottom)) {
                .LineStyle( = xlContinuous)
                .ColorIndex( = 0)
                .Weight( = xlThin)
             }
            with( .Borders(xlEdgeRight)) {
                .LineStyle( = xlContinuous)
                .ColorIndex( = 0)
                .Weight( = xlThin)
             }
            with( .Borders(xlInsideVertical)) {
                .LineStyle( = xlContinuous)
                .ColorIndex( = 0)
                .Weight( = xlThin)
             }
            with( .Borders(xlInsideHorizontal)) {
                .LineStyle( = xlContinuous)
                .ColorIndex( = 0)
                .Weight( = xlThin)
             }
         }
     }
 }
 function ValidateNumber(Col  , Min  ,  Max  ){
    var ForCounter  
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            if( .Range(Col + ForCounter) != Empty ) {
                if( ! IsError(.Range(Col + ForCounter)) ) {
                    if( IsNumeric(.Range(Col + ForCounter).Value) ) {
                        if( CDbl(.Range(Col + ForCounter).Value) < Min ) {
                            MarkWrong( ForCounter, Col)
                         }
                        if( Max > 0 And CDbl(.Range(Col + ForCounter).Value) > Max ) {
                            MarkWrong( ForCounter, Col)
                         }
                    }else{
                        MarkWrong( ForCounter, Col)
                     }
                }else{
                    MarkWrong( ForCounter, Col)
                 }
             }
        }
     }
 }
 function ValidateUnnamedBranches(Col  ){
    var ForCounter  
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            if( .Range(Col + ForCounter) != Empty ) {
                if( isUnnamed(.Range(Col + ForCounter).Value) ) {
                    if( ! CheckUnnamed(.Range(Col + ForCounter).Value) ) {
                        MarkWrong( ForCounter, Col)
                     }
                 }
             }
        }
     }
 }
 function ValidateWithCodeLists(Col  , List  ){
    var ForCounter  
    var Result  Boolean
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            Result = true
            if( .Range(Col + ForCounter) != Empty ) {
                if( List == "CheckType" ) {
                    Result = Check_Value(.Range(Col + ForCounter).Value, Replace$(.Range(FeatureColumn + ForCounter).Value, " ", ""))
                else if( List = "CheckSpecs" ) {
                    if( .Range(FeatureColumn + ForCounter).Value == "Pipe" ) {
                        Result = Check_Value(.Range(Col + ForCounter).Value, "PipeSpec")
                    else if( .Range(FeatureColumn + ForCounter).Value = "Sleeve" ) {
                        Result = Check_Value(.Range(Col + ForCounter).Value, "SleeveSpec")
                     }
                else if( List = "CheckRadius" ) {
                    if( .Range(FeatureColumn + ForCounter).Value == "Field Bend" ) {
                        Result = Check_Value(.Range(Col + ForCounter).Value, "BendRadius_Field")
                    else if( .Range(FeatureColumn + ForCounter).Value = "Mfg Bend" ) {
                        Result = Check_Value(.Range(Col + ForCounter).Value, "BendRadius")
                     }
                else if( List = "CheckOD" ) {
                    if( .Range(FeatureColumn + ForCounter).Value == "Pipe" ) {
                        Result = Check_Value(.Range(Col + ForCounter).Value, "OutsideDiameter")
                    else if( .Range(FeatureColumn + ForCounter).Value = "Sleeve" ) {
                        Result = Check_Value(.Range(Col + ForCounter).Value, "AnyOD2")
                     }
                }else{
                    Result = Check_Value(.Range(Col + ForCounter).Value, List)
                 }
                if( ! Result ) {
                    MarkWrong( ForCounter, Col)
                 }
             }
        }
     }
 }
 function Check_Value(Value, ListName  )  Boolean{
    var SearchArea  Range
    var r  Range
    var resultRange  Range
    Check_Value = false
     SearchArea = ThisWorkbook.Names(ListName).RefersToRange
    for ( var r in SearchArea) {
        if( IsNumeric(r.Value) ) {
            if( r.Value == Value ) {
                Check_Value = true
             }
        }else{
            if( UCase$(CStr(r.Value)) == UCase$(CStr(Value)) ) {
                Check_Value = true
             }
         }
    } r
 }
 function ValidateDate(Col  , MinDate  Date,  MaxDate  Date){
    var ValDate
    var ForCounter  
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            if( ! IsEmpty(.Range(Col + ForCounter)) ) {
                if( IsDate(.Range(Col + ForCounter).Value) ) {
                    ValDate = CDate(.Range(Col + ForCounter).Value)
                    if( ValDate > MinDate ) {
                        if( ! IsMissing(MaxDate) _
                        And ValDate > MaxDate ) {
                            MarkWrong( ForCounter, Col)
                         }
                    }else{
                        MarkWrong( ForCounter, Col)
                     }
                }else{
                    MarkWrong( ForCounter, Col)
                 }
             }
        }
     }
 }
 function ValidateMutuallyExclusive(CheckMeCol  , ExcludeMeCol  ,  FeatureVal$){
    var FeatFound  Boolean
    var ForCounter  
    var FV()  
    var i  
    if( FeatureVal$ != "" ) {
        FV( = Split(UCase$(FeatureVal$), ", "))
     }
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            FeatFound = true
            if( FeatureVal$ != "" ) {
                FeatFound = false
                for( i = LBound(FV) To UBound(FV)) {
                    if( UCase$(.Range(FeatureColumn + ForCounter).Value) == FV(i) ) {
                        FeatFound = true
                         break
                     }
                } i
             }
            if( FeatFound And (! IsEmpty(.Range(CheckMeCol + ForCounter))) ) {
                if( ! IsEmpty(.Range(ExcludeMeCol + ForCounter)) ) {
                    MarkWrong( ForCounter, ExcludeMeCol)
                 }
             }
        } ForCounter
     }
 }
 function ValidateMutuallyInclusive(CheckMeCol  , IncludeMeCol  ,  FeatureVal$){
    var FeatFound  Boolean
    var ForCounter  
    var FV()  
    var i  
    if( FeatureVal$ != "" ) {
        FV( = Split(UCase$(FeatureVal$), ", "))
     }
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            FeatFound = true
            if( FeatureVal$ != "" ) {
                FeatFound = false
                for( i = LBound(FV) To UBound(FV)) {
                    if( UCase$(.Range(FeatureColumn + ForCounter).Value) == FV(i) ) {
                        FeatFound = true
                         break
                     }
                } i
             }
            if( FeatFound And (! IsEmpty(.Range(CheckMeCol + ForCounter))) ) {
                if( IsEmpty(.Range(IncludeMeCol + ForCounter)) ) {
                    MarkWrong( ForCounter, CheckMeCol)
                 }
             }
        } ForCounter
     }
 }
 function ValidateRequiredField(Col  ,  FeatureVal$,  TypeVal$){
    var FeatFound  Boolean
    var ForCounter  
    var FV()  
    var i  
    var TV()  
    var TypeFound  Boolean
    var Var  Variant
    if( FeatureVal$ != "" ) {
        FV( = Split(UCase$(FeatureVal$), ", "))
        if( TypeVal$ != "" ) {
            TV( = Split(UCase$(TypeVal$), ", "))
         }
     }
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            FeatFound = true
            TypeFound = true
            if( FeatureVal$ != "" ) {
                FeatFound = false
                for( i = LBound(FV) To UBound(FV)) {
                    if( UCase$(.Range(FeatureColumn + ForCounter).Value) == FV(i) ) {
                        FeatFound = true
                         break
                     }
                } i
                if( FeatFound ) {
                    if( TypeVal$ != "" ) {
                        TypeFound = false
                        for( i = LBound(TV) To UBound(TV)) {
                            if( UCase$(.Range(TypeColumn + ForCounter).Value) == TV(i) ) {
                                TypeFound = true
                                 break
                             }
                        } i
                     }
                 }
             }
            if( FeatFound And TypeFound ) {
                if( IsEmpty(.Range(Col + ForCounter)) ) {
                    On Error Resume }
                     Var = .Range(Col + ForCounter).Validation.Type
                    On( Error GoTo 0)
                    if( IsEmpty(Var) ) {
                        MarkWrong( ForCounter, Col)
                    }else{
                        if( ! .Range(Col + ForCounter).Validation.Type == "2" ) {
                            MarkWrong( ForCounter, Col)
                        }else{
                            if( .Range(Col + ForCounter).Interior.Color != 65535 ) {
                                MarkWrong( ForCounter, Col)
                             }
                         }
                     }
                 }
             }
        } ForCounter
     }
 }
 function ValidateRequiredEmptyField(Col  ,  FeatureVal$,  TypeVal$){
    var FeatFound  Boolean
    var ForCounter  
    var FV()  
    var i  
    var TV()  
    var TypeFound  Boolean
    if( FeatureVal$ != "" ) {
        FV( = Split(UCase$(FeatureVal$), ", "))
        if( TypeVal$ != "" ) {
            TV( = Split(UCase$(TypeVal$), ", "))
         }
     }
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            FeatFound = true
            TypeFound = true
            if( FeatureVal$ != "" ) {
                FeatFound = false
                for( i = LBound(FV) To UBound(FV)) {
                    if( UCase$(.Range(FeatureColumn + ForCounter).Value) == FV(i) ) {
                        FeatFound = true
                         break
                     }
                } i
                if( FeatFound ) {
                    if( TypeVal$ != "" ) {
                        TypeFound = false
                        for( i = LBound(TV) To UBound(TV)) {
                            if( UCase$(.Range(TypeColumn + ForCounter).Value) == TV(i) ) {
                                TypeFound = true
                                 break
                             }
                        } i
                     }
                 }
             }
            if( FeatFound And TypeFound ) {
                if( ! IsEmpty(.Range(Col + ForCounter)) ) {
                    if( CStr(Application.Evaluate(.Range(Col + ForCounter).Formula)) != "" ) {
                        MarkWrong( ForCounter, Col)
                     }
                 }
             }
        } ForCounter
     }
 }
 function ValidateWithString(Col  , ListToCheck  ){
    var ForCounter  
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        for( ForCounter = FirstRow To lastRow) {
            if( ! IsEmpty(.Range(Col + ForCounter)) ) {
                if( FindInStrForward(UCase$(ListToCheck), UCase$(.Range(Col + ForCounter).Value)) == 0 ) {
                    MarkWrong( ForCounter, Col)
                 }
             }
        } ForCounter
     }
 }
 function MarkWrong(iRow  , Col  ,  CheckYellow  Boolean){
    with( ThisWorkbook.Worksheets("Pipe Data")) {
        if( (! CheckYellow) || (.Range(Col + iRow).Interior.ColorIndex != 6) ) {
            //Change feature cell background to green
            .Range(FeatureNoColumn + iRow).Interior.ColorIndex = 4
            //Change cell border
            with( .Range(Col + iRow)) {
                .Borders((xlDiagonalDown).LineStyle = xlNone)
                .Borders((xlDiagonalUp).LineStyle = xlNone)
                with( .Borders(xlEdgeLeft)) {
                    .LineStyle( = xlContinuous)
                    .ColorIndex( = 3)
                    .Weight( = xlThick)
                 }
                with( .Borders(xlEdgeTop)) {
                    .LineStyle( = xlContinuous)
                    .ColorIndex( = 3)
                    .Weight( = xlThick)
                 }
                with( .Borders(xlEdgeBottom)) {
                    .LineStyle( = xlContinuous)
                    .ColorIndex( = 3)
                    .Weight( = xlThick)
                 }
                with( .Borders(xlEdgeRight)) {
                    .LineStyle( = xlContinuous)
                    .ColorIndex( = 3)
                    .Weight( = xlThick)
                 }
                .Borders((xlInsideVertical).LineStyle = xlNone)
                .Borders((xlInsideHorizontal).LineStyle = xlNone)
             }
         }
     }
 }
 function CheckUnnamed(ShortName  )  Boolean{
    var Chr$
    var DateStr$
    var FindUSLng  
    var i  
    var ShortTypeList$
    var ShortTypes$()
    var StrArray$()
    var StrCount  
    var TempShortName$
    var TempShortNameRemaining$
    var TypeFound  Boolean
    TypeFound = false
    ShortTypeList$( = "DREG,DCUST,GCUST,STA,STUB,DRIP,BD,DFDS,DF")
    ShortTypes( = Split(ShortTypeList$, ","))
    TempShortName$( = UCase$(Trim$(ShortName)))
    for( i = 1 To Len(TempShortName$)) {
        Chr$( = Mid(TempShortName$, i, 1))
        if( ! Chr$ Like "[0-9,A-Z,_,-]" ) {
            CheckUnnamed = false
             function{
         }
    } i
    TempShortNameRemaining$( = TempShortName$)
    FindUSLng( = FindInStrForward(TempShortName$, "_"))
    while FindUSLng > 0) {
        StrCount( = StrCount + 1)
        ReDim( Preserve StrArray$(1 To StrCount))
        TempShortName$( = Left$(TempShortNameRemaining$, FindUSLng - 1))
        TempShortNameRemaining$( = Right$(TempShortNameRemaining$, Len(TempShortNameRemaining$) - FindUSLng))
        StrArray$((StrCount) = TempShortName$)
        FindUSLng( = FindInStrForward(TempShortNameRemaining$, "_"))
        if( FindUSLng == 0 ) {
            StrCount( = StrCount + 1)
            ReDim( Preserve StrArray$(1 To StrCount))
            StrArray$((StrCount) = TempShortNameRemaining$)
         }
    Wend()

    //Check length
    if( StrCount < 3 ) {
        CheckUnnamed = false
         function{
     }
    //Check prefix

    if( StrArray$(1) != "U" ) {
        CheckUnnamed = false
         function{
     }
    //Check short type

    for( i = 0 To UBound(ShortTypes$)) {
        if( ShortTypes(i) == StrArray$(2) ) {
            TypeFound = true
             break
         }
    } i
    if( ! TypeFound ) {
        CheckUnnamed = false
         function{
     }

    //Check date string is a number
    if( ! IsNumeric(StrArray$(3)) ) {
        CheckUnnamed = false
         function{
     }

    //Check the date string format is correct
    if( ! StrArray$(3) Like "############" ) {
        CheckUnnamed = false
         function{
    }else{
        //Check that the dates string//s year is valid
        if( CLng(Left$(StrArray$(3), 4)) > 2013 || CLng(Left$(StrArray$(3), 4)) < 2011 ) {
            CheckUnnamed = false
             function{
         }
     }

    //Check the source route information
    if( StrCount > 3 ) {
        for( i = 4 To StrCount) {
            //Check that unnamed follows the date pattern
            if( Left$(StrArray$(i), 1) == "U" ) {
                //Check that the date string is a number
                if( ! IsNumeric(Right$(StrArray$(i), Len(StrArray$(i)) - 1)) ) {
                    CheckUnnamed = false
                     function{
                 }
                //Check that the date string format is correct
                if( ! Right$(StrArray$(i), Len(StrArray$(i)) - 1) Like "############" ) {
                    CheckUnnamed = false
                     function{
                }else{
                    //Check that the date string//s year is valid
                    if( CLng(Left$(Right$(StrArray$(i), Len(StrArray$(i)) - 1), 4)) > 2013 _
                    || CLng(Left$(Right$(StrArray$(i), Len(StrArray$(i)) - 1), 4)) < 2011 ) {
                        CheckUnnamed = false
                         function{
                     }
                 }
            }else{
                //Check that //R// is the last string and that only U and R are used as prefixes
                if( Left$(StrArray$(i), 1) != "R" || StrCount > i ) {
                    CheckUnnamed = false
                     function{
                 }
             }
        } i
     }
    CheckUnnamed = true
 }
 function FindInStrForward(FindIn  , ToFind  )  {
    //searches for a character starting from the end of a string
    var FindCha  
    for( FindCha = 1 To (Len(FindIn) - Len(ToFind) + 1)) {
        if( Mid(FindIn, FindCha, Len(ToFind)) == ToFind ) {
            FindInStrForward( = FindCha)
             function{
         }
    } FindCha
 }
 function isUnnamed(ByVal Route$)  Boolean{
    var TrimRoute$
    TrimRoute$( = Left(UCase$(Trim(Route$)), 2))
    if( TrimRoute$ == "UN" || TrimRoute$ == "U_" ) {
        isUnnamed = true
     }
 }


