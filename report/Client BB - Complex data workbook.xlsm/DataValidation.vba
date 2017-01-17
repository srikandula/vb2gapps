Attribute VB_Name = "DataValidation"
Option Explicit
Public Version
Public Const FirstRow = 3
Public Const ActSizeColumn = "W"
Public Const AngleColumn = "AD"
Public Const ANSIColumn = "V"
Public Const BarredColumn = "AG"
Public Const BegStaColumn = "A"
Public Const BLineColumn = "CG"
Public Const BuilderColumn = "CI"
Public Const CasTypeColumn = "AN"
Public Const CheckerColumn = "CJ"
Public Const ClassLocColumn = "D"
Public Const CoatDateColumn = "AR"
Public Const CoatTypeColumn = "AP"
Public Const DiscrepancyColumn = "CL"
Public Const DPlatColumn = "BN"
Public Const Dwg1Column = "BA"
Public Const Dwg1QColumn = "BB"
Public Const Dwg2Column = "BC"
Public Const Dwg2QColumn = "BD"
Public Const EndConnectColumn = "AB"
Public Const EndStaColumn = "B"
Public Const FabAssColumn = "AY"
Public Const FeatureColumn = "H"
Public Const FeatureNoColumn = "G"
Public Const FieldSTColumn = "M"
Public Const GISPipeSegIDColumn = "C"
Public Const Img10Column = "CM"
Public Const Img11Column = "CN"
Public Const Img12Column = "CO"
Public Const Img13Column = "CP"
Public Const Img14Column = "CQ"
Public Const Img1Column = "BE"
Public Const Img1QColumn = "BF"
Public Const Img2Column = "BG"
Public Const Img2QColumn = "BH"
Public Const Img3Column = "BI"
Public Const Img3QColumn = "BJ"
Public Const Img4Column = "BK"
Public Const Img4QColumn = "BL"
Public Const InsertionColumn = "AI"
Public Const InstallDateColumn = "F"
Public Const InsulatorColumn = "AO"
Public Const JobNoColumn = "E"
Public Const LengthColumn = "J"
Public Const MaterialColumn = "AC"
Public Const MaxPressColumn = "X"
Public Const MethodColumn = "AH"
Public Const MPColumn = "L"
Public Const OD1Column = "O"
Public Const OD2Column = "Q"
Public Const OperatorColumn = "AK"
Public Const OrientationColumn = "AF"
Public Const PercentSMYSColumn = "CW"
Public Const PipeSTColumn = "N"
Public Const PreFabColumn = "AX"
Public Const PurDateColumn = "AV"
Public Const PurOtherColumn = "AW"
Public Const RadiusColumn = "AE"
Public Const ReconSalvColumn = "AZ"
Public Const SealColumn = "AM"
Public Const SeamTypeColumn = "S"
Public Const ShellPressColumn = "AJ"
Public Const SMYSColumn = "U"
Public Const SpecsColumn = "T"
Public Const STPRAdjPressColumn = "BT"
Public Const STPRBChartColumn = "CC"
Public Const STPRDateColumn = "BU"
Public Const STPRDatePriorColumn = "BX"
Public Const STPRDeadLogColumn = "CD"
Public Const STPRDurColumn = "BS"
Public Const STPRFChartColumn = "CB"
Public Const STPRImgColumn = "CA"
Public Const STPRMAOPColumn = "BZ"
Public Const STPRMediaColumn = "BQ"
Public Const STPRPressColumn = "BR"
Public Const STPRPressPriorColumn = "BY"
Public Const STPRQColumn = "CF"
Public Const STPRSketchColumn = "CE"
Public Const STPRTypeColumn = "BO"
Public Const TypeColumn = "I"
Public Const VentedColumn = "AL"
Public Const WT1Column = "P"
Public Const WT2Column = "R"
Public Const XRayColumn = "CH"
Public lastRow As Long
Public Sub DataValidation_Click()
    ValidateData
End Sub
Public Sub ValidClean_Click()
    ValidationCleanup
    ThisWorkbook.Worksheets("Pipe Data").Activate
End Sub
Private Sub ValidateData()
    Dim ForCounter As Long
    Dim ListName As Name
    Dim PipeDataSht As Worksheet
    Dim RowCtr As Long
    Application.ScreenUpdating = False
    Set PipeDataSht = ThisWorkbook.Worksheets("Pipe Data")

    ValidationCleanup

    'Count all the rows
    lastRow = PipeDataSht.Range("H" & PipeDataSht.Rows.Count).End(xlUp).row
    If lastRow < FirstRow Then Exit Sub

    'Remove corrupt List Names
    For Each ListName In ActiveWorkbook.Names
        If UCase$(ListName.RefersTo) Like ("*.XL*") Then
            ListName.Delete
        End If
    Next ListName

    With PipeDataSht
        'Check standard PFL data
        If IsEmpty(.Range("G1")) Then
            .Range("G1").Interior.ColorIndex = 4
        End If
        If Not IsEmpty(.Range("J1")) Then
            If IsEmpty(.Range("L1")) Then
                .Range("L1").Interior.ColorIndex = 4
            End If
        Else
            If Not IsEmpty(.Range("L1")) Then
                .Range("J1").Interior.ColorIndex = 4
            End If
        End If

        'Begin and End Mile Point
        If IsEmpty(.Range(MPColumn & FirstRow)) Then MarkWrong FirstRow, MPColumn
        If IsEmpty(.Range(MPColumn & lastRow)) Then MarkWrong lastRow, MPColumn

        'Iterate through each row
        For ForCounter = FirstRow To lastRow
            'Check % SMYS
            If IsNumeric(.Range(PercentSMYSColumn & ForCounter)) Then
                If .Range(PercentSMYSColumn & ForCounter).Value >= 1 Then
                    .Range(FeatureNoColumn & ForCounter).Interior.Color = 65535
                End If
            End If
            'Beg Station
            If Not IsEmpty(.Range(BegStaColumn & ForCounter)) Then
                If Not IsNumeric(.Range(BegStaColumn & ForCounter)) Then
                    MarkWrong ForCounter, BegStaColumn
                End If
            Else
                MarkWrong ForCounter, BegStaColumn
            End If
            'End Station
            If Not IsEmpty(.Range(EndStaColumn & ForCounter)) Then
                If Not IsNumeric(.Range(EndStaColumn & ForCounter)) Then
                    MarkWrong ForCounter, EndStaColumn
                End If
            Else
                MarkWrong ForCounter, EndStaColumn
            End If
            'Pipe Segment ID
            If .Range(GISPipeSegIDColumn & ForCounter) = Empty Then
                MarkWrong ForCounter, GISPipeSegIDColumn
            End If
            'Job Number
            If IsEmpty(.Range(JobNoColumn & ForCounter)) _
            And .Range(FeatureColumn & ForCounter).Value <> "Tap" Then
                MarkWrong ForCounter, JobNoColumn
            End If
            'Check Feature Number column
            If IsNumeric(.Range(FeatureNoColumn & ForCounter).Value) Then
                If ForCounter > FirstRow Then
                    If .Range(FeatureNoColumn & ForCounter).Value _
                    <= .Range(FeatureNoColumn & ForCounter - 1).Value Then
                        MarkWrong ForCounter, FeatureNoColumn
                    End If
                End If
                If ForCounter < lastRow Then
                    If .Range(FeatureNoColumn & ForCounter).Value _
                    >= .Range(FeatureNoColumn & ForCounter + 1).Value Then
                        MarkWrong ForCounter, FeatureNoColumn
                    End If
                End If
            Else
                MarkWrong ForCounter, FeatureNoColumn
            End If
            'Check Required Fields
            If IsEmpty(.Range(FeatureColumn & ForCounter)) Then
                MarkWrong ForCounter, FeatureColumn
            End If
            If IsEmpty(.Range(TypeColumn & ForCounter)) Then
                MarkWrong ForCounter, TypeColumn
            End If
            
            'Check length field
            'If .Range(LengthColumn & ForCounter).Value _
            <> .Range(EndStaColumn & ForCounter).Value _
            - .Range(BegStaColumn & ForCounter).Value Then
            '    MarkWrong ForCounter, LengthColumn
            'End If
            
            'Check Current MAOP field
            If IsEmpty(.Range(STPRMAOPColumn & ForCounter)) Then
                MarkWrong ForCounter, STPRMAOPColumn
            Else
                If Not IsNumeric(.Range(STPRMAOPColumn & ForCounter)) Then
                    MarkWrong ForCounter, STPRMAOPColumn
                End If
            End If
            
            'Check that seam type and install date line up
            If UCase$(.Range(SeamTypeColumn & ForCounter)) = UCase$("Unknown > 4 - Modern") Then
                If Not IsEmpty(.Range(InstallDateColumn & ForCounter)) Then
                    If IsDate(.Range(InstallDateColumn & ForCounter).Value) Then
                        If Year(CDate(.Range(InstallDateColumn & ForCounter).Value)) < 1960 Then
                            If IsEmpty(.Range(OD1Column & ForCounter)) Then
                                MarkWrong ForCounter, SeamTypeColumn
                            Else
                                If .Range(OD1Column & ForCounter).Value < 28 Then
                                    MarkWrong ForCounter, SeamTypeColumn
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
        Next ForCounter
    End With

    'Testing Class Loc.
    ValidateWithCodeLists ClassLocColumn, "ClassLocation"

    'Testing Install Date
    ValidateDate InstallDateColumn, CDate("1/1/1910"), CDate("12/31/2015")
    ValidateRequiredField InstallDateColumn

    'Testing Feature Type
    ValidateWithCodeLists FeatureColumn, "EventType"

    'Testing Type, Type List varies based on feature type
    ValidateWithCodeLists TypeColumn, "CheckType"

    'Testing Mile Point
    ValidateNumber MPColumn, Min:=0

    'Testing Field Station
    ValidateNumber FieldSTColumn, Min:=0

    'Testing Pipe Station
    ValidateNumber PipeSTColumn, Min:=0
       
    '******* MAOP Specifications *******
    'Testing O.D.1
    ValidateRequiredField OD1Column, FeatureVal:="Cap, Regulator, Pipe, Tap, Field Bend, Mfg Bend, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateRequiredEmptyField OD1Column, FeatureVal:="Appurtenance"
    ValidateWithCodeLists OD1Column, "CheckOD"
    'Testing W.T.1
    ValidateRequiredField WT1Column, FeatureVal:="Pipe, Field Bend, Mfg Bend, Tee, Reducer, Sleeve, Other"
    ValidateRequiredField WT1Column, FeatureVal:="Flange", TypeVal:="Lap Joint, Threaded, Reducing, Socket Weld, Insulating"
    ValidateRequiredEmptyField WT1Column, FeatureVal:="PCF, Valve, Appurtenance, Meter, Pig Trap, Regulator, Relief Valve"
    ValidateRequiredEmptyField WT1Column, FeatureVal:="Flange", TypeVal:="Weld Neck, Blind, Slip On"
    ValidateNumber WT1Column, Min:=0.01, Max:=1.5
    'Testing O.D.2
    ValidateRequiredField OD2Column, FeatureVal:="Tee, Reducer"
    ValidateRequiredField OD2Column, FeatureVal:="Mfg Bend, Flange", TypeVal:="Reducing"
    ValidateRequiredField OD2Column, FeatureVal:="Pipe", TypeVal:="Casing"
    ValidateRequiredEmptyField OD2Column, FeatureVal:="Field Bend, Flange, PCF, Tap, Valve, Sleeve, Appurtenance, Meter, Pig Trap, Regulator, Relief Valve, Cap, Other"
    ValidateRequiredEmptyField OD2Column, FeatureVal:="Mfg Bend", TypeVal:="Forged, Miter, Bell Bell Chill Ring, Wrinkle, Rolled Plate, Socket Weld"
    ValidateRequiredEmptyField OD2Column, FeatureVal:="Pipe", TypeVal:="No Casing,Pipe Bridge,Pipe Span,Pipe Liner, Pipe Encased"
    ValidateRequiredEmptyField OD2Column, FeatureVal:="Flange", TypeVal:="Weld Neck, Blind, Slip On, Lap Joint, Threaded, Socket Weld, Insulating"
    ValidateWithCodeLists OD2Column, "OutsideDiameter"
    'Testing W.T.2
    ValidateRequiredField WT2Column, FeatureVal:="Tee, Reducer"
    ValidateRequiredField WT2Column, FeatureVal:="Mfg Bend, Flange", TypeVal:="Reducing"
    ValidateRequiredField WT2Column, FeatureVal:="Pipe", TypeVal:="Casing"
    ValidateRequiredEmptyField WT2Column, FeatureVal:="Pipe", TypeVal:="No Casing,Pipe Bridge,Pipe Span,Pipe Liner, Pipe Encased"
    ValidateRequiredEmptyField WT2Column, FeatureVal:="Field Bend, Flange, PCF, Tap, Valve, Sleeve, Appurtenance, Meter, Pig Trap, Regulator, Relief Valve, Cap, Other"
    ValidateRequiredEmptyField WT2Column, FeatureVal:="Pipe", TypeVal:="No Casing,Pipe Bridge,Pipe Span,Pipe Liner, Pipe Encased"
    ValidateRequiredEmptyField WT2Column, FeatureVal:="Mfg Bend", TypeVal:="Forged, Miter, Bell Bell Chill Ring, Wrinkle, Rolled Plate, Socket Weld"
    ValidateRequiredEmptyField WT2Column, FeatureVal:="Flange", TypeVal:="Weld Neck, Blind, Slip On, Lap Joint, Threaded, Socket Weld, Insulating"
    ValidateNumber WT2Column, 0.01, 1.5
    'Testing Seam Type
    ValidateRequiredField SeamTypeColumn, FeatureVal:="Pipe, Field Bend, Mfg Bend, Tee, Reducer"
    ValidateRequiredEmptyField SeamTypeColumn, FeatureVal:="Flange, Appurtenance, Pig Trap"
    ValidateWithCodeLists SeamTypeColumn, "SeamType"
    'Testing Specification / Rating
    ValidateRequiredField SpecsColumn, FeatureVal:="Cap, Pipe, Tap, Field Bend, Mfg Bend, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other"
    ValidateRequiredField SpecsColumn, FeatureVal:="Relief Valve"
    ValidateRequiredEmptyField SpecsColumn, FeatureVal:="Valve, Appurtenance, Regulator"
    ValidateWithCodeLists SpecsColumn, "CheckSpecs"
    'Testing SMYS Column
    ValidateRequiredField SMYSColumn, FeatureVal:="Cap, Pipe, Tap, Field Bend, Mfg Bend, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other"
    ValidateRequiredEmptyField SMYSColumn, FeatureVal:="Valve, Appurtenance, Regulator"
    ValidateWithCodeLists SMYSColumn, "SMYS"
    'Testing ANSI - WOG Rating
    ValidateRequiredField ANSIColumn, FeatureVal:="Regulator, Meter, Cap, Tap, Sleeve, PCF, Flange, Relief Valve"
    ValidateRequiredEmptyField ANSIColumn, FeatureVal:="Field Bend, Pipe, Appurtenance, Pig Trap"
    ValidateRequiredEmptyField ANSIColumn, FeatureVal:="Other", TypeVal:="Collar"
    ValidateRequiredField ANSIColumn, FeatureVal:="Other", TypeVal:="Closure, Insulating Joint, Filter/Strainer, Separator, Pig Signal, Pressure Transmitter, Dresser Mechanical, Internal Baffle, Threaded Weldolet, Weldolet or variations, Plug, Union, Repair Can, Other, Unknown"
    ValidateMutuallyExclusive ANSIColumn, SMYSColumn, FeatureVal:="Reducer,Tee,Mfg Bend,Flange"
    ValidateWithCodeLists ANSIColumn, "Ansi"
    'Testing Actual Size or Opening
    ValidateRequiredField ActSizeColumn, FeatureVal:="Tap, Valve, Sleeve, Regulator, Relief Valve, Other"
    ValidateRequiredEmptyField ActSizeColumn, FeatureVal:="Cap, Pipe, Field Bend, Mfg Bend, Tee, Reducer, Appurtenance, PCF, Flange, Meter, Pig Trap"
    ValidateNumber ActSizeColumn, 0.1, 125
    'Testing Max. Working Pressure
    ValidateNumber MaxPressColumn, 100, 5000
    ValidateRequiredEmptyField MaxPressColumn, FeatureVal:="Field Bend, Pipe, Appurtenance, Meter, Pig Trap, Cap"
    ValidateRequiredField MaxPressColumn, FeatureVal:="Tap, Regulator, Sleeve, PCF, Flange, Relief Valve"
    ValidateRequiredField MaxPressColumn, FeatureVal:="Other", TypeVal:="Closure, Insulating Joint, Filter/Strainer, Separator, Pig Signal, Pressure Transmitter, Dresser Mechanical, Internal Baffle, Threaded Weldolet, Weldolet or variations, Plug, Union, Repair Can, Other, Unknown"
    '******* Feature Specs *******
    'Testing End Connect
    ValidateRequiredField EndConnectColumn, FeatureVal:="Pipe, Valve, Flange, Regulator, Other"
    ValidateWithCodeLists EndConnectColumn, "EndConnect"
    'Testing Material Type
    ValidateWithCodeLists MaterialColumn, "MaterialType"
    ValidateRequiredField MaterialColumn, FeatureVal:="Pipe, Mfg Bend, Sleeve"
    '******* Bend Data *******
    'Testing Angle
    ValidateRequiredField AngleColumn, FeatureVal:="Field Bend, Mfg Bend"
    ValidateRequiredEmptyField AngleColumn, FeatureVal:="Appurtenance, Cap, Regulator, Pipe, Tap, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateNumber AngleColumn, 0, 90
    'Testing Radius
    ValidateRequiredField RadiusColumn, FeatureVal:="Field Bend, Mfg Bend", TypeVal:="Forged, Bell Bell Chill Ring, Wrinkle, Rolled Plate, Reducing, Socket Weld, Unknown, Bending Machine, Int Mandrel-Bending Machine"
    ValidateRequiredEmptyField RadiusColumn, FeatureVal:="Appurtenance, Cap, Regulator, Pipe, Tap, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithCodeLists RadiusColumn, "CheckRadius"
    'Testing Orient
    ValidateRequiredField OrientationColumn, FeatureVal:="Field Bend, Mfg Bend"
    ValidateRequiredEmptyField OrientationColumn, FeatureVal:="Appurtenance, Cap, Regulator, Pipe, Tap, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithCodeLists OrientationColumn, "Orientation"
    
    '******* Tee Data *******
    'Testing Barred?
    ValidateRequiredField BarredColumn, FeatureVal:="Tee"
    ValidateRequiredEmptyField BarredColumn, FeatureVal:="Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tap, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithString BarredColumn, "Yes, No, Unknown"
    
    '******* Tap Data *******
    'Testing Method
    ValidateRequiredField MethodColumn, FeatureVal:="Tap"
    ValidateRequiredEmptyField MethodColumn, FeatureVal:="Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithCodeLists MethodColumn, "HotCold"
    'Testing Insertion
    ValidateRequiredField InsertionColumn, FeatureVal:="Tap"
    ValidateRequiredEmptyField InsertionColumn, FeatureVal:="Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Valve, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithString InsertionColumn, "Yes, No, Unknown"
    
    '******* Valves *******
    'Testing Shell Test Pressure
    ValidateRequiredField ShellPressColumn, FeatureVal:="Valve"
    ValidateRequiredEmptyField ShellPressColumn, FeatureVal:="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other"
    ValidateNumber ShellPressColumn, 100, 5000
    'Testing Operator Type
    ValidateRequiredField OperatorColumn, FeatureVal:="Valve"
    ValidateRequiredEmptyField OperatorColumn, FeatureVal:="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Pipe, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Other"
    ValidateWithCodeLists OperatorColumn, "ValveOperatorType"

    '******* Casing Data *******
    'Testing Vented?
    ValidateRequiredField VentedColumn, FeatureVal:="Pipe", TypeVal:="Casing"
    ValidateRequiredEmptyField VentedColumn, FeatureVal:="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithCodeLists VentedColumn, "YesNoUnknown"
    'Testing Seal Type
    ValidateRequiredField SealColumn, FeatureVal:="Pipe", TypeVal:="Casing"
    ValidateRequiredEmptyField SealColumn, FeatureVal:="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithCodeLists SealColumn, "CasingSealType"
    'Testing Type
    ValidateRequiredField CasTypeColumn, FeatureVal:="Pipe", TypeVal:="Casing"
    ValidateRequiredEmptyField CasTypeColumn, FeatureVal:="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithCodeLists CasTypeColumn, "CasingType"
    'Testing Insulator Type
    ValidateRequiredField InsulatorColumn, FeatureVal:="Pipe", TypeVal:="Casing"
    ValidateRequiredEmptyField InsulatorColumn, FeatureVal:="Tap, Appurtenance, Mfg Bend, Field Bend, Cap, Regulator, Valve, Tee, Reducer, Sleeve, PCF, Flange, Meter, Pig Trap, Relief Valve, Other"
    ValidateWithCodeLists InsulatorColumn, "CasingInsulateType"

    '******* Coating *******
    'Testing Coating Type
    ValidateWithCodeLists CoatTypeColumn, "CoatingType"
    'Testing Coating Install Date
    ValidateDate CoatDateColumn, MinDate:=CDate("1/1/1915"), MaxDate:=CDate("12/31/2015")

    '******* Purchase and Install Information *******
    'Testing Purchase Date of Feature
    ValidateDate PurDateColumn, MinDate:="1/1/1915", MaxDate:="1/1/2020"
    'Testing Purchased from Other Company?
    ValidateWithCodeLists PurOtherColumn, "YesNoUnknown"
    'Testing Pre-Fabricated Feature?
    ValidateWithCodeLists PreFabColumn, "YesNoUnknown"
    'Testing Fabricated Assembly?
    ValidateWithCodeLists FabAssColumn, "YesNoUnknown"
    'Testing Reconditioned / Salvaged?
    ValidateWithCodeLists ReconSalvColumn, "YesNoUnknown"

    '******* Reference Document Images [ELS-EDMS] *******
    'Testing Drawing 1 Quality
    ValidateMutuallyInclusive Dwg1Column, Dwg1QColumn
    ValidateMutuallyInclusive Dwg1QColumn, Dwg1Column
    'Testing Drawing 2 Quality
    ValidateMutuallyInclusive Dwg2Column, Dwg2QColumn
    ValidateMutuallyInclusive Dwg2QColumn, Dwg2Column

    '******* Reference Document Images [ECTS] ******
    'Testing Image 1 Quality
    ValidateMutuallyInclusive Img1Column, Img1QColumn
    ValidateMutuallyInclusive Img1QColumn, Img1Column
    'Testing Image 2 Quality
    ValidateMutuallyInclusive Img2Column, Img2QColumn
    ValidateMutuallyInclusive Img2QColumn, Img2Column
    'Testing Image 3 Quality
    ValidateMutuallyInclusive Img3Column, Img3QColumn
    ValidateMutuallyInclusive Img3QColumn, Img3Column
    'Testing Image 4 Quality
    ValidateMutuallyInclusive Img4Column, Img4QColumn
    ValidateMutuallyInclusive Img4QColumn, Img4Column

    '******* Reference Maps *******
    'Testing Distribution Wall Map and Plat Sheet
    '******* Strength Test Information *******
    'Testing Type
    ValidateWithCodeLists STPRTypeColumn, "PipeTestType"
    'Testing Media
    ValidateMutuallyInclusive STPRMediaColumn, STPRTypeColumn
    ValidateWithCodeLists STPRMediaColumn, "TestMedium"
    'Testing Test Pressure
    ValidateMutuallyInclusive STPRPressColumn, STPRTypeColumn
    ValidateNumber STPRPressColumn, 100, 5000
    'Testing Duration (hrs)
    ValidateMutuallyInclusive STPRDurColumn, STPRTypeColumn
    ValidateNumber STPRDurColumn, 0.5, 90
    'Testing Adj Test Pressure
    ValidateNumber STPRAdjPressColumn, 100, 5000
    'Testing Test Date
    ValidateMutuallyInclusive STPRDateColumn, STPRTypeColumn
    ValidateDate STPRDateColumn, MinDate:=CDate("1/1/1901"), MaxDate:=CDate("12/31/2020")
    'Testing Date Prior Test
    ValidateDate STPRDatePriorColumn, MinDate:=CDate("1/1/1901"), MaxDate:=CDate("12/31/2020")
    'Testing Pressure Prior Test
    ValidateMutuallyInclusive STPRPressPriorColumn, STPRDatePriorColumn
    ValidateNumber STPRPressPriorColumn, 100, 5000
    'Testing Current MAOP
    ValidateNumber STPRMAOPColumn, 35, 3000
    
    '******* STPR Reference Images [ECTS] *******
    'Testing STPR Quality Validation
    With PipeDataSht
        For ForCounter = FirstRow To lastRow
            If (.Cells(ForCounter, STPRImgColumn) <> Empty Or .Cells(ForCounter, STPRFChartColumn) <> Empty Or .Cells(ForCounter, STPRBChartColumn) <> Empty Or .Cells(ForCounter, STPRDeadLogColumn) <> Empty Or .Cells(ForCounter, STPRSketchColumn) <> Empty) And .Cells(ForCounter, STPRQColumn) = Empty Then
                MarkWrong ForCounter, STPRQColumn
            ElseIf .Cells(ForCounter, STPRImgColumn) = Empty And .Cells(ForCounter, STPRFChartColumn) = Empty And .Cells(ForCounter, STPRBChartColumn) = Empty And .Cells(ForCounter, STPRDeadLogColumn) = Empty And .Cells(ForCounter, STPRSketchColumn) = Empty And .Cells(ForCounter, STPRQColumn) <> Empty Then
                MarkWrong ForCounter, STPRQColumn
            End If
        Next
    End With

    'Testing STPR Quality
    ValidateWithCodeLists STPRQColumn, "QualityCodeSTPR"

    '******* Branch *******
    ValidateRequiredField BLineColumn, FeatureVal:="Tee"
    ValidateRequiredField BLineColumn, FeatureVal:="Tap", TypeVal:="Drip-External,Drip-Internal,Insertion Probe,Thermowell,Take-off,Take-off Service Tee,Take-off Pin off-Valve Tee,Take-off Curb Valve Tee w/ nut,Other,Unknown"
    ValidateRequiredField BLineColumn, FeatureVal:="PCF", TypeVal:="Unknown,Other,TDW,Bottom Tap,Side Tap,Top Tap"
    ValidateUnnamedBranches BLineColumn

    '******* X-Ray *******
    ValidateWithCodeLists XRayColumn, "PercentXRay"

    '******* Responsible example Lan Id ********
    ValidateRequiredField BuilderColumn
    ValidateRequiredField CheckerColumn
    ValidateWithCodeLists DiscrepancyColumn, "YesNo"

    PipeDataSht.Activate

    Application.ScreenUpdating = True
End Sub
Private Sub ValidationCleanup()
    Application.ScreenUpdating = False
    Dim BeginColumn As String
    Dim EndColumn As String
    'Setting Initial values
    BeginColumn = "A"
    EndColumn = "CW"
    With ThisWorkbook.Worksheets("Pipe Data")
        lastRow = .Rows.Count
        'Setting Name fields to no color
        .Range(.Cells(1, FeatureNoColumn), .Cells(1, "L")).Interior.Color = RGB(255, 255, 255)
        'Setting Feature Number cells background to no color
        .Range(.Cells(FirstRow, FeatureNoColumn), .Cells(lastRow, FeatureNoColumn)).Interior.Color = RGB(255, 255, 255)
        'Changing cells border
        With .Range(BeginColumn & FirstRow & ":" & EndColumn & lastRow)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .Weight = xlThin
            End With
        End With
    End With
End Sub
Private Sub ValidateNumber(Col As String, Min As Double, Optional Max As Double)
    Dim ForCounter As Long
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            If .Range(Col & ForCounter) <> Empty Then
                If Not IsError(.Range(Col & ForCounter)) Then
                    If IsNumeric(.Range(Col & ForCounter).Value) Then
                        If CDbl(.Range(Col & ForCounter).Value) < Min Then
                            MarkWrong ForCounter, Col
                        End If
                        If Max > 0 And CDbl(.Range(Col & ForCounter).Value) > Max Then
                            MarkWrong ForCounter, Col
                        End If
                    Else
                        MarkWrong ForCounter, Col
                    End If
                Else
                    MarkWrong ForCounter, Col
                End If
            End If
        Next
    End With
End Sub
Private Sub ValidateUnnamedBranches(Col As String)
    Dim ForCounter As Long
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            If .Range(Col & ForCounter) <> Empty Then
                If isUnnamed(.Range(Col & ForCounter).Value) Then
                    If Not CheckUnnamed(.Range(Col & ForCounter).Value) Then
                        MarkWrong ForCounter, Col
                    End If
                End If
            End If
        Next
    End With
End Sub
Private Sub ValidateWithCodeLists(Col As String, List As String)
    Dim ForCounter As Long
    Dim Result As Boolean
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            Result = True
            If .Range(Col & ForCounter) <> Empty Then
                If List = "CheckType" Then
                    Result = Check_Value(.Range(Col & ForCounter).Value, Replace$(.Range(FeatureColumn & ForCounter).Value, " ", ""))
                ElseIf List = "CheckSpecs" Then
                    If .Range(FeatureColumn & ForCounter).Value = "Pipe" Then
                        Result = Check_Value(.Range(Col & ForCounter).Value, "PipeSpec")
                    ElseIf .Range(FeatureColumn & ForCounter).Value = "Sleeve" Then
                        Result = Check_Value(.Range(Col & ForCounter).Value, "SleeveSpec")
                    End If
                ElseIf List = "CheckRadius" Then
                    If .Range(FeatureColumn & ForCounter).Value = "Field Bend" Then
                        Result = Check_Value(.Range(Col & ForCounter).Value, "BendRadius_Field")
                    ElseIf .Range(FeatureColumn & ForCounter).Value = "Mfg Bend" Then
                        Result = Check_Value(.Range(Col & ForCounter).Value, "BendRadius")
                    End If
                ElseIf List = "CheckOD" Then
                    If .Range(FeatureColumn & ForCounter).Value = "Pipe" Then
                        Result = Check_Value(.Range(Col & ForCounter).Value, "OutsideDiameter")
                    ElseIf .Range(FeatureColumn & ForCounter).Value = "Sleeve" Then
                        Result = Check_Value(.Range(Col & ForCounter).Value, "AnyOD2")
                    End If
                Else
                    Result = Check_Value(.Range(Col & ForCounter).Value, List)
                End If
                If Not Result Then
                    MarkWrong ForCounter, Col
                End If
            End If
        Next
    End With
End Sub
Private Function Check_Value(Value, ListName As String) As Boolean
    Dim SearchArea As Range
    Dim r As Range
    Dim resultRange As Range
    Check_Value = False
    Set SearchArea = ThisWorkbook.Names(ListName).RefersToRange
    For Each r In SearchArea
        If IsNumeric(r.Value) Then
            If r.Value = Value Then
                Check_Value = True
            End If
        Else
            If UCase$(CStr(r.Value)) = UCase$(CStr(Value)) Then
                Check_Value = True
            End If
        End If
    Next r
End Function
Private Sub ValidateDate(Col As String, MinDate As Date, Optional MaxDate As Date)
    Dim ValDate
    Dim ForCounter As Long
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            If Not IsEmpty(.Range(Col & ForCounter)) Then
                If IsDate(.Range(Col & ForCounter).Value) Then
                    ValDate = CDate(.Range(Col & ForCounter).Value)
                    If ValDate > MinDate Then
                        If Not IsMissing(MaxDate) _
                        And ValDate > MaxDate Then
                            MarkWrong ForCounter, Col
                        End If
                    Else
                        MarkWrong ForCounter, Col
                    End If
                Else
                    MarkWrong ForCounter, Col
                End If
            End If
        Next
    End With
End Sub
Private Sub ValidateMutuallyExclusive(CheckMeCol As String, ExcludeMeCol As String, Optional FeatureVal$)
    Dim FeatFound As Boolean
    Dim ForCounter As Long
    Dim FV() As String
    Dim i As Long
    If FeatureVal$ <> "" Then
        FV = Split(UCase$(FeatureVal$), ", ")
    End If
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            FeatFound = True
            If FeatureVal$ <> "" Then
                FeatFound = False
                For i = LBound(FV) To UBound(FV)
                    If UCase$(.Range(FeatureColumn & ForCounter).Value) = FV(i) Then
                        FeatFound = True
                        Exit For
                    End If
                Next i
            End If
            If FeatFound And (Not IsEmpty(.Range(CheckMeCol & ForCounter))) Then
                If Not IsEmpty(.Range(ExcludeMeCol & ForCounter)) Then
                    MarkWrong ForCounter, ExcludeMeCol
                End If
            End If
        Next ForCounter
    End With
End Sub
Private Sub ValidateMutuallyInclusive(CheckMeCol As String, IncludeMeCol As String, Optional FeatureVal$)
    Dim FeatFound As Boolean
    Dim ForCounter As Long
    Dim FV() As String
    Dim i As Long
    If FeatureVal$ <> "" Then
        FV = Split(UCase$(FeatureVal$), ", ")
    End If
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            FeatFound = True
            If FeatureVal$ <> "" Then
                FeatFound = False
                For i = LBound(FV) To UBound(FV)
                    If UCase$(.Range(FeatureColumn & ForCounter).Value) = FV(i) Then
                        FeatFound = True
                        Exit For
                    End If
                Next i
            End If
            If FeatFound And (Not IsEmpty(.Range(CheckMeCol & ForCounter))) Then
                If IsEmpty(.Range(IncludeMeCol & ForCounter)) Then
                    MarkWrong ForCounter, CheckMeCol
                End If
            End If
        Next ForCounter
    End With
End Sub
Private Sub ValidateRequiredField(Col As String, Optional FeatureVal$, Optional TypeVal$)
    Dim FeatFound As Boolean
    Dim ForCounter As Long
    Dim FV() As String
    Dim i As Long
    Dim TV() As String
    Dim TypeFound As Boolean
    Dim Var As Variant
    If FeatureVal$ <> "" Then
        FV = Split(UCase$(FeatureVal$), ", ")
        If TypeVal$ <> "" Then
            TV = Split(UCase$(TypeVal$), ", ")
        End If
    End If
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            FeatFound = True
            TypeFound = True
            If FeatureVal$ <> "" Then
                FeatFound = False
                For i = LBound(FV) To UBound(FV)
                    If UCase$(.Range(FeatureColumn & ForCounter).Value) = FV(i) Then
                        FeatFound = True
                        Exit For
                    End If
                Next i
                If FeatFound Then
                    If TypeVal$ <> "" Then
                        TypeFound = False
                        For i = LBound(TV) To UBound(TV)
                            If UCase$(.Range(TypeColumn & ForCounter).Value) = TV(i) Then
                                TypeFound = True
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
            If FeatFound And TypeFound Then
                If IsEmpty(.Range(Col & ForCounter)) Then
                    On Error Resume Next
                    Set Var = .Range(Col & ForCounter).Validation.Type
                    On Error GoTo 0
                    If IsEmpty(Var) Then
                        MarkWrong ForCounter, Col
                    Else
                        If Not .Range(Col & ForCounter).Validation.Type = "2" Then
                            MarkWrong ForCounter, Col
                        Else
                            If .Range(Col & ForCounter).Interior.Color <> 65535 Then
                                MarkWrong ForCounter, Col
                            End If
                        End If
                    End If
                End If
            End If
        Next ForCounter
    End With
End Sub
Private Sub ValidateRequiredEmptyField(Col As String, Optional FeatureVal$, Optional TypeVal$)
    Dim FeatFound As Boolean
    Dim ForCounter As Long
    Dim FV() As String
    Dim i As Long
    Dim TV() As String
    Dim TypeFound As Boolean
    If FeatureVal$ <> "" Then
        FV = Split(UCase$(FeatureVal$), ", ")
        If TypeVal$ <> "" Then
            TV = Split(UCase$(TypeVal$), ", ")
        End If
    End If
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            FeatFound = True
            TypeFound = True
            If FeatureVal$ <> "" Then
                FeatFound = False
                For i = LBound(FV) To UBound(FV)
                    If UCase$(.Range(FeatureColumn & ForCounter).Value) = FV(i) Then
                        FeatFound = True
                        Exit For
                    End If
                Next i
                If FeatFound Then
                    If TypeVal$ <> "" Then
                        TypeFound = False
                        For i = LBound(TV) To UBound(TV)
                            If UCase$(.Range(TypeColumn & ForCounter).Value) = TV(i) Then
                                TypeFound = True
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
            If FeatFound And TypeFound Then
                If Not IsEmpty(.Range(Col & ForCounter)) Then
                    If CStr(Application.Evaluate(.Range(Col & ForCounter).Formula)) <> "" Then
                        MarkWrong ForCounter, Col
                    End If
                End If
            End If
        Next ForCounter
    End With
End Sub
Private Sub ValidateWithString(Col As String, ListToCheck As String)
    Dim ForCounter As Long
    With ThisWorkbook.Worksheets("Pipe Data")
        For ForCounter = FirstRow To lastRow
            If Not IsEmpty(.Range(Col & ForCounter)) Then
                If FindInStrForward(UCase$(ListToCheck), UCase$(.Range(Col & ForCounter).Value)) = 0 Then
                    MarkWrong ForCounter, Col
                End If
            End If
        Next ForCounter
    End With
End Sub
Private Sub MarkWrong(iRow As Long, Col As String, Optional CheckYellow As Boolean)
    With ThisWorkbook.Worksheets("Pipe Data")
        If (Not CheckYellow) Or (.Range(Col & iRow).Interior.ColorIndex <> 6) Then
            'Change feature cell background to green
            .Range(FeatureNoColumn & iRow).Interior.ColorIndex = 4
            'Change cell border
            With .Range(Col & iRow)
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 3
                    .Weight = xlThick
                End With
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 3
                    .Weight = xlThick
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 3
                    .Weight = xlThick
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 3
                    .Weight = xlThick
                End With
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End With
        End If
    End With
End Sub
Private Function CheckUnnamed(ShortName As String) As Boolean
    Dim Chr$
    Dim DateStr$
    Dim FindUSLng As Long
    Dim i As Long
    Dim ShortTypeList$
    Dim ShortTypes$()
    Dim StrArray$()
    Dim StrCount As Long
    Dim TempShortName$
    Dim TempShortNameRemaining$
    Dim TypeFound As Boolean
    TypeFound = False
    ShortTypeList$ = "DREG,DCUST,GCUST,STA,STUB,DRIP,BD,DFDS,DF"
    ShortTypes = Split(ShortTypeList$, ",")
    TempShortName$ = UCase$(Trim$(ShortName))
    For i = 1 To Len(TempShortName$)
        Chr$ = Mid(TempShortName$, i, 1)
        If Not Chr$ Like "[0-9,A-Z,_,-]" Then
            CheckUnnamed = False
            Exit Function
        End If
    Next i
    TempShortNameRemaining$ = TempShortName$
    FindUSLng = FindInStrForward(TempShortName$, "_")
    While FindUSLng > 0
        StrCount = StrCount + 1
        ReDim Preserve StrArray$(1 To StrCount)
        TempShortName$ = Left$(TempShortNameRemaining$, FindUSLng - 1)
        TempShortNameRemaining$ = Right$(TempShortNameRemaining$, Len(TempShortNameRemaining$) - FindUSLng)
        StrArray$(StrCount) = TempShortName$
        FindUSLng = FindInStrForward(TempShortNameRemaining$, "_")
        If FindUSLng = 0 Then
            StrCount = StrCount + 1
            ReDim Preserve StrArray$(1 To StrCount)
            StrArray$(StrCount) = TempShortNameRemaining$
        End If
    Wend

    'Check length
    If StrCount < 3 Then
        CheckUnnamed = False
        Exit Function
    End If
    'Check prefix

    If StrArray$(1) <> "U" Then
        CheckUnnamed = False
        Exit Function
    End If
    'Check short type

    For i = 0 To UBound(ShortTypes$)
        If ShortTypes(i) = StrArray$(2) Then
            TypeFound = True
            Exit For
        End If
    Next i
    If Not TypeFound Then
        CheckUnnamed = False
        Exit Function
    End If

    'Check date string is a number
    If Not IsNumeric(StrArray$(3)) Then
        CheckUnnamed = False
        Exit Function
    End If

    'Check the date string format is correct
    If Not StrArray$(3) Like "############" Then
        CheckUnnamed = False
        Exit Function
    Else
        'Check that the dates string's year is valid
        If CLng(Left$(StrArray$(3), 4)) > 2013 Or CLng(Left$(StrArray$(3), 4)) < 2011 Then
            CheckUnnamed = False
            Exit Function
        End If
    End If

    'Check the source route information
    If StrCount > 3 Then
        For i = 4 To StrCount
            'Check that unnamed follows the date pattern
            If Left$(StrArray$(i), 1) = "U" Then
                'Check that the date string is a number
                If Not IsNumeric(Right$(StrArray$(i), Len(StrArray$(i)) - 1)) Then
                    CheckUnnamed = False
                    Exit Function
                End If
                'Check that the date string format is correct
                If Not Right$(StrArray$(i), Len(StrArray$(i)) - 1) Like "############" Then
                    CheckUnnamed = False
                    Exit Function
                Else
                    'Check that the date string's year is valid
                    If CLng(Left$(Right$(StrArray$(i), Len(StrArray$(i)) - 1), 4)) > 2013 _
                    Or CLng(Left$(Right$(StrArray$(i), Len(StrArray$(i)) - 1), 4)) < 2011 Then
                        CheckUnnamed = False
                        Exit Function
                    End If
                End If
            Else
                'Check that 'R' is the last string and that only U and R are used as prefixes
                If Left$(StrArray$(i), 1) <> "R" Or StrCount > i Then
                    CheckUnnamed = False
                    Exit Function
                End If
            End If
        Next i
    End If
    CheckUnnamed = True
End Function
Private Function FindInStrForward(FindIn As String, ToFind As String) As Long
    'searches for a character starting from the end of a string
    Dim FindCha As Long
    For FindCha = 1 To (Len(FindIn) - Len(ToFind) + 1)
        If Mid(FindIn, FindCha, Len(ToFind)) = ToFind Then
            FindInStrForward = FindCha
            Exit Function
        End If
    Next FindCha
End Function
Private Function isUnnamed(ByVal Route$) As Boolean
    Dim TrimRoute$
    TrimRoute$ = Left(UCase$(Trim(Route$)), 2)
    If TrimRoute$ = "UN" Or TrimRoute$ = "U_" Then
        isUnnamed = True
    End If
End Function


