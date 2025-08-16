Operation =3
Name ="dbo_OLCCScanCode"
Option =0
Where ="(((Len([dbo_OLCCScanCode].[UPCCode]))=10))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
    Name ="qrydbo_OLCCUPCItem_01_UPC_FirstChar"
End
Begin OutputColumns
    Name ="Code"
    Expression ="dbo_OLCCScanCode.Code"
    Alias ="upc12"
    Name ="UPCCode"
    Expression ="checkdigit11([thirddigit] & [dbo_OLCCScanCode].[upccode])"
End
Begin Joins
    LeftTable ="qrydbo_OLCCUPCItem_01_UPC_FirstChar"
    RightTable ="dbo_OLCCScanCode"
    Expression ="qrydbo_OLCCUPCItem_01_UPC_FirstChar.Code = dbo_OLCCScanCode.Code"
    Flag =1
End
Begin Groups
    Expression ="dbo_OLCCScanCode.Code"
    GroupLevel =0
    Expression ="checkdigit11([thirddigit] & [dbo_OLCCScanCode].[upccode])"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="dbo_OLCCScanCode.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCScanCode12DigitCalc.ThirdDigit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCCScanCode12DigitCalc.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="upc12"
        dbInteger "ColumnWidth" ="2475"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryOLCC_8_and_12_UPC_FirstChar.ThirdDigit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Len([dbo_OLCCScanCode].[UPCCode])"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =205
    Top =46
    Right =1388
    Bottom =756
    Left =-1
    Top =-1
    Right =1160
    Bottom =328
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =462
        Top =13
        Right =606
        Bottom =157
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
    Begin
        Left =33
        Top =13
        Right =177
        Bottom =157
        Top =0
        Name ="qrydbo_OLCCUPCItem_01_UPC_FirstChar"
        Name =""
    End
End
