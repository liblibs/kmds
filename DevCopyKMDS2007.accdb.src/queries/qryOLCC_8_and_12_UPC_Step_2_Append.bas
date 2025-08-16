Operation =3
Name ="OLCC_8_and_12_UPC"
Option =0
Where ="(((Len([dbo_OLCCScanCode].[UPCCode]))=10))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
    Name ="qryOLCC_8_and_12_UPC_FirstChar"
End
Begin OutputColumns
    Name ="Code"
    Expression ="dbo_OLCCScanCode.Code"
    Alias ="lid"
    Name ="lid"
    Expression ="Len([dbo_OLCCScanCode].[UPCCode])"
    Name ="ThirdDigit"
    Expression ="qryOLCC_8_and_12_UPC_FirstChar.ThirdDigit"
    Name ="UPCCode"
    Expression ="dbo_OLCCScanCode.UPCCode"
    Alias ="upc12"
    Name ="upc12"
    Expression ="checkdigit11([thirddigit] & [dbo_OLCCScanCode].[upccode])"
End
Begin Joins
    LeftTable ="dbo_OLCCScanCode"
    RightTable ="qryOLCC_8_and_12_UPC_FirstChar"
    Expression ="dbo_OLCCScanCode.Code = qryOLCC_8_and_12_UPC_FirstChar.Code"
    Flag =1
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
End
Begin
    State =0
    Left =46
    Top =126
    Right =1229
    Bottom =836
    Left =-1
    Top =-1
    Right =1160
    Bottom =408
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =472
        Top =-3
        Right =616
        Bottom =141
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =287
        Bottom =156
        Top =0
        Name ="qryOLCC_8_and_12_UPC_FirstChar"
        Name =""
    End
End
