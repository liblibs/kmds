Operation =3
Name ="InvoLine"
Option =0
Begin InputTables
    Name ="Sheet1"
    Name ="InvoLine"
End
Begin OutputColumns
    Alias ="InvoHeadID"
    Name ="InvoHeadID"
    Expression ="99901"
    Alias ="SetupDate"
    Name ="SetupDate"
    Expression ="#11/1/2019#"
    Alias ="Quantity"
    Name ="Quantity"
    Expression ="Sheet1.NOV"
    Name ="ID"
    Expression ="Sheet1.ID"
    Alias ="Code"
    Name ="Code"
    Expression ="Sheet1.ITEMID"
    Alias ="InputCode"
    Name ="InputCode"
    Expression ="\"KMDS Conv\""
    Alias ="TransType"
    Name ="TransType"
    Expression ="\"S\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet1.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetupDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =32
    Top =49
    Right =1259
    Bottom =790
    Left =-1
    Top =-1
    Right =1195
    Bottom =457
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =56
        Top =42
        Right =200
        Bottom =186
        Top =0
        Name ="Sheet1"
        Name =""
    End
    Begin
        Left =371
        Top =18
        Right =515
        Bottom =458
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
