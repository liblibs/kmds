Operation =3
Name ="InvoLine"
Option =0
Where ="(((sheet2x.Aug)<>0))"
Begin InputTables
    Name ="sheet2x"
End
Begin OutputColumns
    Alias ="sd"
    Name ="SetupDate"
    Expression ="#8/2/2020#"
    Name ="Quantity"
    Expression ="sheet2x.Aug"
    Alias ="InvoHeadID"
    Name ="InvoHeadID"
    Expression ="999"
    Alias ="CashSysID"
    Name ="CashSysID"
    Expression ="99"
    Alias ="CashierID"
    Name ="CashierID"
    Expression ="99"
    Name ="Code"
    Expression ="sheet2x.ItemID"
    Alias ="InputCode"
    Name ="InputCode"
    Expression ="\"KMDS Conversion\""
    Alias ="TransType"
    Name ="TransType"
    Expression ="\"S\""
    Name ="id"
    Expression ="sheet2x.id"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet2.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sheet2.Aug"
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
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.Aug"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.sep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.oct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.nov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.dec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.jan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.feb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.mar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.apr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.may"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.jun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sheet2x.jul"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-220
    Top =98
    Right =1286
    Bottom =934
    Left =-1
    Top =-1
    Right =1474
    Bottom =535
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =120
        Top =18
        Right =264
        Bottom =473
        Top =0
        Name ="sheet2x"
        Name =""
    End
End
