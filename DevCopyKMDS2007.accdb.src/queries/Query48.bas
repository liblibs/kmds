Operation =3
Name ="InvoLine"
Option =0
Where ="(((EugSouthHill.Apr)<>0))"
Begin InputTables
    Name ="EugSouthHill"
End
Begin OutputColumns
    Alias ="SetupDate"
    Name ="SetupDate"
    Expression ="#4/1/2018#"
    Name ="Quantity"
    Expression ="EugSouthHill.Apr"
    Alias ="InvoHeadID"
    Name ="InvoHeadID"
    Expression ="9999"
    Alias ="CashSysID"
    Name ="CashSysID"
    Expression ="99"
    Alias ="CashierID"
    Name ="CashierID"
    Expression ="99"
    Name ="Code"
    Expression ="EugSouthHill.Code"
    Alias ="SalesDept"
    Name ="SalesDept"
    Expression ="1"
    Alias ="InputCode"
    Name ="InputCode"
    Expression ="\"KMDS Convert\""
    Alias ="TransType"
    Name ="TransType"
    Expression ="\"S\""
    Name ="ID"
    Expression ="EugSouthHill.ID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetupDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.May"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Jun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Jul"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Aug"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Sep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Oct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Nov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Dec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Jan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Feb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Mar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EugSouthHill.Apr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-72
    Top =410
    Right =1219
    Bottom =1216
    Left =-1
    Top =-1
    Right =1259
    Bottom =505
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =250
        Bottom =337
        Top =0
        Name ="EugSouthHill"
        Name =""
    End
End
