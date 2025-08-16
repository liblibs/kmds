Operation =1
Option =0
Where ="(((InvoLine5.Code)=10) AND ((InvoLine5.TransType)=\"T\"))"
Begin InputTables
    Name ="InvoLine5"
End
Begin OutputColumns
    Expression ="InvoLine5.InvoHeadID"
    Expression ="InvoLine5.ID"
    Expression ="InvoLine5.CashSysID"
    Expression ="InvoLine5.SetupDate"
    Expression ="InvoLine5.Code"
    Expression ="InvoLine5.TransType"
    Expression ="InvoLine5.CashierID"
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
Begin
    Begin
        dbText "Name" ="InvoLine5.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.CashierID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-560
    Top =146
    Right =992
    Bottom =941
    Left =-1
    Top =-1
    Right =1529
    Bottom =345
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =98
        Top =16
        Right =242
        Bottom =332
        Top =0
        Name ="InvoLine5"
        Name =""
    End
End
