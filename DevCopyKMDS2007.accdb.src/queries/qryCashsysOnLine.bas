Operation =1
Option =0
Where ="(((CashSys.ComputerName)<>\"\") AND ((CashSys.XfrSmode)<>0))"
Begin InputTables
    Name ="CashSys"
End
Begin OutputColumns
    Expression ="CashSys.ID"
    Expression ="CashSys.ComputerName"
    Expression ="CashSys.XfrSmode"
End
Begin OrderBy
    Expression ="CashSys.ID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CashSys.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSys.ComputerName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSys.XfrSmode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =412
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CashSys"
        Name =""
    End
End
