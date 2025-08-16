Operation =1
Option =0
Begin InputTables
    Name ="TimeClockDetail"
End
Begin OutputColumns
    Alias ="PunchDate"
    Expression ="Format([PunchTime],'mm/dd/yyyy')"
    Alias ="Expr1"
    Expression ="TimeClockDetail.CashierID"
    Alias ="FirstPunch"
    Expression ="Min(TimeClockDetail.PunchTime)"
    Alias ="LastPunch"
    Expression ="Max(TimeClockDetail.PunchTime)"
End
Begin Groups
    Expression ="Format([PunchTime],'mm/dd/yyyy')"
    GroupLevel =0
    Expression ="TimeClockDetail.CashierID"
    GroupLevel =0
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
        dbText "Name" ="FirstPunch"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastPunch"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PunchDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="TimeClockDetail"
        Name =""
    End
End
