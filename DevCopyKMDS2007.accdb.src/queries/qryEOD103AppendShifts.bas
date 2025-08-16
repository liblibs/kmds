Operation =3
Name ="Shifts"
Option =0
Where ="(((Shifts.ID) Is Null))"
Begin InputTables
    Name ="InvoHead5"
    Name ="Shifts"
End
Begin OutputColumns
    Name ="ID"
    Expression ="InvoHead5.ShiftID"
    Name ="CashSysID"
    Expression ="InvoHead5.CashSysID"
    Alias ="BusinessDate"
    Name ="BusinessDate"
    Expression ="Last(Format(InvoHead5.SetupDate,\"mm/dd/yyyy\"))"
End
Begin Joins
    LeftTable ="InvoHead5"
    RightTable ="Shifts"
    Expression ="InvoHead5.CashSysID = Shifts.CashSysID"
    Flag =2
    LeftTable ="InvoHead5"
    RightTable ="Shifts"
    Expression ="InvoHead5.ShiftID = Shifts.ID"
    Flag =2
End
Begin OrderBy
    Expression ="InvoHead5.ShiftID"
    Flag =0
    Expression ="InvoHead5.CashSysID"
    Flag =0
    Expression ="Last(Format(InvoHead5.SetupDate,\"mm/dd/yyyy\"))"
    Flag =0
End
Begin Groups
    Expression ="InvoHead5.ShiftID"
    GroupLevel =0
    Expression ="InvoHead5.CashSysID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoHead5.SetupDate"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfSetupDate"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BusinessDate"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =173
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =26
        Top =9
        Right =138
        Bottom =168
        Top =0
        Name ="InvoHead5"
        Name =""
    End
    Begin
        Left =221
        Top =6
        Right =352
        Bottom =150
        Top =0
        Name ="Shifts"
        Name =""
    End
End
