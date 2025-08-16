Operation =1
Option =0
Begin InputTables
    Name ="TimeClockPunches"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="TimeClockPunches.CashierID"
    Expression ="TimeClockPunches.PunchDate"
    Expression ="TimeClockPunches.Punch1"
    Expression ="TimeClockPunches.Punch2"
    Expression ="TimeClockPunches.Punch3"
    Expression ="TimeClockPunches.Punch4"
    Expression ="TimeClockPunches.SetupDate"
    Expression ="TimeClockPunches.ChangeDate"
    Alias ="HoursWorked"
    Expression ="(DateDiff(\"n\",[Punch1],[Punch2])/60)+(DateDiff(\"n\",[Punch3],[Punch4])/60)"
    Alias ="Wages"
    Expression ="(((DateDiff(\"n\",[Punch1],[Punch2])/60)+(DateDiff(\"n\",[Punch3],[Punch4])/60))"
        "*Cashier.HourlyWage)"
End
Begin Joins
    LeftTable ="TimeClockPunches"
    RightTable ="Cashier"
    Expression ="TimeClockPunches.CashierID=Cashier.ID"
    Flag =1
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
        dbText "Name" ="HoursWorked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wages"
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
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="TimeClockPunches"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Cashier"
        Name =""
    End
End
