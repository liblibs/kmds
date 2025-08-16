Operation =3
Name ="CashiersAndPrograms"
Option =0
Where ="(((CashiersAndPrograms.CashierID) Is Null))"
Begin InputTables
    Name ="qryCashiersAndPrograms"
    Name ="CashiersAndPrograms"
End
Begin OutputColumns
    Name ="CashierID"
    Expression ="qryCashiersAndPrograms.CashierID"
    Name ="ProgramName"
    Expression ="qryCashiersAndPrograms.ProgramName"
    Name ="Type"
    Expression ="qryCashiersAndPrograms.Type"
End
Begin Joins
    LeftTable ="qryCashiersAndPrograms"
    RightTable ="CashiersAndPrograms"
    Expression ="qryCashiersAndPrograms.CashierID = CashiersAndPrograms.CashierID"
    Flag =2
    LeftTable ="qryCashiersAndPrograms"
    RightTable ="CashiersAndPrograms"
    Expression ="qryCashiersAndPrograms.ProgramName = CashiersAndPrograms.ProgramName"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="qryCashiersAndPrograms.CashierID"
        dbInteger "ColumnWidth" ="3300"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashiersAndPrograms.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCashiersAndPrograms.ProgramName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCashiersAndPrograms.Type"
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
    Bottom =184
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =283
        Bottom =168
        Top =0
        Name ="qryCashiersAndPrograms"
        Name =""
    End
    Begin
        Left =331
        Top =12
        Right =475
        Bottom =156
        Top =0
        Name ="CashiersAndPrograms"
        Name =""
    End
End
