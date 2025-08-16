Operation =1
Option =0
Where ="(((CashiersAndPrograms.CashierID)=[parCashierID]) AND ((CashiersAndPrograms.Prog"
    "ramName)=[parProgramName]) AND ((CashiersAndPrograms.Allowed)=True))"
Begin InputTables
    Name ="CashiersAndPrograms"
End
Begin OutputColumns
    Expression ="CashiersAndPrograms.CashierID"
    Expression ="CashiersAndPrograms.ProgramName"
    Expression ="CashiersAndPrograms.Allowed"
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
        dbText "Name" ="CashiersAndPrograms.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashiersAndPrograms.ProgramName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashiersAndPrograms.Allowed"
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
    Bottom =330
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CashiersAndPrograms"
        Name =""
    End
End
