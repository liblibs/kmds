Operation =3
Name ="ItemsSoldToday"
Option =0
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Name ="CashSysID"
    Expression ="InvoLine.CashSysID"
    Name ="Code"
    Expression ="InvoLine.Code"
    Alias ="SumOfQuantity"
    Name ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
End
Begin Groups
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.Code"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="SumOfQuantity"
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
    ColumnsShown =655
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
