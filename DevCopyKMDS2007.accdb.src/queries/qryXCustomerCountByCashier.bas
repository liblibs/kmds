Operation =1
Option =0
Begin InputTables
    Name ="qryXCustomerCount"
End
Begin OutputColumns
    Expression ="qryXCustomerCount.CashSysID"
    Expression ="qryXCustomerCount.CashierID"
    Alias ="COUNT"
    Expression ="Sum(1)"
    Alias ="CashierSales"
    Expression ="Sum(qryXCustomerCount.Sales)"
End
Begin Groups
    Expression ="qryXCustomerCount.CashSysID"
    GroupLevel =0
    Expression ="qryXCustomerCount.CashierID"
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
        dbText "Name" ="COUNT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierSales"
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
        Top =14
        Right =284
        Bottom =113
        Top =0
        Name ="qryXCustomerCount"
        Name =""
    End
End
