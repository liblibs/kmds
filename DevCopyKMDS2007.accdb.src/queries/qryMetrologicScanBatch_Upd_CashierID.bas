Operation =4
Option =0
Where ="(((MetrologicScanbatch.CashierID)=0))"
Begin InputTables
    Name ="MetrologicScanbatch"
End
Begin OutputColumns
    Name ="MetrologicScanbatch.CashierID"
    Expression ="[parCashierID]"
End
Begin Parameters
    Name ="parCashierID"
    Flag =4
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="MetrologicScanbatch.txtQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.ScanCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Len([txtQuantity])"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MetrologicScanbatch.CashierID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =101
    Right =1619
    Bottom =866
    Left =-1
    Top =-1
    Right =1535
    Bottom =332
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =282
        Top =0
        Name ="MetrologicScanbatch"
        Name =""
    End
End
