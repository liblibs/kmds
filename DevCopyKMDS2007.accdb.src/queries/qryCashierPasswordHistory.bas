Operation =1
Option =0
Where ="(((Cashier.ID)=[parCashierID]))"
Begin InputTables
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="Cashier.ID"
    Expression ="Cashier.Password"
    Expression ="Cashier.PasswordHistory1"
    Expression ="Cashier.PasswordHistory2"
    Expression ="Cashier.PasswordHistory3"
    Expression ="Cashier.PasswordHistory4"
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
        dbText "Name" ="Cashier.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.Password"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.PasswordHistory1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.PasswordHistory2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.PasswordHistory3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.PasswordHistory4"
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
    Bottom =288
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =404
        Bottom =156
        Top =0
        Name ="Cashier"
        Name =""
    End
End
