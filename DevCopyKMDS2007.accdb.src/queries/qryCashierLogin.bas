Operation =1
Option =0
Where ="(((Cashier.ID)=[parID]) AND ((Cashier.Password)=[parPassword]))"
Begin InputTables
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="Cashier.ID"
    Expression ="Cashier.Password"
    Expression ="Cashier.Name"
    Expression ="Cashier.AllowChangePrice"
    Expression ="Cashier.AllowChangeOnhand"
    Expression ="Cashier.AllowChangeCost"
    Expression ="Cashier.AllowChangeOrderAt"
    Expression ="Cashier.AllowChangeOrderTo"
    Expression ="Cashier.AllowChangeDescription"
    Expression ="Cashier.AllowHourlyWageView"
    Expression ="Cashier.AllowDeleteDocument"
    Expression ="Cashier.PasswordDate"
    Expression ="Cashier.AccountLocked"
    Expression ="Cashier.AccountLockedDate"
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
        dbText "Name" ="Cashier.PasswordDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.AccountLocked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.AccountLockedDate"
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
    Bottom =167
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =189
        Bottom =120
        Top =0
        Name ="Cashier"
        Name =""
    End
End
