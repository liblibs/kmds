Operation =2
Name ="DepositTable"
Option =0
Begin InputTables
    Name ="Item"
    Name ="Sysfile"
End
Begin OutputColumns
    Alias ="DepositTotal"
    Expression ="Sum([onhand]*[deposit])"
    Expression ="Sysfile.DepositItemID"
End
Begin Groups
    Expression ="Sysfile.DepositItemID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="DepositTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.DepositItemID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1465
    Bottom =834
    Left =-1
    Top =-1
    Right =1395
    Bottom =431
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =280
        Bottom =443
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =372
        Top =19
        Right =1015
        Bottom =442
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
